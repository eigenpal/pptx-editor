// The HarfBuzz measurer: exact advances from the real font bytes, the same approach as
// docx-editor/packages/core/src/layout/harfbuzz-shaper.ts and shaped-measurer.ts, reduced to
// the advances and line metrics that slide text layout needs.
//
// The WASM module loads through a dynamic import so importing layout never adds top-level
// await to a host's import graph. Call `initializeHarfBuzz()` once before measuring.

import { AdvanceCache } from './advance-cache.ts';
import { os2LineExtents, withFaceLineMetrics } from './face-line-metrics.ts';
import { isIdeographic } from './line-break.ts';
import {
  DEFAULT_FALLBACK_FAMILY,
  FONT_SUBSTITUTES,
  type FaceMetrics,
  type FontRequest,
  type ResolvedFace,
  type TextMeasurer,
} from './measurer.ts';
import {
  harfBuzzUnsupportedRuntimeDiagnostic,
  harfBuzzVersionMismatchDiagnostic,
  harfBuzzWasmUnavailableDiagnostic,
  isUnsupportedNodeRuntime,
} from '../harfbuzz-wasm-binary.ts';

type HarfBuzzModule = typeof import('harfbuzzjs');

/** The HarfBuzz build the engine measures with. A different runtime is refused. */
export const HARFBUZZ_VERSION = '15.0.0';

let hb: HarfBuzzModule | null = null;
let loading: Promise<void> | null = null;

/** Load and verify the HarfBuzz runtime. Idempotent. */
export function initializeHarfBuzz(): Promise<void> {
  if (hb) return Promise.resolve();
  loading ??= import('harfbuzzjs').then(
    (module) => {
      const version = module.versionString();
      if (version !== HARFBUZZ_VERSION)
        throw new Error(harfBuzzVersionMismatchDiagnostic(HARFBUZZ_VERSION, version));
      hb = module;
    },
    (cause: unknown) => {
      // A failed load can be retried after the host fixes the binary location.
      loading = null;
      throw new Error(
        isUnsupportedNodeRuntime(cause)
          ? harfBuzzUnsupportedRuntimeDiagnostic(cause)
          : harfBuzzWasmUnavailableDiagnostic(cause),
        { cause }
      );
    }
  );
  return loading;
}

/** Whether `initializeHarfBuzz()` has completed. */
export function isHarfBuzzReady(): boolean {
  return hb !== null;
}

const featureSets = new Map<boolean, InstanceType<HarfBuzzModule['Feature']>[]>();
/**
 * Slide text takes no standard or contextual ligatures, and kerning only when the run asks
 * for it (`a:rPr/@kern` at or below the run size).
 */
function shapingFeatures(
  module: HarfBuzzModule,
  kerning: boolean
): InstanceType<HarfBuzzModule['Feature']>[] {
  let set = featureSets.get(kerning);
  if (!set) {
    set = ['liga', 'clig', 'dlig', 'hlig'].map((tag) => new module.Feature(tag, 0));
    if (!kerning) set.push(new module.Feature('kern', 0));
    featureSets.set(kerning, set);
  }
  return set;
}

/** One font face the host makes available: family name, style, and sfnt bytes. */
export interface FontSource {
  readonly family: string;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly bytes: Uint8Array;
  /** Where the bytes came from, for the report: `embedded`, `substitute`, `system`. */
  readonly origin: string;
}

/** A registered face: the source plus the internal id paint uses. */
export interface RegisteredFace extends FontSource {
  readonly id: string;
}

interface LoadedFace {
  readonly entry: RegisteredFace;
  readonly font: InstanceType<HarfBuzzModule['Font']>;
  readonly upem: number;
  readonly metrics: FaceMetrics;
}

const MAX_TEXT = 100_000;
/** Advance arrays kept per measurer; the least recently used go first (`AdvanceCache`). */
const CACHE_LIMIT = 20_000;

function readOs2Underline(
  face: InstanceType<HarfBuzzModule['Face']>,
  upem: number
): [number, number] {
  const post = face.referenceTable('post');
  if (post && post.length >= 12) {
    const view = new DataView(post.buffer, post.byteOffset, post.byteLength);
    return [view.getInt16(8) / upem, Math.max(1, view.getInt16(10)) / upem];
  }
  return [-0.1, 0.05];
}

/** How many measurers this module made: each one names its faces with its own number. */
let measurers = 0;

/**
 * Create a measurer over a set of faces. Registration order matters only for ties: the first
 * face registered for a family/style wins.
 *
 * Face ids are unique for each measurer (`pptx-f<measurer>-<face>`), so the faces of two
 * documents never share a name, for example in a rasterizer's font registry.
 */
export function createHarfBuzzMeasurer(sources: readonly FontSource[]): TextMeasurer & {
  readonly faces: readonly RegisteredFace[];
} {
  if (!hb) throw new Error('Call initializeHarfBuzz() before creating a HarfBuzz measurer');
  const module = hb;
  const prefix = `pptx-f${(measurers += 1)}-`;
  const faces: RegisteredFace[] = sources.map((s, i) => ({ ...s, id: `${prefix}${i}` }));
  const loaded = new Map<string, LoadedFace>();
  const resolved = new Map<string, ResolvedFace>();
  const cache = new AdvanceCache<Float64Array>(CACHE_LIMIT);
  // One shaping buffer, reset before each use. Shaping is synchronous, so it is never shared.
  let buffer: InstanceType<HarfBuzzModule['Buffer']> | null = null;

  const load = (entry: RegisteredFace): LoadedFace | null => {
    const hit = loaded.get(entry.id);
    if (hit) return hit;
    try {
      const face = new module.Face(new module.Blob(entry.bytes), 0);
      const font = new module.Font(face);
      const upem = face.upem || 1000;
      const ext = font.hExtents();
      const line = os2LineExtents(face.referenceTable('OS/2'));
      const [underlinePosition, underlineThickness] = readOs2Underline(face, upem);
      const out: LoadedFace = {
        entry,
        font,
        upem,
        metrics: {
          ascent: (line ? line[0] : Math.abs(ext.ascender)) / upem,
          descent: (line ? line[1] : Math.abs(ext.descender)) / upem,
          lineGap: ext.lineGap / upem,
          underlinePosition,
          underlineThickness,
        },
      };
      loaded.set(entry.id, out);
      return out;
    } catch {
      return null;
    }
  };

  const byFamily = (family: string): RegisteredFace[] =>
    faces.filter((f) => f.family.toLowerCase() === family.toLowerCase());

  // Style before weight, as CSS Fonts Level 4 §5.2 matches faces: a deck that embeds bold,
  // italic, and bold italic but no regular draws regular text upright (with the bold face),
  // never slanted (measured on reference PDFs).
  const pick = (candidates: RegisteredFace[], r: FontRequest): RegisteredFace | undefined =>
    candidates.find((f) => f.bold === r.bold && f.italic === r.italic) ??
    candidates.find((f) => f.italic === r.italic) ??
    candidates.find((f) => !f.bold && !f.italic) ??
    candidates[0];

  const resolve = (r: FontRequest): ResolvedFace => {
    const key = `${r.family.toLowerCase()}|${r.bold}|${r.italic}`;
    const hit = resolved.get(key);
    if (hit) return hit;
    const order = [
      r.family,
      ...(FONT_SUBSTITUTES[r.family.toLowerCase()] ?? []),
      DEFAULT_FALLBACK_FAMILY,
    ];
    let chosen: LoadedFace | null = null;
    let substituted = false;
    for (let i = 0; i < order.length && !chosen; i += 1) {
      const entry = pick(byFamily(order[i]), r);
      if (entry) {
        chosen = load(entry);
        substituted = i > 0;
      }
    }
    if (!chosen && faces[0]) {
      chosen = load(faces[0]);
      substituted = true;
    }
    const face: ResolvedFace = chosen
      ? {
          id: chosen.entry.id,
          family: chosen.entry.family,
          bold: chosen.entry.bold,
          italic: chosen.entry.italic,
          synthBold: r.bold && !chosen.entry.bold,
          synthItalic: r.italic && !chosen.entry.italic,
          substituted,
          // A missing face keeps its own line metrics when the table knows them.
          metrics: substituted ? withFaceLineMetrics(r.family, chosen.metrics) : chosen.metrics,
        }
      : {
          id: r.family,
          family: r.family,
          bold: r.bold,
          italic: r.italic,
          synthBold: false,
          synthItalic: false,
          substituted: true,
          metrics: {
            ascent: 0.9,
            descent: 0.25,
            lineGap: 0,
            underlinePosition: -0.1,
            underlineThickness: 0.05,
          },
        };
    resolved.set(key, face);
    return face;
  };

  const advances = (
    face: ResolvedFace,
    text: string,
    options?: { readonly kerning?: boolean }
  ): Float64Array => {
    const kerning = options?.kerning === true;
    const key = `${face.id}${kerning ? 'k' : ''}\u0000${text}`;
    const hit = cache.get(key);
    if (hit) return hit;
    const out = new Float64Array(text.length);
    const lf = loaded.get(face.id);
    if (!lf || text.length === 0 || text.length > MAX_TEXT) {
      for (let i = 0; i < text.length; i += 1) out[i] = 0.5;
      return out;
    }
    buffer ??= new module.Buffer();
    buffer.reset();
    buffer.addText(text);
    buffer.guessSegmentProperties();
    // Slide text takes no standard or contextual ligatures, and embedded subset fonts
    // often lack the ligature glyphs. Paint disables them too.
    module.shape(lf.font, buffer, shapingFeatures(module, kerning));
    // `addText` adds UTF-16, so clusters are UTF-16 offsets into `text`.
    const infos = buffer.getGlyphInfos();
    const positions = buffer.getGlyphPositions();
    for (let g = 0; g < infos.length; g += 1) {
      const cluster = infos[g].cluster;
      if (cluster < 0 || cluster >= out.length) continue;
      // Glyph 0 is `.notdef`: the face lacks the character, and paint falls back to a system
      // face. Measure the fallback, not the empty box: East Asian characters are full width
      // (1 em), others half width. An approximation, listed on the public fidelity page.
      if (infos[g].codepoint === 0)
        out[cluster] += isIdeographic(text.codePointAt(cluster) ?? 0) ? 1 : 0.5;
      else out[cluster] += positions[g].xAdvance / lf.upem;
    }
    cache.set(key, out);
    return out;
  };

  return { resolve, advances, faces };
}
