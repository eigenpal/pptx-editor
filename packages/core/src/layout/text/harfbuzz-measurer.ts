// The HarfBuzz measurer: exact advances from the real font bytes, the same approach as
// docx-editor/packages/core/src/layout/harfbuzz-shaper.ts and
// packages/core/src/layout/shaped-measurer.ts, reduced to
// the advances and line metrics that slide text layout needs. DOCX Editor commit adf0e178.
//
// The WASM module loads through a dynamic import so importing layout never adds top-level
// await to a host's import graph. Call `initializeHarfBuzz()` once before measuring.

import type { PathCommand } from '@eigenpal/ooxml-core';
import { AdvanceCache } from './advance-cache.ts';
import { os2LineExtents, withFaceLineMetrics } from './face-line-metrics.ts';
import { isIdeographic } from './line-break.ts';
import { unkernedTextRanges } from './run-kerning.ts';
import {
  cmapMapsCodePoint,
  unicodeCmapSubtable,
  type UnicodeSubtable,
} from './sfnt-cmap-coverage.ts';
import { sfntStrikeoutStrokeEm } from './sfnt-strikeout-metrics.ts';
import {
  DEFAULT_FALLBACK_FAMILY,
  eastAsianFallbackFamilies,
  fontSubstitutes,
  type FaceMetrics,
  type FontRequest,
  type ResolvedFace,
  type TextMeasurer,
  type ShapedGlyph,
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
 * for it (`a:rPr/@kern` at or below the run size) and the face can kern
 * (`LoadedFace.kerns`).
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

/**
 * The features for one shaping call: `shapingFeatures`, and with kerning on, `kern` off over
 * full-width East Asian text (`unkernedTextRanges`). Clusters are UTF-16 offsets, as the ranges.
 */
function featuresFor(
  module: HarfBuzzModule,
  kerning: boolean,
  text: string
): InstanceType<HarfBuzzModule['Feature']>[] {
  const set = shapingFeatures(module, kerning);
  if (!kerning) return set;
  const ranges = unkernedTextRanges(text);
  if (ranges.length === 0) return set;
  return [...set, ...ranges.map(([start, end]) => new module.Feature('kern', 0, start, end))];
}

/**
 * Characters that HarfBuzz may draw with a glyph the face's cmap does not give them, and that
 * still draw right: spaces take the face's space (Unicode spaces fallback), and default
 * ignorable characters draw nothing.
 */
const SHAPER_DRAWS = /^[\p{Zs}\p{Default_Ignorable_Code_Point}]$/u;

/** One font face the host makes available: family name, style, and sfnt bytes. */
export interface FontSource {
  readonly family: string;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly bytes: Uint8Array;
  /**
   * Where the bytes came from, for the report: `embedded`, `substitute`, `system`, or
   * `HOST_SUBSTITUTE_ORIGIN`.
   */
  readonly origin: string;
}

/**
 * The `origin` of a host substitution's face: the face of another family, registered under the
 * family that the deck names.
 */
export const HOST_SUBSTITUTE_ORIGIN = 'host-substitute';

/** A registered face: the source plus the internal id paint uses. */
export interface RegisteredFace extends FontSource {
  readonly id: string;
}

interface LoadedFace {
  readonly entry: RegisteredFace;
  readonly font: InstanceType<HarfBuzzModule['Font']>;
  readonly upem: number;
  readonly metrics: FaceMetrics;
  /**
   * Whether the face kerns when a run asks for it. A face that the deck names kerns only when it
   * has a `kern` table: the reference PDFs draw a face whose pairs are only in `GPOS` (Carlito)
   * without kerning, also where the run asks for it (`font-effects`, `diagrams`, and every
   * fixture through the default text style), and kern Liberation Sans, which has both tables
   * (`rtl-text`, `east-asian-line-breaking`). A stand-in for a missing face keeps its `GPOS`
   * pairs, as the face it replaces may have a `kern` table: a face that
   * the measurer substitutes, and a host substitution (`HOST_SUBSTITUTE_ORIGIN`).
   * Unit tests cover this substitution rule; no reference PDF measures it.
   */
  readonly kerns: boolean;
  /**
   * The face's Unicode cmap subtable (`unicodeCmapSubtable`), or null for a face without one,
   * such as a symbol face. Coverage then comes from HarfBuzz alone.
   */
  readonly cmap: UnicodeSubtable | null;
}

const MAX_TEXT = 100_000;
/** Advance arrays kept per measurer; the least recently used go first (`AdvanceCache`). */
const CACHE_LIMIT = 20_000;
/**
 * The characters of the advance arrays kept per measurer: about 10 MB with their keys. Each
 * keystroke in a paragraph measures its run again, so a count limit alone would keep
 * `CACHE_LIMIT` copies of a long paragraph.
 */
const CACHE_CHARS = 1_000_000;
/**
 * Outline commands for one glyph: a real glyph has a few hundred at most. A glyph with more has
 * no outline, so a text warp skips it.
 */
const MAX_OUTLINE_COMMANDS = 2_000;
/** Outline commands kept in the outline cache, over all glyphs. */
const MAX_CACHED_OUTLINE_COMMANDS = 1_000_000;

/**
 * An sfnt table of face 0, read from the font bytes in JavaScript. `Face.referenceTable` makes
 * a HarfBuzz blob that is never freed and holds a copy of the whole font, so each new measurer
 * leaked its fonts into the WebAssembly heap until it was full. @internal
 */
export function sfntTable(bytes: Uint8Array, tag: string): Uint8Array | undefined {
  if (bytes.length < 12) return undefined;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  // A collection ('ttcf') lists the offset of each font's table directory from byte 12.
  const ttc = view.getUint32(0) === 0x74746366;
  const base = ttc ? (bytes.length >= 16 ? view.getUint32(12) : -1) : 0;
  if (base < 0 || base + 12 > bytes.length) return undefined;
  const count = view.getUint16(base + 4);
  for (let i = 0; i < count; i += 1) {
    const rec = base + 12 + i * 16;
    if (rec + 16 > bytes.length) return undefined;
    const name = String.fromCharCode(
      bytes[rec]!,
      bytes[rec + 1]!,
      bytes[rec + 2]!,
      bytes[rec + 3]!
    );
    if (name !== tag) continue;
    const offset = view.getUint32(rec + 8);
    const length = view.getUint32(rec + 12);
    return length > 0 && offset <= bytes.length && length <= bytes.length - offset
      ? bytes.subarray(offset, offset + length)
      : undefined;
  }
  return undefined;
}

function readOs2Underline(bytes: Uint8Array, upem: number): [number, number] {
  const post = sfntTable(bytes, 'post');
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
  const cache = new AdvanceCache<Float64Array>(CACHE_LIMIT, CACHE_CHARS);
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
      const line = os2LineExtents(sfntTable(entry.bytes, 'OS/2'));
      const [underlinePosition, underlineThickness] = readOs2Underline(entry.bytes, upem);
      const strike = sfntStrikeoutStrokeEm(entry.bytes, 0);
      const out: LoadedFace = {
        entry,
        font,
        upem,
        kerns:
          entry.origin === HOST_SUBSTITUTE_ORIGIN || sfntTable(entry.bytes, 'kern') !== undefined,
        cmap: unicodeCmapSubtable(entry.bytes, 0),
        metrics: {
          ascent: (line ? line[0] : Math.abs(ext.ascender)) / upem,
          descent: (line ? line[1] : Math.abs(ext.descender)) / upem,
          lineGap: ext.lineGap / upem,
          underlinePosition,
          underlineThickness,
          ...(strike
            ? { strikeoutPosition: strike.offsetEm, strikeoutThickness: strike.thicknessEm }
            : {}),
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
    const eastAsian = r.script === 'eastAsian';
    const fallbackOrder = eastAsian ? eastAsianFallbackFamilies(r.lang) : [];
    const key = `${r.family.toLowerCase()}|${r.bold}|${r.italic}${eastAsian ? `|ea|${fallbackOrder[0]}` : ''}`;
    const hit = resolved.get(key);
    if (hit) return hit;
    const order = [
      r.family,
      ...fontSubstitutes(r.family),
      ...fallbackOrder,
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

  // The cmap answer for each BMP code point of a face: 0 not read yet, 1 mapped, 2 missing.
  // A lookup reads the font bytes (a binary search over thousands of groups in an East Asian
  // face), and the measurer asks for every non-ASCII character it shapes. 64 KiB for each face
  // that measures non-ASCII text.
  const bmpCoverage = new WeakMap<LoadedFace, Uint8Array>();
  const cmapHas = (lf: LoadedFace, cmap: UnicodeSubtable, cp: number): boolean => {
    if (cp > 0xffff) return cmapMapsCodePoint(lf.entry.bytes, cmap, cp);
    let known = bmpCoverage.get(lf);
    if (!known) bmpCoverage.set(lf, (known = new Uint8Array(0x10000)));
    if (known[cp] === 0) known[cp] = cmapMapsCodePoint(lf.entry.bytes, cmap, cp) ? 1 : 2;
    return known[cp] === 1;
  };

  /**
   * Whether the face draws a non-ASCII `cp` with its own glyph. Spaces and default ignorable
   * characters keep HarfBuzz's answer (`SHAPER_DRAWS`), and so does every character of a face
   * without a Unicode subtable. Other characters need the face's cmap.
   */
  const covers = (lf: LoadedFace, cp: number): boolean => {
    if (!lf.cmap) return true;
    return cmapHas(lf, lf.cmap, cp) || SHAPER_DRAWS.test(String.fromCodePoint(cp));
  };

  // The face's own hyphen advance in em (U+2010, else U+002D), or null when it has neither.
  const hyphens = new WeakMap<LoadedFace, number | null>();
  const hyphenAdvance = (lf: LoadedFace): number | null => {
    let known = hyphens.get(lf);
    if (known === undefined) {
      const cp = lf.cmap && !cmapHas(lf, lf.cmap, 0x2010) ? 0x2d : 0x2010;
      const glyph = lf.font.nominalGlyph(cp) || lf.font.nominalGlyph(0x2d);
      known = glyph ? lf.font.glyphHAdvance(glyph) / lf.upem : null;
      hyphens.set(lf, known);
    }
    return known;
  };

  /**
   * The estimated advance of a character the face lacks. A hyphen or dash (U+2010 to U+2013)
   * takes the face's own hyphen advance, as a browser or the PDF fallback draws a hyphen of
   * about that width: U+2011 in Liberation Sans is 0.333 em, not 0.5 em. Of the other
   * characters, East Asian ones are full width (1 em), and the rest half width.
   */
  const missingAdvance = (lf: LoadedFace, cp: number): number => {
    if (cp >= 0x2010 && cp <= 0x2013) {
      const hyphen = hyphenAdvance(lf);
      if (hyphen !== null) return hyphen;
    }
    return isIdeographic(cp) ? 1 : 0.5;
  };

  const advances = (
    face: ResolvedFace,
    text: string,
    options?: { readonly kerning?: boolean; readonly mathScript?: 1 | 2 }
  ): Float64Array => {
    const lf = loaded.get(face.id);
    const kerning = options?.kerning === true && (lf?.kerns === true || face.substituted);
    const ssty = options?.mathScript;
    const key = `${face.id}${kerning ? 'k' : ''}${ssty ? `s${ssty}` : ''}\u0000${text}`;
    const hit = cache.get(key);
    if (hit) return hit;
    const out = new Float64Array(text.length);
    if (!lf || text.length === 0 || text.length > MAX_TEXT) {
      for (let i = 0; i < text.length; i += 1) out[i] = 0.5;
      return out;
    }
    buffer ??= new module.Buffer();
    buffer.reset();
    buffer.addText(text);
    // Script-style forms are features of the math script (`Zmth`), which no character of the
    // text names: set it before the other properties are guessed.
    if (ssty) buffer.setScript('Zmth');
    buffer.guessSegmentProperties();
    // Slide text takes no standard or contextual ligatures, and embedded subset fonts
    // often lack the ligature glyphs. Paint disables them too.
    const features = featuresFor(module, kerning, text);
    module.shape(
      lf.font,
      buffer,
      ssty ? [...features, new module.Feature('ssty', ssty)] : features
    );
    // `addText` adds UTF-16, so clusters are UTF-16 offsets into `text`.
    const infos = buffer.getGlyphInfos();
    const positions = buffer.getGlyphPositions();
    let estimated = -1;
    for (let g = 0; g < infos.length; g += 1) {
      const cluster = infos[g].cluster;
      if (cluster < 0 || cluster >= out.length) continue;
      if (cluster === estimated) continue;
      // Glyph 0 is `.notdef`: the face lacks the character, and paint falls back to a system
      // face. A character the face's cmap lacks is missing too, also when HarfBuzz draws it
      // with another of the face's glyphs (U+2010 for U+2011, or a decomposition): paint
      // draws characters, not HarfBuzz's glyphs. Measure the fallback, not the face's glyph,
      // once per cluster (`missingAdvance`). An approximation, listed on the public fidelity
      // page.
      const own = infos[g].codepoint !== 0;
      if (own && (text.charCodeAt(cluster) < 0x80 || covers(lf, text.codePointAt(cluster)!)))
        out[cluster] += positions[g].xAdvance / lf.upem;
      else {
        out[cluster] += missingAdvance(lf, text.codePointAt(cluster)!);
        estimated = cluster;
      }
    }
    cache.set(key, out, text.length);
    return out;
  };

  const shape: NonNullable<TextMeasurer['shape']> = (face, text, options) => {
    const lf = loaded.get(face.id);
    if (!lf || !text.length || text.length > MAX_TEXT) return [];
    buffer ??= new module.Buffer();
    buffer.reset();
    buffer.addText(text);
    buffer.setDirection(options.rtl ? module.Direction.RTL : module.Direction.LTR);
    buffer.guessSegmentProperties();
    const kerning = options.kerning === true && (lf.kerns || face.substituted);
    module.shape(lf.font, buffer, featuresFor(module, kerning, text));
    const infos = buffer.getGlyphInfos();
    const positions = buffer.getGlyphPositions();
    return infos.map((info, i): ShapedGlyph => ({
      id: info.codepoint,
      cluster: info.cluster,
      advance: positions[i].xAdvance / lf.upem,
      offsetX: positions[i].xOffset / lf.upem,
      offsetY: positions[i].yOffset / lf.upem,
    }));
  };

  // Glyph coverage, by face id and code point; bounded like the advance cache.
  const fallbacks = new Map<string, ResolvedFace | null>();
  const has = (entry: RegisteredFace, cp: number): boolean => {
    const lf = load(entry);
    if (!lf) return false;
    if (lf.cmap) return cmapHas(lf, lf.cmap, cp);
    const glyph = lf.font.nominalGlyph(cp);
    return glyph !== undefined && glyph !== 0;
  };
  /** One regular face for each family, in registration order: the fallback candidates. */
  const families: RegisteredFace[] = [];
  for (const f of faces)
    if (!families.some((g) => g.family.toLowerCase() === f.family.toLowerCase()))
      families.push(pick(byFamily(f.family), { family: f.family, bold: false, italic: false })!);

  const fallbackFace = (
    face: ResolvedFace,
    cp: number,
    lang?: string
  ): ResolvedFace | undefined => {
    const order = eastAsianFallbackFamilies(lang);
    const key = `${face.id}|${face.bold}|${face.italic}|${cp}|${order[0]}`;
    const hit = fallbacks.get(key);
    if (hit !== undefined) return hit ?? undefined;
    const own = faces.find((f) => f.id === face.id);
    let found: ResolvedFace | null = null;
    if (own && !has(own, cp)) {
      const eastAsian = isIdeographic(cp) || (cp >= 0x2460 && cp <= 0x27bf);
      const named = eastAsian ? order.flatMap((name) => byFamily(name).slice(0, 1)) : [];
      for (const entry of [...named, ...families]) {
        if (entry.family.toLowerCase() === own.family.toLowerCase() || !has(entry, cp)) continue;
        found = resolve({ family: entry.family, bold: face.bold, italic: face.italic });
        break;
      }
    }
    if (fallbacks.size >= CACHE_LIMIT) fallbacks.clear();
    fallbacks.set(key, found);
    return found ?? undefined;
  };

  // Glyph outlines in em, y up, by face id and code point; bounded by entries and by commands.
  const outlines = new Map<string, readonly PathCommand[] | null>();
  let cachedCommands = 0;
  let draw: InstanceType<HarfBuzzModule['DrawFuncs']> | null = null;
  let sink: PathCommand[] = [];
  let overflow = false;
  let k = 1;
  const glyphOutline = (face: ResolvedFace, cp: number): readonly PathCommand[] | undefined => {
    const key = `${face.id}|${cp}`;
    const hit = outlines.get(key);
    if (hit !== undefined) return hit ?? undefined;
    const entry = faces.find((f) => f.id === face.id);
    const lf = entry ? load(entry) : null;
    const glyph = lf?.font.nominalGlyph(cp);
    let out: PathCommand[] | null = null;
    if (lf && glyph) {
      k = 1 / lf.upem;
      // Drawing is synchronous, so one set of callbacks writes into the current sink.
      if (!draw) {
        draw = new module.DrawFuncs();
        const add = (c: PathCommand): void => {
          if (sink.length < MAX_OUTLINE_COMMANDS) sink.push(c);
          else overflow = true;
        };
        draw.setMoveToFunc((x, y) => add({ op: 'M', x: x * k, y: y * k }));
        draw.setLineToFunc((x, y) => add({ op: 'L', x: x * k, y: y * k }));
        draw.setQuadraticToFunc((x1, y1, x, y) =>
          add({ op: 'Q', x1: x1 * k, y1: y1 * k, x: x * k, y: y * k })
        );
        draw.setCubicToFunc((x1, y1, x2, y2, x, y) =>
          add({ op: 'C', x1: x1 * k, y1: y1 * k, x2: x2 * k, y2: y2 * k, x: x * k, y: y * k })
        );
        draw.setClosePathFunc(() => add({ op: 'Z' }));
      }
      sink = [];
      overflow = false;
      lf.font.drawGlyph(glyph, draw);
      out = overflow ? null : sink;
      sink = [];
    }
    const commands = out?.length ?? 0;
    if (outlines.size >= CACHE_LIMIT || cachedCommands + commands > MAX_CACHED_OUTLINE_COMMANDS) {
      outlines.clear();
      cachedCommands = 0;
    }
    cachedCommands += commands;
    outlines.set(key, out);
    return out ?? undefined;
  };

  return { resolve, advances, shape, fallbackFace, glyphOutline, faces };
}
