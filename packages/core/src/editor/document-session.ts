// One open presentation: the store, its fonts, the measurer, cached slide layouts, and image
// URLs. The editor facade and the fidelity harness both open documents through this, so the
// harness measures and paints exactly what the editor shows.

import { withoutEditorMarks } from '../automation/commands/table-edit-marks.ts';
import {
  loadDefaultFonts,
  ALL_DEFAULT_FONT_FAMILIES,
  FONT_ASSET_ROOT,
} from '@docx-editor.dev/fonts';
import {
  binaryPartOf,
  decodeEmbeddedFont,
  fontEmbeddingPermission,
  presentationInfo,
  PptxDocumentStore,
  attr,
  NS,
  readPptx,
  relTarget,
  type Element,
  type XmlPart,
  type PackageRejection,
  type PptxPackage,
  type PresentationInfo,
  imageSize,
  imageSizeEmu,
} from '../store/index.ts';
import { partNameKey } from '@eigenpal/ooxml-core';
import { resolveSlide } from '../cascade/index.ts';
import {
  HARD_MAX_FONT_BYTES,
  createFixedMeasurer,
  createHarfBuzzMeasurer,
  initializeHarfBuzz,
  isMetricCompatible,
  layoutSlide,
  type FontSource,
  type RegisteredFace,
  type SlideLayout,
  type TextMeasurer,
} from '../layout/index.ts';
import { OpBatch } from '../automation/commands/op-batch.ts';
import { createPackagedFileFetch } from './packaged-file-fetch.ts';
import { replicaBindingOf } from './replica-ids.ts';
import { sameDependencies, slideDependencies } from './slide-dependencies.ts';

/** The resolution of a tiled image whose header states none (measured on `fixtures/fills-lines.pdf`). */
const TILE_DPI = 72;

/** Options for opening a presentation. */
export interface OpenOptions {
  /** Extra faces (for example, local system fonts in the fidelity harness). Tried first. */
  readonly fonts?: readonly FontSource[];
  /** Load the metric-compatible substitutes from `@docx-editor.dev/fonts`. Default true. */
  readonly substitutes?: boolean;
  /** Use the deterministic fixed measurer (tests only). */
  readonly fixedMeasurer?: boolean;
}

/** A family the deck names that measures with another face. @public */
export interface FontSubstitution {
  readonly family: string;
  readonly substitute: string;
  /** True when the substitute keeps the original's advance widths, so line breaks match. */
  readonly metricCompatible: boolean;
}

/** Why a file could not be opened. */
export type OpenError = PackageRejection | { readonly reason: 'empty' };

/** Summary of the fonts a session uses, for the settings panel and the report. */
export interface FontReport {
  readonly embeddedFaces: number;
  readonly refusedEmbedded: readonly string[];
  /**
   * Packaged substitute faces that failed to load, each as `family (file): reason`. Text in
   * those families measures and paints with another face, or with none.
   */
  readonly substituteFailures: readonly string[];
  readonly faces: readonly { readonly family: string; readonly origin: string }[];
}

/** Elements whose `typeface` names a family layout measures with. */
const TYPEFACE_ELEMENTS = new Set(['latin', 'sym', 'buFont']);
const familyCache = new WeakMap<XmlPart, ReadonlySet<string>>();

function partFamilies(part: XmlPart): ReadonlySet<string> {
  const hit = familyCache.get(part);
  if (hit) return hit;
  const out = new Set<string>();
  const walk = (el: Element, depth: number): void => {
    if (depth > 256) return;
    if (el.ns === NS.a && TYPEFACE_ELEMENTS.has(el.local)) {
      const face = attr(el, 'typeface');
      if (face && !face.startsWith('+') && face.length <= 64) out.add(face);
    }
    for (const c of el.children) if (c.type === 'element') walk(c, depth + 1);
  };
  walk(part.root, 0);
  familyCache.set(part, out);
  return out;
}

const FAMILY_PARTS = /^\/ppt\/(slides|slideLayouts|slideMasters)\//;

/**
 * Every family the deck names for measured text: run fonts, symbol fonts, bullet fonts, and
 * the theme's font scheme. Cached per part identity, so an edit rescans only the parts it
 * changed.
 */
export function namedFamilies(pkg: PptxPackage): string[] {
  const byKey = new Map<string, string>();
  // Only the themes slide masters use: a notes or handout master theme never reaches a slide.
  const themes = new Set(presentationInfo(pkg).slides.map((s) => s.themePart));
  for (const part of pkg.parts.values()) {
    if (part.kind !== 'xml' || !(FAMILY_PARTS.test(part.partName) || themes.has(part.partName)))
      continue;
    for (const f of partFamilies(part.xml))
      if (!byKey.has(f.toLowerCase())) byKey.set(f.toLowerCase(), f);
  }
  return [...byKey.values()].sort((a, b) => a.localeCompare(b));
}

const MIME_BY_EXT: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  jpe: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  webp: 'image/webp',
};

/** The image MIME type a part name maps to, or undefined when the browser cannot draw it. */
export function imageMime(partName: string): string | undefined {
  const ext = partName.slice(partName.lastIndexOf('.') + 1).toLowerCase();
  return MIME_BY_EXT[ext];
}

type DefaultFamily = (typeof ALL_DEFAULT_FONT_FAMILIES)[number];

/** Which packaged substitute family covers a typeface the deck names. */
function substituteFamilyFor(typeface: string): DefaultFamily {
  const t = typeface.toLowerCase();
  if (t.startsWith('calibri') || t.startsWith('aptos')) return 'Calibri';
  if (t.startsWith('cambria')) return 'Cambria';
  if (t.startsWith('times') || t.includes('serif') || t.startsWith('georgia'))
    return 'Times New Roman';
  if (t.startsWith('courier') || t.includes('mono') || t.startsWith('consolas'))
    return 'Courier New';
  if (t.startsWith('century gothic')) return 'Century Gothic';
  return 'Arial';
}

/** The substitute families a deck needs: every typeface it names, plus the Arial fallback. */
export function substituteFamiliesFor(pkg: PptxPackage): DefaultFamily[] {
  const families = new Set<DefaultFamily>(['Arial']);
  for (const f of namedFamilies(pkg)) families.add(substituteFamilyFor(f));
  return [...families];
}

/** The packaged substitutes of one family, and the faces that failed to load. */
interface SubstituteLoad {
  readonly sources: readonly FontSource[];
  readonly failures: readonly string[];
}

const substituteCache = new Map<DefaultFamily, Promise<SubstituteLoad>>();

/**
 * How the substitutes are fetched. Under Node the packaged faces are `file:` URLs, which
 * Node's `fetch` rejects, so a bounded reader confined to the package's `assets/` folder
 * reads them. Browsers and Bun keep the global `fetch` (undefined here).
 */
let substituteFetcher: typeof fetch | undefined | null = null;
function substituteFetch(): typeof fetch | undefined {
  if (substituteFetcher === null)
    substituteFetcher = createPackagedFileFetch({
      trustedRoot: FONT_ASSET_ROOT,
      maxBytes: HARD_MAX_FONT_BYTES,
    });
  return substituteFetcher;
}

function loadSubstitute(family: DefaultFamily): Promise<SubstituteLoad> {
  const fetcher = substituteFetch();
  return loadDefaultFonts({ families: [family], ...(fetcher ? { fetcher } : {}) }).then(
    (fragment) => ({
      sources: fragment.sources.map((s) => ({
        family: s.request.family,
        bold: s.request.weight >= 600,
        italic: s.request.style === 'italic',
        bytes: s.bytes,
        origin: 'substitute',
      })),
      failures: fragment.failures.map((f) => `${f.family} (${f.file}): ${f.diagnostic}`),
    }),
    (error: unknown) => ({
      sources: [],
      failures: [`${family}: ${error instanceof Error ? error.message : String(error)}`],
    })
  );
}

/**
 * The packaged substitutes for some families, each family loaded once per page. A family
 * with a failed face loads again at the next open, so a passing fault does not stick.
 */
function substituteFonts(families: readonly DefaultFamily[]): Promise<SubstituteLoad> {
  return Promise.all(
    families.map((family) => {
      let hit = substituteCache.get(family);
      if (!hit) {
        hit = loadSubstitute(family);
        substituteCache.set(family, hit);
        void hit.then((load) => {
          if (load.failures.length > 0 && substituteCache.get(family) === hit)
            substituteCache.delete(family);
        });
      }
      return hit;
    })
  ).then((loads) => ({
    sources: loads.flatMap((l) => l.sources),
    failures: loads.flatMap((l) => l.failures),
  }));
}

/** Decode the presentation's embedded fonts in memory, honoring embedding permissions. */
export function embeddedFontSources(pkg: PptxPackage): {
  sources: FontSource[];
  refused: string[];
} {
  const info = presentationInfo(pkg);
  const sources: FontSource[] = [];
  const refused: string[] = [];
  for (const font of info.embeddedFonts) {
    for (const face of font.faces) {
      const bytes = binaryPartOf(pkg, face.partName);
      if (!bytes) continue;
      const decoded = decodeEmbeddedFont(bytes);
      if (!decoded.ok) {
        refused.push(`${font.typeface} (${face.style}): ${decoded.reason}`);
        continue;
      }
      const permission = fontEmbeddingPermission(decoded.info.fsType);
      if (permission === 'restricted') {
        refused.push(`${font.typeface} (${face.style}): restricted license`);
        continue;
      }
      // Sessions support editing, including after a viewer switches modes. Preview/print
      // embedding does not permit editing with this face. Keep its package part unchanged,
      // but measure and paint with a host face or a substitute instead.
      if (permission === 'preview-print') {
        refused.push(`${font.typeface} (${face.style}): preview and print only`);
        continue;
      }
      sources.push({
        family: font.typeface,
        bold: face.style === 'bold' || face.style === 'boldItalic',
        italic: face.style === 'italic' || face.style === 'boldItalic',
        bytes: decoded.sfnt,
        origin: 'embedded',
      });
    }
  }
  return { sources, refused };
}

/**
 * Host faces from the `fonts` option and its shims, after the embedded faces: the faces under
 * their own family, and host substitutions (`FontSourceSubstitution`) registered under the
 * requested family with the target family they draw.
 */
export interface HostFaces {
  readonly faces: readonly FontSource[];
  readonly substitutions: readonly { readonly face: FontSource; readonly target: string }[];
}

export const NO_HOST_FACES: HostFaces = Object.freeze({ faces: [], substitutions: [] });

/** The `origin` of a host substitution's face, for the report. */
export const HOST_SUBSTITUTE_ORIGIN = 'host-substitute';

/** The face sources a session measures with, in priority order. */
interface FontSet {
  readonly host: readonly FontSource[];
  readonly embedded: readonly FontSource[];
  readonly configured: HostFaces;
  readonly substitutes: readonly FontSource[];
  readonly refused: readonly string[];
  readonly substituteFailures: readonly string[];
  readonly fixed: boolean;
}

function buildMeasurer(set: FontSet): { measurer: TextMeasurer; faces: readonly RegisteredFace[] } {
  if (set.fixed) return { measurer: createFixedMeasurer(), faces: [] };
  const measurer = createHarfBuzzMeasurer([
    ...set.host,
    ...set.embedded,
    ...set.configured.faces,
    ...set.configured.substitutions.map((s) => s.face),
    ...set.substitutes,
  ]);
  return { measurer, faces: measurer.faces };
}

/** An open presentation. */
let sessions = 0;

/** A random id for this session, so a clipboard payload knows where it came from. */
function sessionId(): string {
  sessions += 1;
  const random =
    typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function'
      ? crypto.randomUUID()
      : Math.random().toString(36).slice(2);
  return `${random}-${sessions}`;
}

export class DocumentSession {
  readonly store: PptxDocumentStore;
  /** Identifies this open document (clipboard payloads carry it). */
  readonly id = sessionId();
  #fonts: FontSet;
  #measurer: TextMeasurer;
  #faces: readonly RegisteredFace[];
  #generation = 0;
  /** Layouts by package, then index: repeated reads of one revision. */
  #layouts = new WeakMap<PptxPackage, Map<number, SlideLayout>>();
  /** The last layout of each slide part, with what it was made from (`slideDependencies`). */
  #bySlide = new Map<string, { readonly deps: readonly unknown[]; readonly layout: SlideLayout }>();
  #substitutions: { key: string; list: FontSubstitution[] } | null = null;
  #urls = new Map<string, string>();
  #usedPartNames = new Set<string>();

  constructor(store: PptxDocumentStore, fonts: FontSet) {
    this.store = store;
    this.#fonts = fonts;
    const built = buildMeasurer(fonts);
    this.#measurer = built.measurer;
    this.#faces = built.faces;
    for (const name of store.package.parts.keys()) this.#usedPartNames.add(partNameKey(name));
    // Changed parts include added and removed ones: a removed name stays taken.
    store.subscribe((change) => {
      for (const name of change.parts) this.#usedPartNames.add(partNameKey(name));
      // A slide part that is gone (deleted, or undone) keeps no layout.
      for (const name of change.parts)
        if (this.#bySlide.has(name) && !store.package.parts.has(name)) this.#bySlide.delete(name);
    });
  }

  /**
   * Every part name this session has held, case-folded the OPC way. New parts never take one,
   * so a part name in the history, the clipboard, or a relationship always means one part.
   */
  get usedPartNames(): ReadonlySet<string> {
    return this.#usedPartNames;
  }

  /**
   * A batch over the store's package, or over `base` (a preview the store has not committed),
   * that keeps new part names unique in this session.
   */
  batch(base: PptxPackage = this.store.package): OpBatch {
    return this.#batch(base, false);
  }

  /**
   * A batch for a probe (the `can` check): it mints on copies of the record, so it uses up no
   * number. @internal
   */
  probeBatch(base: PptxPackage = this.store.package): OpBatch {
    return this.#batch(base, true);
  }

  #batch(base: PptxPackage, probe: boolean): OpBatch {
    // While a collaboration session is attached (`bindReplicaIds`), new ids are minted in the
    // replica's stripe, from the record of its collaboration session.
    const replica = replicaBindingOf(this.store);
    return new OpBatch(base, {
      reservedPartNames: this.#usedPartNames,
      ...(replica ? { ...replica, probe } : {}),
    });
  }

  /** Layout of slide `index` over a package that is not committed (a preview); not cached. */
  layoutOf(pkg: PptxPackage, index: number): SlideLayout | undefined {
    const resolved = resolveSlide(pkg, index);
    return resolved ? layoutSlide(resolved, this.measurer) : undefined;
  }

  /** The measurer in force. Replaced when the host faces change. */
  get measurer(): TextMeasurer {
    return this.#measurer;
  }

  /** Registered faces, for paint registration. Replaced when the host faces change. */
  get faces(): readonly RegisteredFace[] {
    return this.#faces;
  }

  /** Bumped every time the face set changes. */
  get fontGeneration(): number {
    return this.#generation;
  }

  get fontReport(): FontReport {
    return {
      embeddedFaces: this.#fonts.embedded.length,
      refusedEmbedded: this.#fonts.refused,
      substituteFailures: this.#fonts.substituteFailures,
      faces: this.#faces.map((f) => ({ family: f.family, origin: f.origin })),
    };
  }

  /** The host faces in force (`setHostFaces`). @internal */
  get hostFaces(): HostFaces {
    return this.#fonts.configured;
  }

  /**
   * Replace the host faces and re-measure. They rank after the embedded faces and before the
   * packaged substitutes. The document, the selection, and the undo history are not touched:
   * fonts change measurement, not content. The editor's `setFonts` calls it. @internal
   */
  setHostFaces(next: HostFaces): void {
    const now = this.#fonts.configured;
    if (now.faces.length + now.substitutions.length === 0 && next === NO_HOST_FACES) return;
    this.#fonts = { ...this.#fonts, configured: next };
    this.#remeasure();
  }

  /** Remove every host face. @internal */
  clearHostFaces(): void {
    this.setHostFaces(NO_HOST_FACES);
  }

  #remeasure(): void {
    const built = buildMeasurer(this.#fonts);
    this.#measurer = built.measurer;
    this.#faces = built.faces;
    this.#layouts = new WeakMap();
    this.#bySlide.clear();
    this.#substitutions = null;
    this.#generation += 1;
  }

  /**
   * Families the deck names that measure with a different family. Fixed-measurer sessions
   * report none (they model no fonts).
   */
  substitutedFamilies(): FontSubstitution[] {
    if (this.#fonts.fixed) return [];
    // Cached by the families the deck names and the face set: most edits change neither.
    const families = namedFamilies(this.store.package);
    const key = `${this.#generation}\u0000${families.join('\u0000')}`;
    if (this.#substitutions?.key === key) return [...this.#substitutions.list];
    const out: FontSubstitution[] = [];
    // A host substitution draws its target face under the requested family's name.
    const targets = new Map<string, string>();
    const registered = this.#fonts.configured.substitutions;
    if (registered.length > 0)
      for (const f of this.#faces) {
        const hit = registered.find((s) => s.face.bytes === f.bytes && s.face.family === f.family);
        if (hit && f.origin === HOST_SUBSTITUTE_ORIGIN) targets.set(f.id, hit.target);
      }
    for (const family of families) {
      const face = this.#measurer.resolve({ family, bold: false, italic: false });
      const target = targets.get(face.id);
      if (target && target.toLowerCase() !== family.toLowerCase())
        out.push({
          family,
          substitute: target,
          metricCompatible: isMetricCompatible(family, target),
        });
      else if (face.substituted && face.family.toLowerCase() !== family.toLowerCase())
        out.push({
          family,
          substitute: face.family,
          metricCompatible: isMetricCompatible(family, face.family),
        });
    }
    this.#substitutions = { key, list: out };
    return [...out];
  }

  get info(): PresentationInfo {
    return presentationInfo(this.store.package);
  }

  /**
   * Layout of slide `index` at the current revision. A slide whose part, layout, master, theme,
   * and the other parts it reads are unchanged keeps its layout object across edits elsewhere
   * (`slideDependencies`), so its thumbnail and its paint are reused too.
   */
  layout(index: number): SlideLayout | undefined {
    const pkg = this.store.package;
    let byIndex = this.#layouts.get(pkg);
    if (!byIndex) {
      byIndex = new Map();
      this.#layouts.set(pkg, byIndex);
    }
    const hit = byIndex.get(index);
    if (hit) return hit;
    const deps = slideDependencies(pkg, index);
    const partName = presentationInfo(pkg).slides[index]?.partName;
    if (!deps || !partName) return undefined;
    const last = this.#bySlide.get(partName);
    if (last && sameDependencies(last.deps, deps)) {
      byIndex.set(index, last.layout);
      return last.layout;
    }
    const resolved = resolveSlide(pkg, index);
    if (!resolved) return undefined;
    const laid = layoutSlide(resolved, this.measurer);
    byIndex.set(index, laid);
    this.#bySlide.set(partName, { deps, layout: laid });
    return laid;
  }

  /**
   * A URL for an image relationship. Internal parts only: an external target is never fetched.
   * Blob URLs are created lazily and revoked by `dispose()`.
   */
  imageUrl = (sourcePart: string, rId: string): string | undefined => {
    const pkg = this.store.package;
    const target = relTarget(pkg, sourcePart, rId);
    if (!target) return undefined;
    const mime = imageMime(target);
    const bytes = binaryPartOf(pkg, target);
    if (!mime || !bytes) return undefined;
    const key = `${target}\u0000${bytes.length}`;
    const hit = this.#urls.get(key);
    if (hit) return hit;
    if (typeof URL === 'undefined' || typeof Blob === 'undefined') return undefined;
    const url = URL.createObjectURL(new Blob([bytes.slice().buffer], { type: mime }));
    this.#urls.set(key, url);
    return url;
  };

  /**
   * The natural size of an image relationship in EMU, from its header and resolution, at 72 DPI
   * when it states none: the size of a tiled picture. The stored pixels count, not a JPEG's
   * EXIF orientation. Paint uses it for
   * tiled pictures, and Crop > Fill and Fit for the aspect ratio. Undefined for an external
   * target or an unreadable header.
   */
  imageSize = (sourcePart: string, rId: string): { cx: number; cy: number } | undefined => {
    const pkg = this.store.package;
    const target = relTarget(pkg, sourcePart, rId);
    const bytes = target ? binaryPartOf(pkg, target) : undefined;
    const size = bytes ? imageSize(bytes) : null;
    return size
      ? imageSizeEmu({ ...size, dpiX: size.dpiX ?? TILE_DPI, dpiY: size.dpiY ?? TILE_DPI })
      : undefined;
  };

  /** Release blob URLs. */
  dispose(): void {
    for (const url of this.#urls.values()) URL.revokeObjectURL(url);
    this.#urls.clear();
  }
}

/** Open `.pptx` bytes. Never throws for a malformed file: it returns the typed reason. */
export async function openPresentation(
  bytes: Uint8Array,
  options: OpenOptions = {}
): Promise<{ ok: true; session: DocumentSession } | { ok: false; error: OpenError }> {
  if (bytes.length === 0) return { ok: false, error: { reason: 'empty' } };
  const read = readPptx(bytes);
  if (!read.ok) return { ok: false, error: read.error };
  const { sources: embedded, refused } = embeddedFontSources(read.pkg);
  const store = new PptxDocumentStore(withoutEditorMarks(read.pkg));
  const host = options.fonts ?? [];
  const base = {
    host,
    embedded,
    refused,
    configured: NO_HOST_FACES,
    substitutes: [] as readonly FontSource[],
    substituteFailures: [] as readonly string[],
  };
  if (options.fixedMeasurer)
    return { ok: true, session: new DocumentSession(store, { ...base, fixed: true }) };
  await initializeHarfBuzz();
  const loaded =
    options.substitutes === false
      ? { sources: [], failures: [] }
      : await substituteFonts(substituteFamiliesFor(read.pkg));
  const fonts: FontSet = {
    ...base,
    substitutes: loaded.sources,
    substituteFailures: loaded.failures,
    fixed: false,
  };
  return { ok: true, session: new DocumentSession(store, fonts) };
}
