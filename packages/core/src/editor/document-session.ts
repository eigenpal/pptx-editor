// One open presentation: the store, its fonts, the measurer, cached slide layouts, and image
// URLs. The editor facade and the fidelity harness both open documents through this, so the
// harness measures and paints exactly what the editor shows.

import { partWithoutEditorMarks } from '../automation/commands/table-edit-marks.ts';
import { openStore } from '../automation/commands/open-store.ts';
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
  readPptxDeferred,
  relTarget,
  isDeferredPending,
  parseDeferredPart,
  parseDeferredParts,
  pendingDeferredBytes,
  deferredRejection,
  PPTX_ZIP_LIMITS,
  type Element,
  type XmlPart,
  type PackageRejection,
  type PptxPackage,
  type PackagePart,
  type PresentationInfo,
  changedPartNames,
  deferredParseCount,
  imageSize,
  imageSizeEmu,
} from '../store/index.ts';
import { metafileSizeEmu, partNameKey } from '@eigenpal/ooxml-core';
import { resolveSlide } from '../cascade/index.ts';
import {
  HARD_MAX_FONT_BYTES,
  HOST_SUBSTITUTE_ORIGIN,
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
import { carryFits } from '../automation/commands/table-fit.ts';
import { refitEditedText } from './late-fit.ts';
import { MetafilePictures, type DrawablePicture } from './metafile-pictures.ts';
import type { PicturesWait, PictureStatus } from './metafile-pictures.ts';
import { within } from './within.ts';
import {
  clampSlide,
  nextTask,
  OPEN_STEP_MS,
  openingOrder,
  ProgressiveOpen,
  PROGRESSIVE_OPEN_CONTENT_BYTES,
} from './progressive-open.ts';

/** The longest the late fit waits for substitutes that slides parsed late name. */
const LATE_SUBSTITUTE_WAIT_MS = 10_000;

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
  /**
   * The longest the open waits for the text engine (HarfBuzz WebAssembly) and the packaged
   * substitutes, in milliseconds. After it, the session measures with the fixed measurer and
   * switches to shaped measurement when they arrive (`DocumentSession.shaping`), as DOCX Editor
   * opens before its fonts. Omitted, the open waits for them: a PDF export or a server session
   * needs shaped measurement from the start. @internal
   */
  readonly shapingWaitMs?: number;
  /**
   * Open a long deck progressively: only slide `startSlide` (and its notes) parses now, and
   * the session's `opening` parses the rest (`progressive-open.ts`). A deck whose other slides
   * hold less than `PROGRESSIVE_OPEN_CONTENT_BYTES` parses whole. @internal
   */
  readonly progressive?: { readonly startSlide?: number };
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
   * Packaged substitute faces that failed to load. Text in those families measures and paints
   * with another face, or with none.
   */
  readonly substituteFailures: readonly SubstituteFontFailure[];
  readonly faces: readonly { readonly family: string; readonly origin: string }[];
}

/**
 * A packaged substitute face that failed to load, as DOCX Editor's `DefaultFontLoadFailure`
 * reports it. @public
 */
export interface SubstituteFontFailure {
  /** The family the face substitutes for, for example `Calibri`. */
  readonly family: string;
  /** The packaged file, for example `Carlito-Regular.ttf`. Empty when the whole family failed. */
  readonly file: string;
  /** Why it failed. Do not use this text for program control. */
  readonly diagnostic: string;
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

/** The families of the parts that name measured text, for the last package read. */
interface FamilyCount {
  /** `deferredParseCount()` when counted: a slide that parses changes no part identity. */
  readonly parsed: number;
  readonly slides: PresentationInfo['slides'];
  readonly themes: ReadonlySet<string | undefined>;
  /** The number of parts that name each family, by its exact spelling. */
  readonly counts: ReadonlyMap<string, number>;
  readonly list: readonly string[];
}

const familyCounts = new WeakMap<PptxPackage, FamilyCount>();
/** The last package counted, held weakly: a closed deck is not kept alive for it. */
let lastFamilyPackage: WeakRef<PptxPackage> | null = null;
const EMPTY_FAMILIES: ReadonlySet<string> = new Set();

/** The families `part` adds to the count, or none. */
function countedFamilies(
  part: PackagePart | undefined,
  themes: ReadonlySet<string | undefined>
): ReadonlySet<string> {
  if (part?.kind !== 'xml' || !(FAMILY_PARTS.test(part.partName) || themes.has(part.partName)))
    return EMPTY_FAMILIES;
  // A slide still opening names its families once it parses (`progressive-open.ts`).
  if (isDeferredPending(part)) return EMPTY_FAMILIES;
  return partFamilies(part.xml);
}

/**
 * The sorted list, one entry per case-folded family, or null when two spellings of one family
 * remain: a full pass then picks the spelling of the first part that names it.
 */
function familyList(counts: ReadonlyMap<string, number>): string[] | null {
  const folded = new Set<string>();
  for (const f of counts.keys()) {
    const key = f.toLowerCase();
    if (folded.has(key)) return null;
    folded.add(key);
  }
  return [...counts.keys()].sort((a, b) => a.localeCompare(b));
}

/**
 * Every family the deck names for measured text: run fonts, symbol fonts, bullet fonts, and
 * the theme's font scheme. Cached per part identity, so an edit rescans only the parts it
 * changed. After an edit that keeps the slide list, the count of the last package changes by
 * the changed parts only, so a keystroke walks no other part. The result is the one a full
 * pass gives.
 */
export function namedFamilies(pkg: PptxPackage): string[] {
  const info = presentationInfo(pkg);
  const parsed = deferredParseCount();
  const hit = familyCounts.get(pkg);
  if (hit?.parsed === parsed) return rememberFamilies(pkg, hit);
  const lastPkg = lastFamilyPackage?.deref();
  const last = lastPkg && familyCounts.get(lastPkg);
  const changed =
    lastPkg && last && last.parsed === parsed && last.slides === info.slides
      ? changedPartNames(lastPkg, pkg, 64)
      : null;
  if (last && changed) {
    const counts = new Map(last.counts);
    const add = (families: ReadonlySet<string>, by: number): void => {
      for (const f of families) {
        const n = (counts.get(f) ?? 0) + by;
        if (n > 0) counts.set(f, n);
        else counts.delete(f);
      }
    };
    for (const name of changed) {
      add(countedFamilies(lastPkg?.parts.get(name), last.themes), -1);
      add(countedFamilies(pkg.parts.get(name), last.themes), 1);
    }
    const list = familyList(counts);
    if (list) {
      return rememberFamilies(pkg, {
        parsed,
        slides: info.slides,
        themes: last.themes,
        counts,
        list,
      });
    }
  }
  // Only the themes slide masters use: a notes or handout master theme never reaches a slide.
  const themes = new Set(info.slides.map((s) => s.themePart));
  const counts = new Map<string, number>();
  const byKey = new Map<string, string>();
  for (const part of pkg.parts.values())
    for (const f of countedFamilies(part, themes)) {
      counts.set(f, (counts.get(f) ?? 0) + 1);
      if (!byKey.has(f.toLowerCase())) byKey.set(f.toLowerCase(), f);
    }
  const list = [...byKey.values()].sort((a, b) => a.localeCompare(b));
  return rememberFamilies(pkg, { parsed, slides: info.slides, themes, counts, list });
}

function rememberFamilies(pkg: PptxPackage, count: FamilyCount): string[] {
  familyCounts.set(pkg, count);
  lastFamilyPackage = new WeakRef(pkg);
  return [...count.list];
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

/**
 * Which packaged substitute family covers a typeface the deck names. A deck can name a packaged
 * open face itself (Carlito, Caladea, TeX Gyre Adventor): it loads the family that ships it, so
 * the face draws as itself and not as the Arial fallback.
 */
function substituteFamilyFor(typeface: string): DefaultFamily {
  const t = typeface.toLowerCase();
  if (t.startsWith('calibri') || t.startsWith('aptos') || t === 'carlito') return 'Calibri';
  if (t.startsWith('cambria') || t === 'caladea') return 'Cambria';
  if (t.startsWith('tex gyre adventor')) return 'Century Gothic';
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

/** The packaged substitutes of one family, and the faces that failed to load. @internal */
export interface SubstituteLoad {
  /** The faces that loaded. */
  readonly sources: readonly FontSource[];
  /** The faces that did not load. */
  readonly failures: readonly SubstituteFontFailure[];
}

const substituteCache = new Map<DefaultFamily, Promise<SubstituteLoad>>();

/**
 * How the substitutes are fetched. Under Node the packaged faces are `file:` URLs, which
 * Node's `fetch` rejects, so a bounded reader confined to the package's `assets/` folder
 * reads them (Bun takes the same reader). Browsers keep the global `fetch` (undefined here).
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

/** What a missing file on disk means: a Node bundle that left the fonts out. */
const MISSING_ASSET =
  'the file is not in the assets folder of @docx-editor.dev/fonts. A bundled Node app must ship that folder and set DOCX_EDITOR_FONT_ASSET_ROOT to it';

/** Load one family's packaged substitutes. `fetcher` is injectable for tests. @internal */
export function loadSubstitute(
  family: DefaultFamily,
  fetcher: typeof fetch | undefined = substituteFetch()
): Promise<SubstituteLoad> {
  return loadDefaultFonts({ families: [family], ...(fetcher ? { fetcher } : {}) }).then(
    (fragment) => ({
      sources: fragment.sources.map((s) => ({
        family: s.request.family,
        bold: s.request.weight >= 600,
        italic: s.request.style === 'italic',
        bytes: s.bytes,
        origin: 'substitute',
      })),
      failures: fragment.failures.map((f) => ({
        family: f.family,
        file: f.file,
        diagnostic:
          fetcher && f.diagnostic === 'HTTP 404'
            ? `${f.diagnostic}: ${MISSING_ASSET}`
            : f.diagnostic,
      })),
    }),
    (error: unknown) => ({
      sources: [],
      failures: [
        {
          family,
          file: '',
          diagnostic: error instanceof Error ? error.message : String(error),
        },
      ],
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

export { HOST_SUBSTITUTE_ORIGIN } from '../layout/index.ts';

/** The face sources a session measures with, in priority order. */
interface FontSet {
  readonly host: readonly FontSource[];
  readonly embedded: readonly FontSource[];
  readonly configured: HostFaces;
  readonly substitutes: readonly FontSource[];
  readonly refused: readonly string[];
  readonly substituteFailures: readonly SubstituteFontFailure[];
  /** The substitute families loaded, or null when the open loads none. */
  readonly substituteFamilies: readonly DefaultFamily[] | null;
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
  #pictureGeneration = 0;
  /** EMF and WMF parts draw as PNG once converted, then `onPictures` repaints. @internal */
  readonly pictures = new MetafilePictures(
    () => this.#faces.map((f) => f.bytes),
    () => (this.#pictureGeneration++, this.onPictures?.())
  );
  /** Hears that converted pictures are ready to paint. The editor sets it. @internal */
  onPictures: (() => void) | null = null;
  #usedPartNames = new Set<string>();

  /**
   * The background parse of a progressive open, until it settles; null for a deck that parsed
   * whole. The editor runs it (`continueOpening`) and drops it once it settles. @internal
   */
  opening: ProgressiveOpen | null = null;
  /**
   * The refusal of a slide of a progressive open that did not parse, once one did. The store
   * refuses every commit from then on (`PptxDocumentStore.block`). @internal
   */
  damage: Extract<PackageRejection, { reason: 'xml' }> | null = null;
  /** Hears `damage` at once, where the read happened. The editor sets it. @internal */
  onDamage: ((rejection: Extract<PackageRejection, { reason: 'xml' }>) => void) | null = null;
  /**
   * Shaped measurement still on the way: the open stopped waiting for the text engine or the
   * packaged substitutes (`OpenOptions.shapingWaitMs`). Resolves `true` once the session
   * measures with them, `false` when they cannot load and the fixed measurer stays. Null when
   * nothing is on the way. @internal
   */
  shaping: Promise<boolean> | null = null;
  /** Why shaped measurement did not start, once `shaping` resolved `false`. @internal */
  shapingError: unknown = null;
  /**
   * When the parse ended (`performance.now()`), so a font wait after the open counts from the
   * parse and not from the bytes: a long parse does not use up the wait. @internal
   */
  parsedAt = performance.now();
  #disposed = false;

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

  /** Changes when the same slides paint differently: new faces, or converted pictures. @internal */
  get paintGeneration(): string {
    return `${this.#generation}.${this.#pictureGeneration}`;
  }

  /** Changes when slide part `slide` paints differently: new faces, or its pictures. @internal */
  paintKey(slide: string | undefined): string {
    return `${this.#generation}.${this.pictures.stateOf(this.store.package, slide)}`;
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

  /**
   * Load the packaged substitutes that slides parsed after the open need: a progressive open
   * chose them from the parts it had parsed, and a slide parsed later can name another family.
   * Re-measures only when a family was added. Resolves true when it was. @internal
   */
  addMissingSubstitutes(): Promise<boolean> {
    // One load at a time: two at once (the opening lane and late shaping) chose the same
    // families and added each face twice.
    const next = this.#substituteLoad.then(() => this.#loadMissingSubstitutes());
    this.#substituteLoad = next.catch(() => false);
    return next;
  }

  #substituteLoad: Promise<unknown> = Promise.resolve();

  async #loadMissingSubstitutes(): Promise<boolean> {
    const loaded = this.#fonts.substituteFamilies;
    if (this.#fonts.fixed || !loaded) return false;
    const missing = substituteFamiliesFor(this.store.package).filter((f) => !loaded.includes(f));
    if (missing.length === 0) return false;
    const more = await substituteFonts(missing);
    const now = this.#fonts;
    this.#fonts = {
      ...now,
      substitutes: [...now.substitutes, ...more.sources],
      substituteFailures: [...now.substituteFailures, ...more.failures],
      substituteFamilies: [...(now.substituteFamilies ?? []), ...missing],
    };
    this.#remeasure();
    return true;
  }

  /**
   * Switch from the fixed measurer to shaped measurement with the substitutes that loaded
   * late, and measure again. @internal
   */
  useShapedMeasurer(loaded: SubstituteLoad): void {
    if (this.#disposed) return;
    this.#fonts = {
      ...this.#fonts,
      fixed: false,
      substitutes: loaded.sources,
      substituteFailures: loaded.failures,
    };
    this.#remeasure();
  }

  /** Remove every host face. @internal */
  clearHostFaces(): void {
    this.setHostFaces(NO_HOST_FACES);
  }

  #remeasure(): void {
    const built = buildMeasurer(this.#fonts);
    // The fits of the fixed measurer fit again when shaping starts (`late-fit.ts`).
    if (this.#fonts.fixed) carryFits(this.#measurer, built.measurer);
    this.#measurer = built.measurer;
    this.#faces = built.faces;
    this.#layouts = new WeakMap();
    this.#bySlide.clear();
    this.#substitutions = null;
    this.#generation += 1;
    this.pictures.fontsChanged(); // Metafile text draws with the faces.
  }

  /**
   * Whether the session measures with the fixed measurer: `OpenOptions.fixedMeasurer`, or until
   * `shaping` resolves. Read it through `editor.fontMeasurement()`. @internal
   */
  get fixedMeasurer(): boolean {
    return this.#fonts.fixed;
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

  /** Record that a slide of a progressive open did not parse: block commits, and tell. @internal */
  markDamaged(rejection: Extract<PackageRejection, { reason: 'xml' }>): void {
    if (this.damage) return;
    this.damage = rejection;
    // A stable code: the part name comes from the file.
    this.store.block('damaged-slide');
    this.onDamage?.(rejection);
  }

  /**
   * Whether slide `index` is still opening: its part has not parsed (`progressive-open.ts`).
   * Reading its layout parses it at once. `EditorSnapshot.pendingSlides` is the public count.
   * @internal
   */
  slidePending(index: number): boolean {
    const part = this.info.slides[index]?.partName;
    return !!part && isDeferredPending(this.store.package.parts.get(part));
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
    const bytes = target ? binaryPartOf(pkg, target) : undefined;
    const shown = bytes && this.pictures.shown(bytes, imageMime(target!), pkg, sourcePart);
    return shown ? this.pictures.url(target!, shown.picture, shown.key, sourcePart) : undefined;
  };

  /** The bytes of the picture part `target` as they draw (`MetafilePictures.drawable`). @internal */
  picture(target: string): DrawablePicture | undefined {
    const bytes = binaryPartOf(this.store.package, target);
    return bytes ? this.pictures.drawable(bytes, imageMime(target)) : undefined;
  }

  /**
   * Waits until the EMF and WMF pictures of slide `index`, or of every slide when you omit it,
   * have converted. Resolves `true` when all draw, also as a PNG from before a font change, and
   * `false` when one cannot convert, for a slide that does not exist, or after the wait.
   */
  picturesReady(index?: number, options: PicturesWait = {}): Promise<boolean> {
    return this.pictures.readySlides(this.store.package, this.info.slides, index, options);
  }

  /** The EMF and WMF pictures of slide `index`, or of every slide, each with its state. */
  pictureStatus(index?: number): PictureStatus[] {
    return this.pictures.statusOfSlides(this.store.package, this.info.slides, index);
  }

  /**
   * The natural size of an image relationship in EMU, from its header and resolution, at 72 DPI
   * when it states none: the size of a tiled picture. The stored pixels count, not a JPEG's
   * EXIF orientation. An EMF or WMF picture has the size of its frame. Paint uses it for tiled
   * pictures, and Crop > Fill and Fit for the aspect ratio. Undefined for an external target or
   * an unreadable header.
   */
  imageSize = (sourcePart: string, rId: string): { cx: number; cy: number } | undefined => {
    const pkg = this.store.package;
    const target = relTarget(pkg, sourcePart, rId);
    const bytes = target ? binaryPartOf(pkg, target) : undefined;
    const size = bytes ? imageSize(bytes) : null;
    if (size)
      return imageSizeEmu({ ...size, dpiX: size.dpiX ?? TILE_DPI, dpiY: size.dpiY ?? TILE_DPI });
    return bytes ? metafileSizeEmu(bytes) : undefined; // An EMF or WMF picture frame.
  };

  /** The families whose packaged substitutes the session loads; null when it loads none. */
  get substituteFamilies(): readonly DefaultFamily[] | null {
    return this.#fonts.substituteFamilies;
  }

  /** Whether `dispose` ran. @internal */
  get disposed(): boolean {
    return this.#disposed;
  }

  /** Release blob URLs. */
  dispose(): void {
    this.#disposed = true;
    this.pictures.dispose();
  }
}

/** Open `.pptx` bytes. Never throws for a malformed file: it returns the typed reason. */
export async function openPresentation(
  bytes: Uint8Array,
  options: OpenOptions = {}
): Promise<{ ok: true; session: DocumentSession } | { ok: false; error: OpenError }> {
  if (bytes.length === 0) return { ok: false, error: { reason: 'empty' } };
  // A slide that does not parse after the open marks the session at once, wherever it is read.
  let made: DocumentSession | null = null;
  let early: Extract<PackageRejection, { reason: 'xml' }> | null = null;
  const onReject = (rejection: Extract<PackageRejection, { reason: 'xml' }>): void => {
    if (made) made.markDamaged(rejection);
    else early ??= rejection;
  };
  const readAt = performance.now();
  const read = options.progressive
    ? readPptxDeferred(bytes, PPTX_ZIP_LIMITS, {
        defer: () => true,
        finish: partWithoutEditorMarks,
        onReject,
      })
    : readPptx(bytes);
  if (!read.ok) return { ok: false, error: read.error };
  // A long read ends its task: the first slide parses in the next one (`OPEN_STEP_MS`).
  if (options.progressive && performance.now() - readAt > OPEN_STEP_MS) await nextTask();
  const opened = options.progressive ? firstSlideOnly(read.pkg, options.progressive) : null;
  if (opened && !opened.ok) return { ok: false, error: opened.error };
  const { sources: embedded, refused } = embeddedFontSources(read.pkg);
  const store = openStore(read.pkg);
  const host = options.fonts ?? [];
  const base = {
    host,
    embedded,
    refused,
    configured: NO_HOST_FACES,
    substitutes: [] as readonly FontSource[],
    substituteFailures: [] as readonly SubstituteFontFailure[],
    substituteFamilies: null,
  };
  const make = (fonts: FontSet): DocumentSession => {
    made = new DocumentSession(store, fonts);
    made.opening = opened?.opening ?? null;
    if (early) made.markDamaged(early);
    return made;
  };
  if (options.fixedMeasurer) return { ok: true, session: make({ ...base, fixed: true }) };
  const families = options.substitutes === false ? null : substituteFamiliesFor(read.pkg);
  const parsedAt = performance.now();
  const shaping = initializeHarfBuzz().then(() =>
    families ? substituteFonts(families) : { sources: [], failures: [] }
  );
  const shaped = (loaded: SubstituteLoad): FontSet => ({
    ...base,
    substitutes: loaded.sources,
    substituteFailures: loaded.failures,
    substituteFamilies: families,
    fixed: false,
  });
  if (options.shapingWaitMs === undefined)
    return { ok: true, session: make(shaped(await shaping)) };
  // A failure to load the text engine opens with the fixed measurer, as a stall does.
  const settled = shaping.then(
    (loaded) => ({ loaded }),
    (error: unknown) => ({ error })
  );
  const inTime = await within(settled, options.shapingWaitMs, undefined);
  if (inTime && 'loaded' in inTime) {
    const session = make(shaped(inTime.loaded));
    session.parsedAt = parsedAt;
    return { ok: true, session };
  }
  const session = make({ ...base, substituteFamilies: families, fixed: true });
  session.parsedAt = parsedAt;
  session.shapingError = inTime?.error ?? null;
  // The shapes that edits fit with the fixed measurer fit again when shaping starts, unless
  // their XML is back to the opened package's (an undone edit).
  const openedPackage = session.store.package;
  session.shaping = inTime
    ? Promise.resolve(false)
    : settled.then((late) => {
        if (!('loaded' in late)) {
          session.shapingError = late.error;
          return false;
        }
        const fixed = session.measurer;
        session.useShapedMeasurer(late.loaded);
        // Slides that parsed while the measurer was fixed can name families the open did not
        // load: `addMissingSubstitutes` skipped them then.
        // The fit waits for those substitutes, but a stalled or failed load does not stop it.
        const refit = (): true => (refitEditedText(session, fixed, openedPackage), true);
        return within(session.addMissingSubstitutes(), LATE_SUBSTITUTE_WAIT_MS, undefined).then(
          refit,
          refit
        );
      });
  return { ok: true, session };
}

/**
 * Parse slide `startSlide` of a package read with every slide deferred, with its notes; then
 * the rest too when it is small. The background parse of what is left, or null. A deferred part
 * that did not parse here refuses the deck, as the reader does.
 */
function firstSlideOnly(
  pkg: PptxPackage,
  options: { readonly startSlide?: number }
):
  | { readonly ok: true; readonly opening: ProgressiveOpen | null }
  | { readonly ok: false; readonly error: OpenError } {
  const info = presentationInfo(pkg);
  const start = clampSlide(options.startSlide, info.slides.length);
  const first = info.slides[start];
  if (first) parseDeferredPart(pkg, first.partName);
  if (first?.notesPart) parseDeferredPart(pkg, first.notesPart);
  if (pendingDeferredBytes(pkg) < PROGRESSIVE_OPEN_CONTENT_BYTES) parseDeferredParts(pkg);
  const rejection = deferredRejection(pkg);
  if (rejection) return { ok: false, error: rejection };
  if (pendingDeferredBytes(pkg) === 0) return { ok: true, opening: null };
  return { ok: true, opening: new ProgressiveOpen(pkg, openingOrder(info, start)) };
}
