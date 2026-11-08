// Round-trip census of one deck, for comparison R in `eigenpal/pptx-eval` (`CONTRIBUTING.md`,
// "Fidelity harness").
//
// Pass 1: the editor's own engine opens the deck, saves it, and opens the saved copy. Pass 2:
// the same, after a scripted edit through store ops (`setText` on the first run, and a 1 pt move
// of the first shape with a transform), measured against the pass-1 copy. The result holds facts
// only: errors, whether fingerprints match, part types, and XML element and attribute names. It
// never holds paths, file names, or deck text. The caller owns the pass and fail policy.
//
// Both sides of each comparison go through the engine's own reader, so the census cannot see
// what the reader drops (comments, processing instructions, insignificant whitespace). It does
// not show that other consumers open the copy without a repair prompt; pptx-eval checks a sample
// of the kept copies.

import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  openPresentation,
  type DocumentSession,
} from '../../packages/core/src/editor/document-session.ts';
import {
  MAX_COORDINATE,
  MIN_COORDINATE,
} from '../../packages/core/src/automation/commands/package-edits.ts';
import {
  attr,
  canonicalPackageFingerprint,
  changedParts,
  child,
  contentTypeOf,
  NS,
  packageFingerprints,
  presentationInfo,
  semanticDigest,
  writePptx,
  xmlPartOf,
  type Element,
  type Node,
  type PackageRejection,
  type PptxOp,
  type PptxPackage,
} from '../../packages/core/src/store/index.ts';

/**
 * The text the scripted edit writes. Synthetic, so the result never carries deck text. It holds
 * the five XML special characters and a character outside the BMP, to test escaping on save.
 */
export const EDIT_TEXT = 'pptx-eval & <round-trip> "edit" \'1\' \u{1F600}';
/** Names the scripted edit, so a caller can tell a recipe change from an engine change. */
export const EDIT_RECIPE = 'pptx-eval-settext-move-v2';
/** The scripted move: 1 pt on each axis, in EMU. */
export const EDIT_MOVE_EMU = 12_700;
/** At most this many differences are listed per comparison; `count` has them all. */
export const LIST_LIMIT = 20;
/** The version of round-trip mode's `run.json` and `results.jsonl` format. */
export const ROUNDTRIP_SCHEMA = 4;

/** The step a deck is in. `compare` and `edit` failures keep the facts of the steps before. */
export type CensusStage = 'open' | 'write' | 'reopen' | 'digest' | 'compare' | 'edit';

/** Where a deck failed, and why, in codes only. */
export interface CensusError {
  readonly stage: CensusStage;
  readonly reason: string;
  readonly detail?: string;
  /** The part type of the part that failed to parse, for `xml` rejections. */
  readonly partType?: string;
}

/** One part that differs between two packages. */
export interface PartDifference {
  readonly type: string;
  readonly change: 'changed' | 'added' | 'removed';
  /** XML parts: the element path of the first difference, as `prefix:local` names. */
  readonly path?: string;
  /** What differs there: `text`, `attribute a:x value`, `child removed p:extLst`, `bytes`, … */
  readonly what?: string;
}

export interface Differences {
  readonly count: number;
  readonly items: readonly PartDifference[];
}

export interface EditOutcome {
  readonly edit: 'setText' | 'move';
  readonly result: 'applied' | 'refused' | 'no-target';
  /** The store's rejection code for a refused edit. */
  readonly code?: string;
}

export interface DeckCensus {
  /** The first error of pass 1; later fields are null after an open, write, or reopen error. */
  readonly error: CensusError | null;
  readonly slides: number | null;
  readonly parts: number | null;
  readonly roundTrip: {
    /** The canonical fingerprint of the opened deck equals that of the reopened saved copy. */
    readonly fingerprintEqual: boolean;
    /** XML parts that the engine parses and compares by canonical form. */
    readonly xmlParts: number;
    /** Parts that the engine carries as bytes and writes back unchanged: equal by design. */
    readonly binaryParts: number;
    /** Saving the reopened copy again keeps its fingerprint (`semanticDigest`). */
    readonly digestStable: boolean | null;
    /** Parts whose canonical form differs. */
    readonly differences: Differences;
    /** Parts with equal canonical forms whose prefixes or namespace declarations differ. */
    readonly namespaceDifferences: Differences;
  } | null;
  readonly edit: {
    readonly edits: readonly EditOutcome[];
    readonly error: CensusError | null;
    readonly editedParts: number;
    /** Every edited part keeps its edited content after save and reopen. */
    readonly fingerprintEqual: boolean | null;
    /** Edited parts whose reopened content differs from the edited content. */
    readonly differences: Differences;
    /** Parts that no edit touched, but that differ from the pass-1 copy after save and reopen. */
    readonly collateral: Differences;
    /** Edited parts that are back to their pass-1 content after save and reopen. */
    readonly lost: number;
    /** Undoing every applied edit gives back the package before the edit exactly. */
    readonly undoRestores: boolean | null;
  } | null;
}

/** The facts of pass 1, which a caller keeps when the edit step crashes or hangs. */
export type PassOneFacts = Omit<DeckCensus, 'edit'>;

/** Called with each saved copy, so the caller can keep it (for a later check outside this repository). */
export type KeepCopy = (kind: 'saved' | 'edited', bytes: Uint8Array) => void;
/**
 * Called when the deck enters a step, so a caller can name the step of a deck that hangs. On
 * `edit`, `facts` holds the finished pass 1.
 */
export type OnStage = (stage: CensusStage, facts?: PassOneFacts) => void;

export interface CensusOptions {
  readonly keep?: KeepCopy;
  readonly onStage?: OnStage;
  /** The text of the `setText` edit; tests use it to make the store refuse the edit. */
  readonly editText?: string;
}

// ---------------------------------------------------------------- labels (no deck content)

export const TYPE_PREFIXES: readonly (readonly [string, string])[] = [
  ['application/vnd.openxmlformats-officedocument.presentationml.', 'pml:'],
  ['application/vnd.openxmlformats-officedocument.drawingml.', 'dml:'],
  ['application/vnd.openxmlformats-officedocument.', 'office:'],
  ['application/vnd.openxmlformats-package.', 'package:'],
  ['application/vnd.ms-', 'ms:'],
  ['image/', 'image/'],
  ['video/', 'video/'],
  ['audio/', 'audio/'],
  ['font/', 'font/'],
  ['application/', 'application/'],
];
/**
 * The content types that keep their name in the output, after their prefix label. A deck can
 * declare any content type; any other one shows as `<prefix>other`, so a word that the deck's
 * author chose never reaches the output.
 */
export const KNOWN_TYPES: ReadonlyMap<string, string> = new Map(
  [
    'pml:presentation.main pml:slideshow.main pml:template.main pml:slide pml:slideLayout',
    'pml:slideMaster pml:notesSlide pml:notesMaster pml:handoutMaster pml:presProps pml:viewProps',
    'pml:tableStyles pml:comments pml:commentAuthors pml:tags pml:slideUpdateInfo',
    'dml:chart dml:chartshapes dml:diagramData dml:diagramLayout dml:diagramStyle dml:diagramColors',
    'dml:drawing office:theme office:themeOverride office:vmlDrawing office:oleObject',
    'office:extended-properties office:custom-properties office:customXmlProperties',
    'office:obfuscatedFont office:spreadsheetml.sheet office:wordprocessingml.document',
    'package:relationships package:core-properties package:digital-signature-xmlsignature',
    'package:digital-signature-origin package:digital-signature-certificate',
    'ms:office.chartstyle ms:office.chartcolorstyle ms:office.drawingml.diagramDrawing',
    'ms:office.activeX ms:office.vbaProject ms:office.webextension ms:office.webextensiontaskpanes',
    'ms:powerpoint.revisioninfo ms:powerpoint.changesinfo ms:powerpoint.comments',
    'ms:powerpoint.authors ms:office.classificationlabels ms:excel.sheet.macroEnabled.main',
    'image/png image/jpeg image/gif image/bmp image/tiff image/x-emf image/x-wmf image/emf',
    'image/wmf image/svg image/x-icon image/webp image/vnd.ms-photo image/x-pict image/jpg',
    'video/mp4 video/quicktime video/x-ms-wmv video/mpeg video/x-msvideo video/avi',
    'audio/mpeg audio/wav audio/x-wav audio/mp3 audio/x-m4a audio/mp4 audio/x-ms-wma',
    'font/ttf font/otf font/sfnt font/woff font/woff2 application/xml application/octet-stream',
    'application/x-fontdata application/x-font-ttf application/pdf application/binary',
  ]
    .join(' ')
    .split(' ')
    // Content types compare without case; the label is the listed spelling.
    .map((name) => [name.toLowerCase(), name])
);

/**
 * A short label for a content type: a fixed prefix and at most 60 safe characters of the rest.
 * A type with no known prefix is `other`.
 */
export function typeLabel(contentType: string | undefined): string {
  if (contentType === undefined) return 'no-content-type';
  const type = contentType.trim().toLowerCase();
  for (const [prefix, label] of TYPE_PREFIXES) {
    if (!type.startsWith(prefix)) continue;
    const name = `${label}${type.slice(prefix.length).replace(/\+xml$/, '')}`;
    return KNOWN_TYPES.get(name) ?? `${label}other`;
  }
  return 'other';
}

/** The part type of a part: its content type label, or `rels:` and the type of its source. */
export function partType(pkg: PptxPackage, partName: string): string {
  if (partName === '/[Content_Types].xml') return 'content-types';
  if (partName.toLowerCase() === '/_rels/.rels') return 'rels:package';
  const rels = /^(.*)\/_rels\/([^/]+)\.rels$/i.exec(partName);
  if (rels) return `rels:${typeLabel(contentTypeOf(pkg, `${rels[1]}/${rels[2]}`))}`;
  return typeLabel(contentTypeOf(pkg, partName));
}

/** A part type from the name alone, for parts of a package that did not open. */
function partTypeByName(partName: string): string {
  if (partName === '/[Content_Types].xml') return 'content-types';
  const m = /^\/ppt\/(slides|slideLayouts|slideMasters|notesSlides|notesMasters|theme)\//.exec(
    partName
  );
  return m ? `ppt/${m[1]}` : 'other';
}

const PREFIX_OF = new Map<string, string>(Object.entries(NS).map(([p, uri]) => [uri, p]));

/**
 * The element and attribute names of the ECMA-376 Transitional schemas in `reference/`. A deck
 * can hold any name; only these names reach the output.
 */
const SCHEMA_NAMES = ((): { element: Set<string>; attribute: Set<string> } => {
  const folder = join(import.meta.dir, '../../reference/ecma-376/part4/transitional-schemas');
  const names = { element: new Set<string>(), attribute: new Set<string>() };
  for (const file of readdirSync(folder).filter((f) => f.endsWith('.xsd'))) {
    const text = readFileSync(join(folder, file), 'utf8');
    for (const m of text.matchAll(/<xsd:(element|attribute)\b[^>]*\bname="([A-Za-z_][\w.-]*)"/g))
      names[m[1] as 'element' | 'attribute'].add(m[2]);
  }
  return names;
})();

/**
 * `prefix:local` with the engine's conventional prefix. A name that the schemas do not define is
 * `?`, and an element of no namespace or of a namespace the engine does not know is `x:?`: such a
 * name could be a word that the deck's author chose.
 */
function nameOf(ns: string, local: string, kind: 'element' | 'attribute' = 'element'): string {
  const prefix = ns === '' && kind === 'attribute' ? '' : PREFIX_OF.get(ns);
  if (prefix === undefined) return 'x:?';
  const name = SCHEMA_NAMES[kind].has(local) ? local : '?';
  return prefix === '' ? name : `${prefix}:${name}`;
}

// ---------------------------------------------------------------- differences

const attrKey = (a: Element['attrs'][number]): string =>
  `{${a.ns}}${a.local}=${JSON.stringify(a.value)}`;
const nodeName = (n: Node): string => (n.type === 'text' ? '#text' : nameOf(n.ns, n.local));

/**
 * The first difference between two trees, in document order, the way `canonicalForm` sees them
 * (names, attribute values, children, text). With `namespaces`, element and attribute prefixes
 * and namespace declarations count too. Text values are never returned.
 */
export function firstDifference(
  a: Node,
  b: Node,
  namespaces = false,
  path: readonly string[] = []
): { path: string; what: string } | undefined {
  const at = (what: string): { path: string; what: string } => ({
    path: path.slice(-12).join('/'),
    what,
  });
  if (a.type === 'text' || b.type === 'text') {
    if (a.type !== b.type) return at(`node ${nodeName(a)} -> ${nodeName(b)}`);
    return a.type === 'text' && b.type === 'text' && a.value !== b.value ? at('text') : undefined;
  }
  if (a.ns !== b.ns || a.local !== b.local) return at(`element ${nodeName(a)} -> ${nodeName(b)}`);
  const here = [...path, nameOf(a.ns, a.local)];
  const inner = (what: string): { path: string; what: string } => ({
    path: here.slice(-12).join('/'),
    what,
  });
  const ka = a.attrs.map(attrKey).sort().join('\n');
  const kb = b.attrs.map(attrKey).sort().join('\n');
  if (ka !== kb) {
    const names = [...new Set([...a.attrs, ...b.attrs].map((x) => `${x.ns}\n${x.local}`))].sort();
    for (const key of names) {
      const [ns, local] = key.split('\n');
      const va = a.attrs.find((x) => x.ns === ns && x.local === local)?.value;
      const vb = b.attrs.find((x) => x.ns === ns && x.local === local)?.value;
      if (va === vb) continue;
      const change = va === undefined ? 'added' : vb === undefined ? 'removed' : 'value';
      return inner(`attribute ${nameOf(ns, local, 'attribute')} ${change}`);
    }
  }
  if (namespaces) {
    if ((a.prefix ?? '') !== (b.prefix ?? '')) return inner('element prefix');
    const da = a.nsDecls.map((d) => `${d.prefix}=${d.uri}`).sort();
    const db = b.nsDecls.map((d) => `${d.prefix}=${d.uri}`).sort();
    if (da.join('\n') !== db.join('\n'))
      return inner(`namespace declarations ${da.length} -> ${db.length}`);
    const pa = a.attrs.map((x) => `${x.ns}\n${x.local}\n${x.prefix ?? ''}`).sort();
    const pb = b.attrs.map((x) => `${x.ns}\n${x.local}\n${x.prefix ?? ''}`).sort();
    if (pa.join('\n') !== pb.join('\n')) return inner('attribute prefix');
  }
  // Look into the leading children that line up first, so an earlier change is not hidden by a
  // later added or removed child.
  const aligned = Math.min(a.children.length, b.children.length);
  for (let i = 0; i < aligned; i += 1) {
    if (nodeName(a.children[i]) !== nodeName(b.children[i])) break;
    const d = firstDifference(a.children[i], b.children[i], namespaces, here);
    if (d) return d;
  }
  if (a.children.length !== b.children.length) {
    // Name the first child whose count differs, so a reordering does not name the wrong one.
    const counts = (nodes: readonly Node[]): Map<string, number> => {
      const m = new Map<string, number>();
      for (const n of nodes) m.set(nodeName(n), (m.get(nodeName(n)) ?? 0) + 1);
      return m;
    };
    const ca = counts(a.children);
    const cb = counts(b.children);
    for (const n of [...a.children, ...b.children]) {
      const name = nodeName(n);
      const na = ca.get(name) ?? 0;
      const nb = cb.get(name) ?? 0;
      if (na !== nb) return inner(`child ${na > nb ? 'removed' : 'added'} ${name}`);
    }
  }
  for (let i = 0; i < aligned; i += 1) {
    if (nodeName(a.children[i]) === nodeName(b.children[i])) continue;
    return inner(`child ${nodeName(a.children[i])} -> ${nodeName(b.children[i])}`);
  }
  return undefined;
}

const FINGERPRINTS = new WeakMap<PptxPackage, ReadonlyMap<string, string>>();

/** `packageFingerprints`, computed once per package. */
function fingerprints(pkg: PptxPackage): ReadonlyMap<string, string> {
  let value = FINGERPRINTS.get(pkg);
  if (!value) {
    value = packageFingerprints(pkg);
    FINGERPRINTS.set(pkg, value);
  }
  return value;
}

/** The parts whose canonical fingerprints differ between two packages, or among `only`. */
export function packageDifferences(
  before: PptxPackage,
  after: PptxPackage,
  only?: ReadonlySet<string>
): Differences {
  const fa = fingerprints(before);
  const fb = fingerprints(after);
  const items: PartDifference[] = [];
  let count = 0;
  const add = (item: PartDifference): void => {
    count += 1;
    if (items.length < LIST_LIMIT) items.push(item);
  };
  for (const [name, print] of fa) {
    if (only && !only.has(name)) continue;
    if (!fb.has(name)) {
      add({ type: partType(before, name), change: 'removed' });
      continue;
    }
    if (fb.get(name) === print) continue;
    const pa = before.parts.get(name)!;
    const pb = after.parts.get(name)!;
    let found: { path: string; what: string } | undefined;
    if (pa.kind === 'xml' && pb.kind === 'xml') found = firstDifference(pa.xml.root, pb.xml.root);
    else if (pa.kind !== pb.kind) found = { path: '', what: `kind ${pa.kind} -> ${pb.kind}` };
    add({
      type: partType(before, name),
      change: 'changed',
      path: found?.path ?? '',
      what: found?.what ?? (pa.kind === 'binary' ? 'bytes' : 'canonical form'),
    });
  }
  for (const name of fb.keys())
    if (!fa.has(name) && (!only || only.has(name)))
      add({ type: partType(after, name), change: 'added' });
  return { count, items };
}

/** XML parts with equal canonical forms whose prefixes or namespace declarations differ. */
export function namespaceDifferences(before: PptxPackage, after: PptxPackage): Differences {
  const items: PartDifference[] = [];
  let count = 0;
  const fa = fingerprints(before);
  const fb = fingerprints(after);
  for (const [name, pa] of before.parts) {
    const pb = after.parts.get(name);
    if (pa.kind !== 'xml' || pb?.kind !== 'xml') continue;
    // Only parts with equal canonical forms can differ in namespaces alone.
    if (fa.get(name) !== fb.get(name)) continue;
    const d = firstDifference(pa.xml.root, pb.xml.root, true);
    if (!d) continue;
    count += 1;
    if (items.length < LIST_LIMIT)
      items.push({ type: partType(before, name), change: 'changed', ...d });
  }
  return { count, items };
}

// ---------------------------------------------------------------- the scripted edit

/** The `a:t` of the first run of the first slide that has one. */
function textTarget(pkg: PptxPackage): { part: string; node: string } | undefined {
  const firstRunText = (el: Element): Element | undefined => {
    if (el.kind === 'run') {
      const t = el.children.find((c): c is Element => c.type === 'element' && c.kind === 'text');
      if (t) return t;
    }
    for (const c of el.children) {
      if (c.type !== 'element') continue;
      const hit = firstRunText(c);
      if (hit) return hit;
    }
    return undefined;
  };
  for (const slide of presentationInfo(pkg).slides) {
    const root = xmlPartOf(pkg, slide.partName)?.root;
    const t = root && firstRunText(root);
    if (t) return { part: slide.partName, node: t.id };
  }
  return undefined;
}

/** The `a:off` of the first top-level shape with an explicit transform, and its position. */
function moveTarget(pkg: PptxPackage): { part: string; off: Element } | undefined {
  const findTree = (el: Element): Element | undefined => {
    if (el.kind === 'spTree') return el;
    for (const c of el.children) {
      const hit = c.type === 'element' ? findTree(c) : undefined;
      if (hit) return hit;
    }
    return undefined;
  };
  for (const slide of presentationInfo(pkg).slides) {
    const root = xmlPartOf(pkg, slide.partName)?.root;
    const tree = root && findTree(root);
    for (const el of tree?.children ?? []) {
      if (el.type !== 'element') continue;
      const xfrm =
        el.kind === 'graphicFrame'
          ? child(el, 'p:xfrm')
          : el.kind === 'grpSp'
            ? child(child(el, 'p:grpSpPr'), 'a:xfrm')
            : el.kind === 'sp' || el.kind === 'pic' || el.kind === 'cxnSp'
              ? child(child(el, 'p:spPr'), 'a:xfrm')
              : undefined;
      const off = child(xfrm, 'a:off');
      if (
        off &&
        coordinate(attr(off, 'x')) !== undefined &&
        coordinate(attr(off, 'y')) !== undefined
      )
        return { part: slide.partName, off };
    }
  }
  return undefined;
}

/** An `ST_Coordinate` value: a decimal integer within the schema's range. */
function coordinate(value: string | undefined): number | undefined {
  if (value === undefined || !/^-?\d{1,15}$/.test(value)) return undefined;
  const n = Number(value);
  return n >= MIN_COORDINATE && n <= MAX_COORDINATE ? n : undefined;
}

function moveOps(target: { part: string; off: Element }): PptxOp[] {
  const nudge = (axis: 'x' | 'y'): PptxOp => {
    const value = coordinate(attr(target.off, axis))!;
    const moved =
      value + EDIT_MOVE_EMU <= MAX_COORDINATE ? value + EDIT_MOVE_EMU : value - EDIT_MOVE_EMU;
    return {
      op: 'setAttribute',
      part: target.part,
      node: target.off.id,
      name: { ns: '', local: axis },
      value: String(moved),
    };
  };
  return [nudge('x'), nudge('y')];
}

// ---------------------------------------------------------------- one deck

function rejection(
  stage: CensusError['stage'],
  error: PackageRejection | { reason: 'empty' }
): CensusError {
  if (error.reason === 'zip') return { stage, reason: 'zip', detail: error.detail };
  if (error.reason === 'xml')
    return { stage, reason: 'xml', detail: error.detail, partType: partTypeByName(error.partName) };
  return { stage, reason: error.reason };
}

/**
 * A fault of the caller's machine (a kept copy or a progress file that cannot be written), not
 * of the engine. The census never records it as a deck's error: it goes up to the caller.
 */
export class MachineFault extends Error {
  override name = 'MachineFault';
}

export const exception = (stage: CensusStage, e: unknown): CensusError => ({
  stage,
  reason: 'exception',
  detail: e instanceof Error ? e.name : typeof e,
});

async function open(
  bytes: Uint8Array,
  stage: 'open' | 'reopen'
): Promise<{ session: DocumentSession } | { error: CensusError }> {
  try {
    const r = await openPresentation(bytes, { fixedMeasurer: true });
    return r.ok ? { session: r.session } : { error: rejection(stage, r.error) };
  } catch (e) {
    if (e instanceof MachineFault) throw e;
    return { error: exception(stage, e) };
  }
}

function save(pkg: PptxPackage): { bytes: Uint8Array } | { error: CensusError } {
  try {
    return { bytes: writePptx(pkg) };
  } catch (e) {
    if (e instanceof MachineFault) throw e;
    return { error: exception('write', e) };
  }
}

/** Run both passes on one deck. Never throws for a bad deck: every failure is a `CensusError`. */
export async function censusDeck(
  bytes: Uint8Array,
  options: CensusOptions = {}
): Promise<DeckCensus> {
  const { keep, onStage } = options;
  const empty = { slides: null, parts: null, roundTrip: null, edit: null };
  onStage?.('open');
  const opened = await open(bytes, 'open');
  if ('error' in opened) return { error: opened.error, ...empty };
  let reopened: { session: DocumentSession } | { error: CensusError };
  let pkg: PptxPackage;
  let slides: number;
  let parts: number;
  try {
    pkg = opened.session.store.package;
    slides = presentationInfo(pkg).slides.length;
    parts = pkg.parts.size;
    onStage?.('write');
    const saved = save(pkg);
    if ('error' in saved) return { ...empty, error: saved.error, slides, parts };
    keep?.('saved', saved.bytes);
    onStage?.('reopen');
    reopened = await open(saved.bytes, 'reopen');
    if ('error' in reopened) return { ...empty, error: reopened.error, slides, parts };
  } finally {
    opened.session.dispose();
  }
  const session = reopened.session;
  try {
    const copy = session.store.package;
    let digestStable: boolean | null = null;
    let error: CensusError | null = null;
    let roundTrip: DeckCensus['roundTrip'] = null;
    onStage?.('digest');
    try {
      digestStable = semanticDigest(copy) === canonicalPackageFingerprint(copy);
    } catch (e) {
      if (e instanceof MachineFault) throw e;
      error = exception('digest', e);
    }
    onStage?.('compare');
    try {
      const differences = packageDifferences(pkg, copy);
      const kinds = [...pkg.parts.values()].map((part) => part.kind);
      roundTrip = {
        fingerprintEqual: differences.count === 0,
        xmlParts: kinds.filter((kind) => kind === 'xml').length,
        binaryParts: kinds.filter((kind) => kind !== 'xml').length,
        digestStable,
        differences,
        namespaceDifferences: namespaceDifferences(pkg, copy),
      };
    } catch (e) {
      if (e instanceof MachineFault) throw e;
      error ??= exception('compare', e);
    }
    if (roundTrip === null) return { ...empty, error, slides, parts };
    const facts: PassOneFacts = { error, slides, parts, roundTrip };
    onStage?.('edit', facts);
    let edit: DeckCensus['edit'];
    try {
      edit = await editPass(session, options.editText ?? EDIT_TEXT, keep);
    } catch (e) {
      if (e instanceof MachineFault) throw e;
      edit = { ...EMPTY_EDIT, error: exception('edit', e) };
    }
    return { ...facts, edit };
  } finally {
    session.dispose();
  }
}

const NONE: Differences = { count: 0, items: [] };
const EMPTY_EDIT: NonNullable<DeckCensus['edit']> = {
  edits: [],
  error: null,
  editedParts: 0,
  fingerprintEqual: null,
  differences: NONE,
  collateral: NONE,
  lost: 0,
  undoRestores: null,
};

/**
 * Pass 2's comparison. The edited parts are compared with their reopened content (`differences`,
 * and `lost` when they are back to `before`); every other part with `before` (`collateral`).
 */
export function editResult(
  before: PptxPackage,
  edited: PptxPackage,
  after: PptxPackage,
  touched: ReadonlySet<string>
): Pick<
  NonNullable<DeckCensus['edit']>,
  'fingerprintEqual' | 'differences' | 'collateral' | 'lost'
> {
  const differences = packageDifferences(edited, after, touched);
  const base = fingerprints(before);
  const now = fingerprints(after);
  let lost = 0;
  for (const name of touched)
    if ((base.has(name) || now.has(name)) && base.get(name) === now.get(name)) lost += 1;
  // Every other part against `before`, with the element path and what differs, as pass 1 gives.
  const untouched = new Set([...base.keys(), ...now.keys()].filter((name) => !touched.has(name)));
  const collateral = packageDifferences(before, after, untouched);
  return { fingerprintEqual: differences.count === 0, differences, collateral, lost };
}

/**
 * Pass 2: the scripted edit through store ops on the reopened pass-1 copy, then save and reopen.
 * Editing the copy keeps a pass-1 difference out of pass 2. Undoing the applied edits must give
 * back the pass-1 file exactly, as in DOCX Editor's evaluation probe.
 */
async function editPass(
  session: DocumentSession,
  text: string,
  keep?: KeepCopy
): Promise<DeckCensus['edit']> {
  const store = session.store;
  const before = store.package;
  const edits: EditOutcome[] = [];
  const apply = (edit: EditOutcome['edit'], ops: PptxOp[] | undefined): void => {
    if (!ops) {
      edits.push({ edit, result: 'no-target' });
      return;
    }
    const r = store.transact(ops, { label: 'pptx-eval round-trip edit' });
    edits.push(
      r.ok ? { edit, result: 'applied' } : { edit, result: 'refused', code: r.rejection.code }
    );
  };
  const target = textTarget(before);
  apply('setText', target && [{ op: 'setText', part: target.part, node: target.node, text }]);
  const move = moveTarget(before);
  apply('move', move && moveOps(move));
  const edited = store.package;
  const touched = new Set(changedParts(before, edited));
  const applied = edits.filter((e) => e.result === 'applied').length;
  for (let i = 0; i < applied; i += 1) store.undo();
  const undoRestores =
    applied === 0
      ? null
      : canonicalPackageFingerprint(store.package) === canonicalPackageFingerprint(before);
  const result = { ...EMPTY_EDIT, edits, editedParts: touched.size, undoRestores };
  if (touched.size === 0) return result;
  const saved = save(edited);
  if ('error' in saved) return { ...result, error: saved.error };
  keep?.('edited', saved.bytes);
  const reopened = await open(saved.bytes, 'reopen');
  if ('error' in reopened) return { ...result, error: reopened.error };
  try {
    return { ...result, ...editResult(before, edited, reopened.session.store.package, touched) };
  } finally {
    reopened.session.dispose();
  }
}
