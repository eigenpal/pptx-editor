// The OPC package: bounded read, part classification, relationships, and save.
//
// Pipeline: bytes → readZip (bounded, OPC names normalized) → [Content_Types].xml → each XML
// part the engine interprets through parseXmlPart (bounded, entity-free) → PptxPackage. Every
// other part (media, SVG, fonts, embeddings, XML the engine does not read) is carried as bytes
// and re-emitted byte-identical. Nothing here fetches anything: external relationship targets
// are recorded, never resolved.

import { readZip, writeZip, DEFAULT_ZIP_LIMITS, type ZipLimits } from '@eigenpal/ooxml-core';
import {
  resolveInternalTarget,
  validateExternalTarget,
  asciiFold,
  normalizePartName,
  partNameKey,
} from '@eigenpal/ooxml-core';
import { parseXmlPart, type Element, type XmlPart, type TreeRejection } from './tree.ts';
import { serializeXmlPart, canonicalFingerprint, bytesFingerprint } from './serialize.ts';
import { attr, children, elements } from './query.ts';
import { NS, REL } from './namespaces.ts';
import { repairPresentationRoot } from './presentation-repair.ts';
import { markFileTimings, repairTimingRoot } from './timing.ts';
import { sha256FontBytes } from '@eigenpal/ooxml-core';
import { DerivedParts } from './parts-map.ts';
import {
  deferredXmlPart,
  holdsStandIn,
  isDeferredPending,
  type DeferredPartFinish,
  type DeferredPartReject,
} from './deferred-parts.ts';

/** One part of the package. */
export type PackagePart =
  | { readonly kind: 'xml'; readonly partName: string; readonly xml: XmlPart }
  | { readonly kind: 'binary'; readonly partName: string; readonly bytes: Uint8Array };

/** The whole package. Immutable: an edit produces a new package that shares untouched parts. */
export interface PptxPackage {
  readonly parts: ReadonlyMap<string, PackagePart>;
}

/** Why a file was refused. */
export type PackageRejection =
  | { readonly reason: 'zip'; readonly detail: string }
  | { readonly reason: 'no-content-types' }
  | { readonly reason: 'not-a-presentation' }
  | { readonly reason: 'xml'; readonly partName: string; readonly detail: TreeRejection }
  | { readonly reason: 'too-many-parts' };

export type PackageResult =
  | { readonly ok: true; readonly pkg: PptxPackage }
  | { readonly ok: false; readonly error: PackageRejection };

export const CONTENT_TYPES_PART = '/[Content_Types].xml';
const decoder = new TextDecoder('utf-8', { fatal: false });

const PML = 'application/vnd.openxmlformats-officedocument.presentationml.';
const PPT_MACRO = 'application/vnd.ms-powerpoint.';

/**
 * The main part types of a presentation, a slide show, a template, and their macro-enabled
 * forms (ECMA-376 Part 1, 13.3.6). The engine reads them all the same way; a macro part stays
 * as inert bytes and is never run.
 */
const MAIN_TYPES = new Set(
  [
    `${PML}presentation.main+xml`,
    `${PML}slideshow.main+xml`,
    `${PML}template.main+xml`,
    `${PPT_MACRO}presentation.macroEnabled.main+xml`,
    `${PPT_MACRO}slideshow.macroEnabled.main+xml`,
    `${PPT_MACRO}template.macroEnabled.main+xml`,
  ].map(asciiFold)
);

/** Content types of the parts the engine cannot work without. They must parse, or the deck is refused. */
const REQUIRED_TYPES = new Set(
  [
    ...MAIN_TYPES,
    `${PML}slide+xml`,
    `${PML}slideLayout+xml`,
    `${PML}slideMaster+xml`,
    `${PML}notesSlide+xml`,
    `${PML}notesMaster+xml`,
    'application/vnd.openxmlformats-officedocument.theme+xml',
  ].map(asciiFold)
);

/** Content types the engine reads but can do without. A part that does not parse is kept as bytes. */
const OPTIONAL_TYPES = new Set([`${PML}tableStyles+xml`, `${PML}handoutMaster+xml`].map(asciiFold));

/**
 * How the reader treats a part.
 *
 * - `required`: parsed; a part that does not parse refuses the deck.
 * - `optional`: parsed; a part that does not parse is kept as opaque bytes.
 * - `binary`: never parsed. This covers media (an SVG is `image/svg+xml`, but the engine draws it
 *   as a picture and never reads its tree), and every XML part the engine does not interpret, so
 *   those are re-emitted byte-identical instead of normalized.
 */
type PartRole = 'required' | 'optional' | 'binary';

/** The source part of a `.rels` part (`/` for the package), or undefined if the name is not one. */
function relsSource(partName: string): string | undefined {
  if (partName === '/_rels/.rels') return '/';
  const m = /^(.*)\/_rels\/([^/]+)\.rels$/.exec(partName);
  return m ? `${m[1]}/${m[2]}` : undefined;
}

function roleByType(partName: string, contentType: string | undefined): PartRole {
  if (contentType === undefined)
    // No content type is a broken package; keep the old reading of `.xml` names, but never
    // let one refuse the deck.
    return asciiFold(partName).endsWith('.xml') ? 'optional' : 'binary';
  const type = asciiFold(contentType.trim());
  if (REQUIRED_TYPES.has(type)) return 'required';
  if (OPTIONAL_TYPES.has(type)) return 'optional';
  return 'binary';
}

/** The relationship types whose targets make up the deck: they must parse, whatever their type. */
const STRUCTURAL_RELS = new Set<string>([
  REL.officeDocument,
  REL.slide,
  REL.slideLayout,
  REL.slideMaster,
  REL.theme,
  REL.notesSlide,
  REL.notesMaster,
]);

/** What a structural relationship makes of its target: the main part, the deck, or notes. */
type StructuralRole = 'main' | 'deck' | 'notes';

const ROLE_RANK: Readonly<Record<StructuralRole, number>> = { notes: 1, deck: 2, main: 3 };

/** Content types of the notes parts: neither shows a slide. */
const NOTES_TYPES = new Set([`${PML}notesSlide+xml`, `${PML}notesMaster+xml`].map(asciiFold));
/** Content types of the masters whose theme shows no slide. */
const NOTES_MASTER_TYPES = new Set(
  [`${PML}notesMaster+xml`, `${PML}handoutMaster+xml`].map(asciiFold)
);

/** Whether the content type of `partName` is one of `set`. */
function typeIn(types: ContentTypeIndex, partName: string, set: ReadonlySet<string>): boolean {
  const type = lookupContentType(types, partName);
  return type !== undefined && set.has(asciiFold(type.trim()));
}

/**
 * The parts a structural relationship points to, with the role it gives them. A part's role
 * follows what points to it as well as its content type: a slide with a generic content type is
 * still a slide, and a part of a required type that no structural relationship reaches never
 * refuses the deck.
 *
 * Notes are not needed to show the slides. A part is notes when a notes relationship reaches it,
 * when its content type is a notes type (so another relationship that names it, such as a
 * hostile slide relationship, does not make a damaged notes part refuse the deck), or when it is
 * the theme of a notes master or a handout master. A slide or slide master of the main part
 * stays part of the deck, whatever its content type. Keys are folded part names (`partNameKey`),
 * as OPC part names compare. A part that several relationships reach keeps the strongest role:
 * the main part first, notes last.
 */
function structuralTargets(
  rels: ReadonlyMap<string, PackagePart>,
  types: ContentTypeIndex
): Map<string, StructuralRole> {
  const pkg: PptxPackage = { parts: rels };
  const out = new Map<string, StructuralRole>();
  // The slides and slide masters of the main part: a content type never makes them notes.
  const shown = new Set<string>();
  for (const main of relationshipsOf(pkg, '/'))
    if (main.type === REL.officeDocument && main.target !== undefined)
      for (const rel of relationshipsOf(pkg, main.target))
        if ((rel.type === REL.slide || rel.type === REL.slideMaster) && rel.target !== undefined)
          shown.add(partNameKey(rel.target));
  const notesTyped = (part: string, set: ReadonlySet<string>): boolean =>
    !shown.has(partNameKey(part)) && typeIn(types, part, set);
  for (const name of rels.keys()) {
    const source = relsSource(name);
    if (source === undefined) continue;
    const notesMaster = source !== '/' && notesTyped(source, NOTES_MASTER_TYPES);
    for (const rel of relationshipsOf(pkg, source)) {
      if (rel.target === undefined || !STRUCTURAL_RELS.has(rel.type)) continue;
      const key = partNameKey(rel.target);
      const role: StructuralRole =
        rel.type === REL.officeDocument
          ? 'main'
          : rel.type === REL.notesSlide ||
              rel.type === REL.notesMaster ||
              (rel.type === REL.theme && notesMaster) ||
              notesTyped(rel.target, NOTES_TYPES)
            ? 'notes'
            : 'deck';
      const was = out.get(key);
      if (!was || ROLE_RANK[role] > ROLE_RANK[was]) out.set(key, role);
    }
  }
  return out;
}

function partRole(
  partName: string,
  types: ContentTypeIndex,
  structural: ReadonlyMap<string, StructuralRole>
): PartRole {
  if (partName === CONTENT_TYPES_PART) return 'required';
  if (asciiFold(partName).endsWith('.rels')) {
    // Relationships are always parsed. Those of the package and of required parts must parse.
    const source = relsSource(partName);
    if (source === '/') return 'required';
    if (source === undefined) return 'optional';
    // Those of the main part must parse too. Damaged relationships of a slide, a layout, or a
    // master stay as bytes: that part draws without them, and the deck still opens.
    return structural.get(partNameKey(source)) === 'main' ? 'required' : 'optional';
  }
  const contentType = lookupContentType(types, partName);
  const role = roleByType(partName, contentType);
  const via = structural.get(partNameKey(partName));
  // A part that is not XML by its type or its name (a picture that a slide relationship names)
  // stays as bytes, whatever points to it.
  const xml = role !== 'binary' || (contentType !== undefined && isXmlType(contentType));
  if (via === undefined) return role === 'required' ? 'optional' : role;
  // A structural part named `.xml` under a type that is not XML may still be the slide: it
  // parses, and stays as bytes when it does not.
  if (!xml) return asciiFold(partName).endsWith('.xml') ? 'optional' : role;
  // Speaker notes are not needed to show the slides: a damaged notes part stays as bytes.
  if (via === 'notes') return 'optional';
  return 'required';
}

/** Whether a content type names XML: `application/xml`, `text/xml`, or any `+xml` type. */
function isXmlType(contentType: string): boolean {
  const type = asciiFold(contentType.trim()).split(';')[0]!.trim();
  return type.endsWith('+xml') || type === 'application/xml' || type === 'text/xml';
}

/**
 * Decode the bytes of an XML part: a UTF-16 byte order mark is honored, UTF-8 otherwise.
 *
 * @internal
 */
export function decodeXml(bytes: Uint8Array): string {
  // UTF-16 parts exist in the wild; honor a BOM. Default UTF-8.
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes);
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes);
  let text = decoder.decode(bytes);
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
  return text;
}

/**
 * The archive caps of a presentation when a host configures none. A slide is at least two
 * entries (its part and its relationships), and four with speaker notes, so the format-neutral
 * cap of 10,000 entries refused a deck of 5,000 slides. This cap admits about 10,000 slides with
 * notes. It stays below `ZIP_HARD_LIMITS`, and the byte and ratio caps do not change.
 */
export const PPTX_ZIP_LIMITS: ZipLimits = Object.freeze({
  ...DEFAULT_ZIP_LIMITS,
  maxEntries: 40_000,
});

/** How `readPptxDeferred` defers parts; see `deferred-parts.ts`. @internal */
export interface DeferredReadOptions {
  /**
   * Whether to keep the slide or notes slide `partName` as its bytes, to parse on first read.
   * Asked only for parts of those two content types; every other part parses as before.
   */
  readonly defer: (partName: string) => boolean;
  /** Runs on a deferred part's tree once it parses. */
  readonly finish?: DeferredPartFinish;
  /** Hears, at once, a deferred part that does not parse, wherever it is read. */
  readonly onReject?: DeferredPartReject;
}

/** Read a `.pptx` at the trust boundary. Never throws. */
export function readPptx(bytes: Uint8Array, limits: ZipLimits = PPTX_ZIP_LIMITS): PackageResult {
  return readPptxDeferred(bytes, limits);
}

/**
 * `readPptx`, where the slides and notes slides that `deferred` names stay unparsed until
 * first read (`deferred-parts.ts`), for a progressive open. @internal
 */
export function readPptxDeferred(
  bytes: Uint8Array,
  limits: ZipLimits,
  deferred?: DeferredReadOptions
): PackageResult {
  const zip = readZip(bytes, limits);
  if (!zip.ok) return { ok: false, error: { reason: 'zip', detail: zip.reason } };
  const ctBytes = zip.entries.get(CONTENT_TYPES_PART);
  if (!ctBytes) return { ok: false, error: { reason: 'no-content-types' } };
  const ct = parseXmlPart(CONTENT_TYPES_PART, decodeXml(ctBytes));
  if (!ct.ok)
    return { ok: false, error: { reason: 'xml', partName: CONTENT_TYPES_PART, detail: ct.reason } };
  const types = contentTypeIndex(ct.part);
  // Relationships first: what points to a part decides whether it must parse.
  const rels = new Map<string, PackagePart>();
  const relsTrees = new Map<string, ReturnType<typeof parseXmlPart>>();
  for (const [partName, data] of zip.entries) {
    if (!asciiFold(partName).endsWith('.rels')) continue;
    const tree = parseXmlPart(partName, decodeXml(data));
    relsTrees.set(partName, tree);
    if (tree.ok) rels.set(partName, { kind: 'xml', partName, xml: tree.part });
  }
  const structural = structuralTargets(rels, types);
  const parts = new Map<string, PackagePart>();
  for (const [partName, data] of zip.entries) {
    const role = partRole(partName, types, structural);
    if (role === 'binary') {
      parts.set(partName, { kind: 'binary', partName, bytes: data });
      continue;
    }
    const kind =
      deferred && structural.has(partNameKey(partName))
        ? deferrableKind(partName, types)
        : undefined;
    if (kind && deferred!.defer(partName)) {
      const finish = deferred!.finish;
      parts.set(
        partName,
        deferredXmlPart(partName, data, decodeXml, {
          notes: kind === 'notes',
          ...(deferred!.onReject ? { onReject: deferred!.onReject } : {}),
          finish: (tree) => {
            if (isSlideRoot(tree.root)) markFileTimings(tree.root);
            return finish ? finish(tree) : tree;
          },
        })
      );
      continue;
    }
    const tree =
      partName === CONTENT_TYPES_PART
        ? ct
        : (relsTrees.get(partName) ?? parseXmlPart(partName, decodeXml(data)));
    if (tree.ok) {
      parts.set(partName, { kind: 'xml', partName, xml: tree.part });
      if (isSlideRoot(tree.part.root)) markFileTimings(tree.part.root);
    } else if (role === 'optional') parts.set(partName, { kind: 'binary', partName, bytes: data });
    else return { ok: false, error: { reason: 'xml', partName, detail: tree.reason } };
  }
  const pkg: PptxPackage = { parts };
  const main = mainPresentationPart(pkg);
  const mainType = main ? contentTypeOf(pkg, main) : undefined;
  if (!mainType || !MAIN_TYPES.has(asciiFold(mainType.trim())))
    return { ok: false, error: { reason: 'not-a-presentation' } };
  return { ok: true, pkg };
}

const SLIDE_TYPE = asciiFold(`${PML}slide+xml`);

/** Whether `partName` is a slide or a notes slide, the parts a progressive open may defer. */
function deferrableKind(partName: string, types: ContentTypeIndex): 'slide' | 'notes' | undefined {
  const type = lookupContentType(types, partName);
  const folded = type === undefined ? undefined : asciiFold(type.trim());
  // Notes parse with the package: a damaged notes part stays as bytes, as `partRole` says.
  return folded === SLIDE_TYPE ? 'slide' : undefined;
}

/**
 * Write the package back to `.pptx` bytes. XML parts re-emit normalized; binaries verbatim. The
 * bytes own an `ArrayBuffer`, so they are a valid `BlobPart` as they are. A deferred part that
 * has not been read parses here; one that did not parse refuses the write (it throws), so a
 * stand-in never leaves the session.
 */
export function writePptx(pkg: PptxPackage): Uint8Array<ArrayBuffer> {
  for (const part of pkg.parts.values()) {
    if (isDeferredPending(part) && part.kind === 'xml') void part.xml;
    if (holdsStandIn(part)) throw new Error(`writePptx: ${part.partName} did not parse`);
  }
  const encoder = new TextEncoder();
  const entries = new Map<string, Uint8Array>();
  // [Content_Types].xml first: some readers sniff the first entry.
  const ordered = [...pkg.parts.values()].sort((a, b) =>
    a.partName === CONTENT_TYPES_PART ? -1 : b.partName === CONTENT_TYPES_PART ? 1 : 0
  );
  // The presentation part leaves with one extension list, one section list, and one slide id
  // for each slide that has a relationship and a part. A slide whose `p:timing` the session
  // changed leaves with no empty time node list or build list in it (`repairTimingRoot`).
  const main = mainPresentationPart(pkg);
  for (const part of ordered)
    entries.set(
      part.partName,
      part.kind === 'xml'
        ? encoder.encode(
            serializeXmlPart(
              part.partName === main
                ? { ...part.xml, root: repairedPresentationRoot(pkg) ?? part.xml.root }
                : isSlideRoot(part.xml.root)
                  ? { ...part.xml, root: repairTimingRoot(part.xml.root) }
                  : part.xml
            )
          )
        : part.bytes
    );
  const zip = writeZip(entries);
  // The archive is a fresh allocation, so its buffer is an `ArrayBuffer`; copy if it ever is not.
  return zip.buffer instanceof ArrayBuffer ? (zip as Uint8Array<ArrayBuffer>) : new Uint8Array(zip);
}

/**
 * The relationship ids of the presentation part that are internal and name a part of the
 * package, each to the part it names: a `p:sldId` whose `r:id` is not one of them names no
 * slide. A slide that one
 * collaborator deleted while another moved it can come back in `p:sldIdLst` without its
 * relationship. The part lookup folds case, as OPC part names do, so a target
 * that differs only in case stays.
 */
function liveRelationshipIds(pkg: PptxPackage, main: string): Map<string, string> {
  const live = new Map<string, string>();
  for (const r of relationshipsOf(pkg, main)) {
    const part = !r.external && r.target ? equivalentPartName(pkg, r.target) : undefined;
    if (part) live.set(r.id, part);
  }
  return live;
}

const isSlideRoot = (root: Element): boolean => root.ns === NS.p && root.local === 'sld';

const repaired = new WeakMap<PptxPackage, Element | undefined>();
/**
 * The package `repairedPresentationRoot` read last, by its presentation part. An edit that
 * changes the repair changes that part or its relationships, so a package with another
 * presentation part has nothing to share; two documents never share one.
 */
const repairedLast = new WeakMap<XmlPart, WeakRef<PptxPackage>>();

/**
 * The presentation root as the writer saves it (`repairPresentationRoot`), or undefined
 * without a presentation part. The root itself when it needs no repair. Memoized per package.
 * @internal
 */
export function repairedPresentationRoot(pkg: PptxPackage): Element | undefined {
  if (repaired.has(pkg)) return repaired.get(pkg);
  const main = mainPresentationPart(pkg);
  const xml = main ? xmlPartOf(pkg, main) : undefined;
  const prev = xml ? repairedLast.get(xml)?.deref() : undefined;
  // The repair reads the root relationships, the presentation part, its relationships, and
  // the part names. A package that differs from the last one only in other XML parts (a text
  // edit on a slide) has the same repaired root.
  const changed = prev && repaired.has(prev) ? changedPartNames(prev, pkg, 32) : null;
  const shared = changed?.every((name) => {
    const a = prev!.parts.get(name);
    const b = pkg.parts.get(name);
    return (
      a?.kind === 'xml' &&
      b?.kind === 'xml' &&
      name !== main &&
      !name.toLowerCase().endsWith('.rels')
    );
  });
  const out = shared
    ? repaired.get(prev!)
    : main && xml
      ? repairPresentationRoot(xml.root, liveRelationshipIds(pkg, main))
      : undefined;
  repaired.set(pkg, out);
  if (xml) repairedLast.set(xml, new WeakRef(pkg));
  return out;
}

// The slide id form that the writer's repair compares; `normalizePresentation` uses it too.
export { canonicalSlideId } from './presentation-repair.ts';

// ---------------------------------------------------------------- content types

interface ContentTypeIndex {
  readonly defaults: ReadonlyMap<string, string>;
  readonly overrides: ReadonlyMap<string, string>;
}

const ctCache = new WeakMap<XmlPart, ContentTypeIndex>();

function contentTypeIndex(part: XmlPart): ContentTypeIndex {
  const hit = ctCache.get(part);
  if (hit) return hit;
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  for (const el of elements(part.root)) {
    if (el.ns !== NS.ct) continue;
    const type = attr(el, 'ContentType');
    if (!type) continue;
    if (el.local === 'Default') {
      const ext = attr(el, 'Extension');
      if (ext) defaults.set(asciiFold(ext), type);
    } else if (el.local === 'Override') {
      // `PartName` is a URI, percent-encoded as the ZIP item name is. Decode it the same way,
      // once, so it matches the canonical part name the reader keys parts by.
      const name = attr(el, 'PartName');
      const norm = name ? normalizePartName(name) : undefined;
      if (norm?.ok) overrides.set(asciiFold(norm.partName), type);
    }
  }
  const index = { defaults, overrides };
  ctCache.set(part, index);
  return index;
}

function lookupContentType(index: ContentTypeIndex, partName: string): string | undefined {
  const over = index.overrides.get(asciiFold(partName));
  if (over) return over;
  const dot = partName.lastIndexOf('.');
  return dot < 0 ? undefined : index.defaults.get(asciiFold(partName.slice(dot + 1)));
}

/** The content type of a part, from `[Content_Types].xml`. */
export function contentTypeOf(pkg: PptxPackage, partName: string): string | undefined {
  const ct = pkg.parts.get(CONTENT_TYPES_PART);
  if (!ct || ct.kind !== 'xml') return undefined;
  return lookupContentType(contentTypeIndex(ct.xml), partName);
}

// ---------------------------------------------------------------- relationships

/** One relationship, resolved against its source part. */
export interface Relationship {
  readonly id: string;
  readonly type: string;
  /** Internal: the resolved part name. External: undefined. */
  readonly target?: string;
  /** External: the raw URI, retained verbatim and NEVER fetched. */
  readonly externalTarget?: string;
  readonly external: boolean;
}

/** The `.rels` part name for a source part (`/` for the package). */
export function relsPartName(source: string): string {
  if (source === '/') return '/_rels/.rels';
  const slash = source.lastIndexOf('/');
  return `${source.slice(0, slash)}/_rels/${source.slice(slash + 1)}.rels`;
}

const relCache = new WeakMap<XmlPart, readonly Relationship[]>();

/** All relationships of a source part. Unsafe or unresolvable targets are dropped. */
export function relationshipsOf(pkg: PptxPackage, source: string): readonly Relationship[] {
  const part = pkg.parts.get(relsPartName(source));
  if (!part || part.kind !== 'xml') return [];
  const hit = relCache.get(part.xml);
  if (hit) return hit;
  const out: Relationship[] = [];
  const owner = source === '/' ? '/x' : source;
  for (const el of children(part.xml.root, { ns: NS.rel, local: 'Relationship' })) {
    const id = attr(el, 'Id');
    const type = attr(el, 'Type');
    const target = attr(el, 'Target');
    if (!id || !type || target === undefined) continue;
    if (attr(el, 'TargetMode') === 'External') {
      const ok = validateExternalTarget(target);
      out.push({ id, type, external: true, ...(ok.ok ? { externalTarget: target } : {}) });
      continue;
    }
    const resolved = resolveInternalTarget(owner, target);
    if (resolved.ok) out.push({ id, type, target: resolved.partName, external: false });
  }
  relCache.set(part.xml, out);
  return out;
}

/** The relationships of one list by id; the first one wins for an id that repeats. */
const relById = new WeakMap<readonly Relationship[], ReadonlyMap<string, Relationship>>();

/**
 * The relationship of a source part with id `id`, or undefined: the first one, as a search of
 * `relationshipsOf` finds it. One map lookup: the presentation part names every slide by id.
 * @internal
 */
export function relationshipById(
  pkg: PptxPackage,
  source: string,
  id: string
): Relationship | undefined {
  const rels = relationshipsOf(pkg, source);
  let byId = relById.get(rels);
  if (!byId) {
    const map = new Map<string, Relationship>();
    for (const r of rels) if (!map.has(r.id)) map.set(r.id, r);
    relById.set(rels, (byId = map));
  }
  return byId.get(id);
}

/** The internal target of a relationship id, or undefined. */
export function relTarget(pkg: PptxPackage, source: string, id: string): string | undefined {
  const rel = relationshipById(pkg, source, id);
  return rel && !rel.external ? rel.target : undefined;
}

/**
 * The names of the parts that differ between two packages by object: added, removed, or
 * replaced. Null when more than `limit` differ. An edit replaces only the parts it changes, so
 * this is a short list after one edit. Two maps derived from one root are compared by their
 * recorded changes only; other maps in one pass over the part map. @internal
 */
export function changedPartNames(
  prev: PptxPackage,
  next: PptxPackage,
  limit: number
): string[] | null {
  // Two maps derived from one root (`parts-map.ts`) can differ only where they record changes.
  const candidates = DerivedParts.candidates(prev.parts, next.parts);
  if (candidates) {
    const found: string[] = [];
    for (const name of candidates) {
      if (prev.parts.get(name) === next.parts.get(name)) continue;
      found.push(name);
      if (found.length > limit) return null;
    }
    return found;
  }
  const out: string[] = [];
  let added = 0;
  for (const [name, part] of next.parts) {
    const old = prev.parts.get(name);
    if (old === part) continue;
    if (old === undefined) added += 1;
    out.push(name);
    if (out.length > limit) return null;
  }
  // Same names in both, less the added ones: anything else in `prev` was removed.
  if (prev.parts.size !== next.parts.size - added)
    for (const name of prev.parts.keys()) {
      if (next.parts.has(name)) continue;
      out.push(name);
      if (out.length > limit) return null;
    }
  return out;
}

/** The first internal target of a relationship type. */
export function relTargetByType(
  pkg: PptxPackage,
  source: string,
  type: string
): string | undefined {
  return relationshipsOf(pkg, source).find((r) => r.type === type && !r.external)?.target;
}

/** `/ppt/presentation.xml` (or wherever the package root points). */
export function mainPresentationPart(pkg: PptxPackage): string | undefined {
  return relTargetByType(pkg, '/', REL.officeDocument);
}

/**
 * The existing part whose name is OPC-equivalent to `partName` (ASCII case-insensitive), or
 * undefined. A new part name must not match one: the writer folds case, so `Image1.png` and
 * `image1.png` are one ZIP entry and a package holding both cannot be saved.
 */
export function equivalentPartName(pkg: PptxPackage, partName: string): string | undefined {
  if (pkg.parts.has(partName)) return partName;
  const key = partNameKey(partName);
  for (const name of pkg.parts.keys()) if (partNameKey(name) === key) return name;
  return undefined;
}

/** The XML tree of a part, or undefined. */
export function xmlPartOf(pkg: PptxPackage, partName: string | undefined): XmlPart | undefined {
  if (!partName) return undefined;
  const part = pkg.parts.get(partName);
  return part && part.kind === 'xml' ? part.xml : undefined;
}

/** The bytes of a binary part, or undefined. */
export function binaryPartOf(
  pkg: PptxPackage,
  partName: string | undefined
): Uint8Array | undefined {
  if (!partName) return undefined;
  const part = pkg.parts.get(partName);
  return part && part.kind === 'binary' ? part.bytes : undefined;
}

// ---------------------------------------------------------------- round-trip fingerprints

/**
 * Per-part canonical fingerprints. XML parts hash their canonical form; binaries hash bytes.
 * Two packages with the same map are the same document content, however each was written.
 */
export function packageFingerprints(pkg: PptxPackage): ReadonlyMap<string, string> {
  const out = new Map<string, string>();
  for (const part of [...pkg.parts.values()].sort((a, b) => (a.partName < b.partName ? -1 : 1)))
    out.set(
      part.partName,
      part.kind === 'xml' ? canonicalFingerprint(part.xml) : bytesFingerprint(part.bytes)
    );
  return out;
}

/** One hash over all part fingerprints: the package's canonical fingerprint. */
export function canonicalPackageFingerprint(pkg: PptxPackage): string {
  const lines = [...packageFingerprints(pkg)].map(([k, v]) => `${k}\t${v}`).join('\n');
  return sha256FontBytes(new TextEncoder().encode(lines));
}

/**
 * Save, reopen, and fingerprint: the semantic digest. Equal digests before and after a
 * round trip prove save loses nothing the model carries.
 */
export function semanticDigest(pkg: PptxPackage): string {
  const reopened = readPptx(writePptx(pkg));
  if (!reopened.ok)
    throw new Error(`save produced an unreadable package: ${reopened.error.reason}`);
  return canonicalPackageFingerprint(reopened.pkg);
}
