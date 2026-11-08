// Slide structure beyond `p:sldIdLst`: the parts a slide owns (comments, charts, diagrams,
// embedded objects), which a duplicate copies and a delete takes with it, and the section list
// (`p14:sectionLst`, ECMA-376 Part 1 §19 with the `p14` extension) that names slides by id.

import {
  attr,
  canonicalSlideId,
  CONTENT_TYPES_PART,
  contentTypeOf,
  elements,
  NS,
  parseXmlPart,
  presentationInfo,
  relationshipsOf,
  relsPartName,
  resolveInternalTarget,
  serializeXmlPart,
  xmlPartOf,
  type Element,
  type PptxOp,
  type PptxPackage,
  type XmlPart,
} from '../../store/index.ts';
import { partNameKey } from '@eigenpal/ooxml-core';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import {
  addOverride,
  freeIds,
  freePartName,
  MAX_MINTED_ID,
  MAX_UNSIGNED_INT,
  relativeTarget,
  removeOverrides,
  stripedIds,
  unsignedIntId,
  xmlAttr,
  XML_DECL,
} from './package-edits.ts';
import { toBase64 } from './base64.ts';
import { ACTOR_ID_STRIPE } from './actor-scoped-ids.ts';
import { newGuid } from './comment-write.ts';

/** The namespace of threaded comments. */
const THREADED_NS = 'http://schemas.microsoft.com/office/powerpoint/2018/8/main';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const MS2007 = 'http://schemas.microsoft.com/office/2007/relationships/';
const MS2011 = 'http://schemas.microsoft.com/office/2011/relationships/';
const P15 = 'http://schemas.microsoft.com/office/powerpoint/2012/main';

/** Comments by a single author in one list (`p:cmLst`), ECMA-376 Part 1 §19.4. */
const LEGACY_COMMENTS = `${R}comments`;
/** Threaded comments (`p188:cmLst`), which name their slide and their authors by id. */
const THREADED_COMMENTS = 'http://schemas.microsoft.com/office/2018/10/relationships/comments';
const CT_COMMENTS = 'application/vnd.openxmlformats-officedocument.presentationml.comments+xml';
const CT_AUTHORS =
  'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml';

/**
 * Targets a slide owns. A duplicate gets its own copy, so editing one slide never changes the
 * other. Their own relationships follow the same rules.
 */
const OWNED = new Set([
  `${R}chart`,
  `${R}chartUserShapes`,
  `${R}themeOverride`,
  `${R}package`,
  `${R}oleObject`,
  LEGACY_COMMENTS,
  THREADED_COMMENTS,
  `${R}tags`,
  `${R}vmlDrawing`,
  `${R}diagramData`,
  `${R}diagramLayout`,
  `${R}diagramQuickStyle`,
  `${R}diagramColors`,
  `${MS2007}diagramDrawing`,
  `${MS2011}chartStyle`,
  `${MS2011}chartColorStyle`,
]);

/**
 * Targets the engine cannot copy safely: embedded controls (`p:control`) carry binary state. A
 * duplicate of a slide with one is refused, never made to share it.
 */
const NOT_COPYABLE = new Set([`${R}control`]);

/** Never removed with a deleted slide: layouts, masters, themes, and other slides. */
const STRUCTURAL = new Set([
  `${R}slideLayout`,
  `${R}slideMaster`,
  `${R}theme`,
  `${R}notesMaster`,
  `${R}handoutMaster`,
  `${R}slide`,
  `${R}notesSlide`,
]);

/** Bounds for walking and copying part graphs from a file. */
const MAX_DEPTH = 8;
const MAX_PARTS = 2_000;

/** What a copied relationship points at: the same part, another part, or nothing. */
export type RelTarget =
  { readonly keep: true } | { readonly part: string } | { readonly drop: true };

/** The content-type override of a part, if it has one (a default by extension needs none). */
function overrideOf(pkg: PptxPackage, partName: string): string | undefined {
  const ct = xmlPartOf(pkg, CONTENT_TYPES_PART);
  const key = partNameKey(partName);
  const el = elements(ct?.root).find(
    (e) => e.local === 'Override' && partNameKey(attr(e, 'PartName') ?? '') === key
  );
  return el ? attr(el, 'ContentType') : undefined;
}

/** `/ppt/charts/chart12.xml` → [`/ppt/charts/chart`, `.xml`]: a free name keeps the pattern. */
function namePattern(partName: string): [string, string] {
  const slash = partName.lastIndexOf('/');
  const dot = partName.lastIndexOf('.');
  const ext = dot > slash ? partName.slice(dot) : '';
  const stem = partName.slice(0, partName.length - ext.length).replace(/\d+$/, '');
  return [stem, ext];
}

function decode(bytes: Uint8Array): string {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes);
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes);
  const text = new TextDecoder('utf-8').decode(bytes);
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

/**
 * The XML tree of a part. The reader keeps parts it does not interpret (comments, comment
 * authors, charts) as bytes; those are parsed here with the same bounded reader.
 */
export function treeOf(pkg: PptxPackage, partName: string | undefined): XmlPart | undefined {
  const part = partName ? pkg.parts.get(partName) : undefined;
  if (!part) return undefined;
  if (part.kind === 'xml') return part.xml;
  const r = parseXmlPart(part.partName, decode(part.bytes));
  return r.ok ? r.part : undefined;
}

/** Make a part held as bytes an XML part of the batch, so ops can edit it. */
export function editable(batch: OpBatch, partName: string): XmlPart | undefined {
  const xml = xmlPartOf(batch.pkg, partName);
  if (xml) return xml;
  const tree = treeOf(batch.pkg, partName);
  if (!tree) return undefined;
  batch.push({ op: 'removePart', part: partName });
  batch.push({ op: 'addPart', part: partName, xml: serializeXmlPart(tree) });
  return xmlPartOf(batch.pkg, partName);
}

/**
 * Copies the parts a duplicated slide owns, once each, under fresh names. Legacy comments get
 * new indexes, so every author's comment indexes stay unique in the presentation, and threaded
 * comments and replies get new GUIDs. A save points the copied threads at the new slide
 * (`comment-repair.ts`).
 */
export class PartCopier {
  readonly #copies = new Map<string, string>();
  readonly #comments: string[] = [];
  readonly #threaded: string[] = [];

  constructor(
    readonly batch: OpBatch,
    readonly presentation: string
  ) {}

  /**
   * The relationships part for `copy`, from those of `source`. `override` decides a
   * relationship first; otherwise owned targets are copied, shared ones kept, and a target
   * that cannot be copied refuses the whole command.
   */
  relsXml(
    source: string,
    copy: string,
    override?: (type: string, target: string | undefined) => RelTarget | undefined,
    depth = 0
  ): string {
    const rels = xmlPartOf(this.batch.pkg, relsPartName(source));
    const items: string[] = [];
    for (const r of elements(rels?.root)) {
      const type = attr(r, 'Type') ?? '';
      const raw = attr(r, 'Target') ?? '';
      const external = attr(r, 'TargetMode') === 'External';
      let target = raw;
      if (!external) {
        const resolved = resolveInternalTarget(source, raw);
        const part = resolved.ok ? resolved.partName : undefined;
        const decision = override?.(type, part) ?? this.#decide(type, part, depth);
        if ('drop' in decision) continue;
        if ('part' in decision) target = relativeTarget(copy, decision.part);
        else if (part) target = relativeTarget(copy, part);
      }
      items.push(
        `<Relationship Id="${xmlAttr(attr(r, 'Id') ?? '', 'relationship')}" Type="${xmlAttr(type, 'relationship')}" Target="${xmlAttr(target, 'relationship')}"${external ? ' TargetMode="External"' : ''}/>`
      );
    }
    return `${XML_DECL}<Relationships xmlns="${NS.rel}">${items.join('')}</Relationships>`;
  }

  #decide(type: string, part: string | undefined, depth: number): RelTarget {
    if (NOT_COPYABLE.has(type)) throw new CommandRefusal('disabled.slideNotCopyable', type);
    if (!part || !OWNED.has(type) || !this.batch.pkg.parts.has(part)) return { keep: true };
    const comments =
      type === LEGACY_COMMENTS ? 'legacy' : type === THREADED_COMMENTS ? 'threaded' : undefined;
    return { part: this.#copy(part, depth + 1, comments) };
  }

  /** Copy one owned part (and what it owns) under a free name; returns the new name. */
  #copy(source: string, depth: number, comments?: 'legacy' | 'threaded'): string {
    const done = this.#copies.get(source);
    if (done) return done;
    if (depth > MAX_DEPTH || this.#copies.size >= MAX_PARTS)
      throw new CommandRefusal('errors.tooLarge');
    const part = this.batch.pkg.parts.get(source)!;
    const [prefix, suffix] = namePattern(source);
    const name = freePartName(this.batch, prefix, suffix);
    this.#copies.set(source, name);
    // A comment list is copied as a tree, so its indexes can be renumbered.
    const tree = comments ? treeOf(this.batch.pkg, source) : undefined;
    if (comments && !tree) throw new CommandRefusal('errors.malformed', source);
    this.batch.push(
      tree
        ? { op: 'addPart', part: name, xml: serializeXmlPart(tree) }
        : part.kind === 'xml'
          ? { op: 'addPart', part: name, xml: serializeXmlPart(part.xml) }
          : { op: 'addPart', part: name, base64: toBase64(part.bytes) }
    );
    const contentType = overrideOf(this.batch.pkg, source);
    if (contentType) addOverride(this.batch, name, contentType);
    if (xmlPartOf(this.batch.pkg, relsPartName(source)))
      this.batch.push({
        op: 'addPart',
        part: relsPartName(name),
        xml: this.relsXml(source, name, undefined, depth),
      });
    if (comments === 'legacy') this.#comments.push(name);
    if (comments === 'threaded') this.#threaded.push(name);
    return name;
  }

  /** Give the copied comments new indexes and GUIDs. Call once, after the last copy. */
  finish(): void {
    if (this.#comments.length) renumberComments(this.batch, this.#comments, this.presentation);
    for (const part of this.#threaded) {
      const walk = (el: Element, depth: number): void => {
        if (depth > MAX_DEPTH) return;
        for (const c of elements(el)) {
          if (c.ns === THREADED_NS && (c.local === 'cm' || c.local === 'reply') && attr(c, 'id'))
            this.batch.push({
              op: 'setAttribute',
              part,
              node: c.id,
              name: { ns: '', local: 'id' },
              value: newGuid(),
            });
          walk(c, depth + 1);
        }
      };
      const root = xmlPartOf(this.batch.pkg, part)?.root;
      if (root) walk(root, 0);
    }
  }
}

/** How {@link renumberComments} reads copied lists whose author ids changed. */
export interface RenumberOptions {
  /** The author id an element had before the copy changed it. Default: its `authorId`. */
  readonly authorOf?: (part: string, el: Element) => string | undefined;
  /** Remove the threading extension of a reply whose parent is not in the list. */
  readonly dropOrphanReplies?: boolean;
}

/** The first element `ns`:`local` at or below `el`, at most `depth` levels down. */
function descendant(el: Element, ns: string, local: string, depth: number): Element | undefined {
  if (el.ns === ns && el.local === local) return el;
  if (depth === 0) return undefined;
  for (const c of elements(el)) {
    const found = descendant(c, ns, local, depth - 1);
    if (found) return found;
  }
  return undefined;
}

/** The `p15` namespace of comment threading (`p15:threadingInfo`). */
export { P15 };

/**
 * New `p:cm/@idx` values for copied comment lists. An index is unique per author in the
 * presentation (§19.4.1), so each copy takes the author's next free index; replies
 * (`p15:parentCm`) follow their parent, and `p:cmAuthor/@lastIdx` records the last one.
 */
export function renumberComments(
  batch: OpBatch,
  parts: readonly string[],
  presentation: string,
  options: RenumberOptions = {}
): void {
  const authorOf = options.authorOf ?? ((_part: string, el: Element) => attr(el, 'authorId'));
  // Per author: the indexes in use and the largest `lastIdx`. `p:cm/@idx` is `ST_Index` and
  // `p:cmAuthor/@lastIdx` is `xsd:unsignedInt` (`pml.xsd`); any other value counts for nothing.
  // `authorId` and `p:cmAuthor/@id` are `xsd:unsignedInt` too, so "007" and "7" are one author.
  const canonical = (value: string | undefined): string | undefined =>
    value === undefined ? undefined : String(unsignedIntId(value) ?? value);
  const used = new Map<string, Set<number>>();
  const after = new Map<string, number>();
  const bump = (raw: string | undefined, idx: string | undefined, inUse: boolean): void => {
    const author = canonical(raw);
    const n = unsignedIntId(idx);
    if (author === undefined || n === undefined) return;
    if (!inUse) {
      after.set(author, Math.max(after.get(author) ?? 0, n));
      return;
    }
    let own = used.get(author);
    if (!own) used.set(author, (own = new Set()));
    own.add(n);
  };
  // The lists being renumbered do not count: their indexes are about to change.
  const renumbered = new Set(parts);
  for (const name of batch.pkg.parts.keys()) {
    if (renumbered.has(name)) continue;
    const type = contentTypeOf(batch.pkg, name);
    if (type !== CT_COMMENTS && type !== CT_AUTHORS) continue;
    const root = treeOf(batch.pkg, name)?.root;
    if (root?.ns !== NS.p) continue;
    if (root.local === 'cmLst')
      for (const cm of elements(root)) bump(attr(cm, 'authorId'), attr(cm, 'idx'), true);
    if (root.local === 'cmAuthorLst')
      for (const a of elements(root)) bump(attr(a, 'id'), attr(a, 'lastIdx'), false);
  }
  // One past the author's largest index, then the lowest free ones (`freeIds`); in a
  // collaboration session, the free indexes of the actor's stripe (`stripedIds`).
  const sources = new Map<string, () => number>();
  const last = new Map<string, number>();
  const nextIdx = (author: string): number => {
    let source = sources.get(author);
    if (!source) {
      const own = used.get(author) ?? new Set<number>();
      const floor = after.get(author) ?? 0;
      source = batch.allocationActor
        ? // In a collaboration session the index comes from the actor's stripe, so two peers that
          // copy slides with comments at once do not give one author's comments one index. The
          // stripe starts after `lastIdx`, or lower when that leaves no number of the stripe.
          stripedIds(
            batch.allocationActor,
            own,
            Math.min(floor + 1, MAX_MINTED_ID - ACTOR_ID_STRIPE + 1),
            MAX_MINTED_ID,
            batch.mintedOf('commentIdx')
          )
        : freeIds(own, 1, MAX_MINTED_ID, floor);
      sources.set(author, source);
    }
    const idx = source();
    last.set(author, Math.max(last.get(author) ?? 0, idx));
    return idx;
  };
  const set = (part: string, node: Element, name: string, value: string): void =>
    batch.push({ op: 'setAttribute', part, node: node.id, name: { ns: '', local: name }, value });
  // A reply names its parent by author and index: the author id the element had before the copy
  // changed it (`authorOf`).
  const key = (part: string, el: Element): string =>
    `${canonical(authorOf(part, el)) ?? ''}:${canonical(attr(el, 'idx')) ?? ''}`;
  for (const part of parts) {
    const root = xmlPartOf(batch.pkg, part)?.root;
    // Keyed by the author and index each comment had before any change, so two authors that
    // now share one id cannot take each other's replies.
    const fresh = new Map<string, string>();
    for (const cm of elements(root)) {
      const author = canonical(attr(cm, 'authorId'));
      if (cm.local !== 'cm' || author === undefined) continue;
      const idx = nextIdx(author);
      fresh.set(key(part, cm), String(idx));
      set(part, cm, 'idx', String(idx));
    }
    for (const cm of elements(root)) {
      if (cm.local !== 'cm') continue;
      const extLst = elements(cm).find((e) => e.local === 'extLst');
      const exts = elements(extLst);
      for (const ext of exts) {
        const parent = descendant(ext, P15, 'parentCm', 4);
        if (!parent) continue;
        const idx = fresh.get(key(part, parent));
        if (idx) set(part, parent, 'idx', idx);
        else if (options.dropOrphanReplies && extLst)
          // The parent did not come along: the reply becomes a comment of its own.
          batch.push({
            op: 'removeNode',
            part,
            node: exts.length === 1 ? extLst.id : ext.id,
          });
      }
    }
  }
  const authors = relationshipsOf(batch.pkg, presentation).find(
    (r) => r.type === `${R}commentAuthors` && !r.external
  )?.target;
  if (!authors) return;
  const behind = (a: Element): number | undefined => {
    const n = last.get(canonical(attr(a, 'id')) ?? '');
    return n !== undefined && n > (unsignedIntId(attr(a, 'lastIdx')) ?? 0) ? n : undefined;
  };
  if (!elements(treeOf(batch.pkg, authors)?.root).some((a) => behind(a) !== undefined)) return;
  for (const a of elements(editable(batch, authors)?.root)) {
    const n = behind(a);
    if (n !== undefined) set(authors, a, 'lastIdx', String(n));
  }
}

/**
 * Give a copied slide its own `p14:creationId` (the `p14` extension schema; `val` is an
 * `xsd:unsignedInt`): two slides that share one look like one slide to co-authoring and comments.
 */
export function freshCreationId(batch: OpBatch, slide: string, slides: readonly string[]): void {
  freshCreationIdAmong(batch, slide, creationIdsOf(batch, slides));
}

/** The `p14:creationId` of a slide part, if it has one. */
function creationIdOf(batch: OpBatch, part: string): Element | undefined {
  const ext = elements(xmlPartOf(batch.pkg, part)?.root).find((e) => e.local === 'extLst');
  for (const e of elements(ext))
    for (const c of elements(e)) if (c.ns === NS.p14 && c.local === 'creationId') return c;
  return undefined;
}

/** The creation ids that `slides` use. */
export function creationIdsOf(batch: OpBatch, slides: readonly string[]): Set<number> {
  // `p14:creationId/@val` is `xsd:unsignedInt`; a value that is not one counts for nothing.
  const used = new Set<number>();
  for (const s of slides) {
    const n = unsignedIntId(attr(creationIdOf(batch, s), 'val'));
    if (n !== undefined) used.add(n);
  }
  return used;
}

/**
 * Give `slide` a creation id that no id in `used` has, and add it to `used`, so a run of new
 * slides checks each id once instead of every slide again.
 */
export function freshCreationIdAmong(
  batch: OpBatch,
  slide: string,
  used: Set<number>,
  /** Collect the op here instead of applying it, for a caller that applies many at once. */
  out?: PptxOp[]
): void {
  const own = creationIdOf(batch, slide);
  if (!own) return;
  let value: number;
  if (batch.allocationActor) {
    // In a collaboration session the id comes from the actor's stripe, so two peers that copy
    // one slide at once do not give both copies one id (`replica-ids.ts`).
    value = stripedIds(
      batch.allocationActor,
      used,
      1,
      MAX_UNSIGNED_INT,
      batch.mintedOf('creationId')
    )();
  } else {
    value = unsignedIntId(attr(own, 'val')) ?? 0;
    do value = value >= MAX_UNSIGNED_INT ? 1 : value + 1;
    while (used.has(value));
  }
  used.add(value);
  const op: PptxOp = {
    op: 'setAttribute',
    part: slide,
    node: own.id,
    name: { ns: '', local: 'val' },
    value: String(value),
  };
  if (out) out.push(op);
  else batch.push(op);
}

/** Internal parts reachable from `roots` through non-structural relationships. */
export function ownedParts(pkg: PptxPackage, roots: readonly string[]): Set<string> {
  const out = new Set<string>();
  const rootSet = new Set(roots);
  const visit = (source: string, depth: number): void => {
    if (depth > MAX_DEPTH) return;
    for (const rel of relationshipsOf(pkg, source)) {
      if (rel.external || !rel.target || STRUCTURAL.has(rel.type)) continue;
      if (out.has(rel.target) || rootSet.has(rel.target) || out.size >= MAX_PARTS) continue;
      out.add(rel.target);
      visit(rel.target, depth + 1);
    }
  };
  for (const root of roots) visit(root, 0);
  return out;
}

/**
 * Remove the candidates that no remaining part references any more (a deleted slide's
 * comments, charts, and media), with their relationships and content types. A part another
 * slide, layout, or master still names stays. When a relationship in the package cannot be
 * resolved, nothing is removed: that relationship could name a candidate.
 */
export function removeOrphans(batch: OpBatch, candidates: ReadonlySet<string>): void {
  for (let round = 0; round < MAX_DEPTH; round += 1) {
    const referenced = new Set<string>();
    for (const source of ['/', ...batch.pkg.parts.keys()]) {
      if (source.endsWith('.rels')) continue;
      for (const r of elements(xmlPartOf(batch.pkg, relsPartName(source))?.root)) {
        if (attr(r, 'TargetMode') === 'External') continue;
        const resolved = resolveInternalTarget(
          source === '/' ? '/x' : source,
          attr(r, 'Target') ?? ''
        );
        if (!resolved.ok) return;
        referenced.add(partNameKey(resolved.partName));
      }
    }
    const orphans = [...candidates].filter(
      (p) => batch.pkg.parts.has(p) && !referenced.has(partNameKey(p))
    );
    if (orphans.length === 0) return;
    for (const part of orphans) {
      batch.push({ op: 'removePart', part });
      if (batch.pkg.parts.has(relsPartName(part)))
        batch.push({ op: 'removePart', part: relsPartName(part) });
    }
    removeOverrides(batch, orphans);
  }
}

// ------------------------------------------------------------------------------- sections

/** Every `p14:sldId` of the section list, and the section `p14:sldIdLst` elements. */
function sections(batch: OpBatch, presentation: string): { ids: Element[]; lists: Element[] } {
  const ids: Element[] = [];
  const lists: Element[] = [];
  const walk = (el: Element, depth: number): void => {
    if (depth > 32) return;
    if (el.ns === NS.p14 && el.local === 'sldId') ids.push(el);
    if (el.ns === NS.p14 && el.local === 'sldIdLst') lists.push(el);
    for (const c of elements(el)) walk(c, depth + 1);
  };
  const root = xmlPartOf(batch.pkg, presentation)?.root;
  if (root) walk(root, 0);
  return { ids, lists };
}

/**
 * Whether a section entry names slide id `sldId`. Slide ids compare as numbers ("0258" is
 * 258, `ST_SlideId` is an `xsd:unsignedInt`), as the load and save repairs do.
 */
const namesSlide = (entry: Element, sldId: string): boolean =>
  canonicalSlideId(attr(entry, 'id') ?? '') === canonicalSlideId(sldId);

/** Remove a slide id from the section list (`p14:sectionLst`). */
export function dropFromSections(batch: OpBatch, presentation: string, sldId: string): void {
  dropSlidesFromSections(batch, presentation, [sldId]);
}

/** Remove slide ids from the section list, in one walk of it. */
export function dropSlidesFromSections(
  batch: OpBatch,
  presentation: string,
  sldIds: readonly string[]
): void {
  const gone = new Set(sldIds.map(canonicalSlideId));
  for (const el of sections(batch, presentation).ids)
    if (gone.has(canonicalSlideId(attr(el, 'id') ?? '')))
      batch.push({ op: 'removeNode', part: presentation, node: el.id });
}

/**
 * Put a slide id into the section of `neighbour`, right after or before it, so the section
 * list keeps naming every slide once and in slide order. A deck without sections is left as it
 * is. A neighbour no section names puts the slide at the end of the last section (or the start
 * of the first).
 */
/** Whether the presentation lists sections (`p14:sectionLst`). */
export function hasSections(batch: OpBatch, presentation: string): boolean {
  return sections(batch, presentation).lists.length > 0;
}

export function placeInSection(
  batch: OpBatch,
  presentation: string,
  sldId: string,
  neighbour: string,
  where: 'after' | 'before'
): void {
  placeRunInSection(batch, presentation, [sldId], neighbour, where);
}

/**
 * Put slide ids, in order, into the section of `neighbour`, right after or before it, as
 * `placeInSection` puts each one after the one before it. The section list is read once.
 */
function placeRunInSection(
  batch: OpBatch,
  presentation: string,
  run: readonly string[],
  neighbour: string,
  where: 'after' | 'before'
): void {
  dropSlidesFromSections(batch, presentation, run);
  const { ids, lists } = sections(batch, presentation);
  if (lists.length === 0) return;
  const anchor = ids.find((el) => namesSlide(el, neighbour));
  let list: Element;
  let index: number;
  if (anchor) {
    list = lists.find((l) => l.children.some((c) => c.id === anchor.id))!;
    const at = list.children.findIndex((c) => c.id === anchor.id);
    index = where === 'after' ? at + 1 : at;
  } else {
    list = where === 'after' ? lists[lists.length - 1] : lists[0];
    index = where === 'after' ? list.children.length : 0;
  }
  run.forEach((sldId, i) =>
    batch.push({
      op: 'insertNode',
      part: presentation,
      parent: list.id,
      index: index + i,
      spec: {
        ns: NS.p14,
        local: 'sldId',
        ...(list.prefix ? { prefix: list.prefix } : {}),
        attrs: [{ ns: '', local: 'id', value: sldId }],
      },
    })
  );
}

/**
 * Put the slides `block` (slide ids, in slide order, already at their place in `p:sldIdLst`)
 * into one section, in order. `before` and `after` are the slide ids next to the block. Without
 * `section`, the block joins the section of the slide before it (at the start of the deck, the
 * first section). With `section` (a `p14:section/@id`), it joins that section: a drop below a
 * section header, or a step across one, puts the block at the start or the end of the section,
 * as the slide pane does. That section must be the section of `before`, the section of
 * `after`, or an empty section between them; otherwise the command is refused.
 */
export function placeBlockInSection(
  batch: OpBatch,
  presentation: string,
  block: readonly string[],
  before: string | undefined,
  after: string | undefined,
  section?: string
): void {
  if (block.length === 0) return;
  const { lists } = sections(batch, presentation);
  if (section === undefined || lists.length === 0) {
    if (section !== undefined) throw new CommandRefusal('errors.unsupported', 'section');
    // A deck without sections has nothing to place.
    if (lists.length === 0) return;
    if (before) return placeRunInSection(batch, presentation, block, before, 'after');
    if (after) return placeRunInSection(batch, presentation, block, after, 'before');
    block.forEach((sldId, i) => {
      if (i > 0) placeInSection(batch, presentation, sldId, block[i - 1]!, 'after');
      else if (before) placeInSection(batch, presentation, sldId, before, 'after');
      else if (after) placeInSection(batch, presentation, sldId, after, 'before');
    });
    return;
  }
  // `lists` are the sections' `p14:sldIdLst` elements, in section order.
  const owner = (sldId: string | undefined): number =>
    sldId === undefined
      ? -1
      : lists.findIndex((l) => elements(l).some((e) => namesSlide(e, sldId)));
  const target = lists.findIndex((l) => {
    const el = sectionOfList(batch, presentation, l);
    return el !== undefined && (attr(el, 'id') ?? '') === section;
  });
  if (target < 0) throw new CommandRefusal('errors.unsupported', 'section');
  const low = owner(before);
  const high = after === undefined ? lists.length : owner(after);
  if (target < low || (high >= 0 && target > high))
    throw new CommandRefusal('errors.unsupported', 'section');
  if (target === low) return placeRunInSection(batch, presentation, block, before!, 'after');
  if (target === high) return placeRunInSection(batch, presentation, block, after!, 'before');
  block.forEach((sldId, i) => {
    if (i > 0) placeInSection(batch, presentation, sldId, block[i - 1]!, 'after');
    else {
      // An empty section between the two: the block becomes its slides.
      dropFromSections(batch, presentation, sldId);
      const list = sections(batch, presentation).lists[target]!;
      batch.push({
        op: 'insertNode',
        part: presentation,
        parent: list.id,
        index: 0,
        spec: {
          ns: NS.p14,
          local: 'sldId',
          ...(list.prefix ? { prefix: list.prefix } : {}),
          attrs: [{ ns: '', local: 'id', value: sldId }],
        },
      });
    }
  });
}

/** The `p14:section` that holds the section slide list `list`. */
function sectionOfList(batch: OpBatch, presentation: string, list: Element): Element | undefined {
  let found: Element | undefined;
  const walk = (el: Element, depth: number): void => {
    if (found || depth > 32) return;
    if (el.ns === NS.p14 && el.local === 'section' && el.children.includes(list)) found = el;
    else for (const c of elements(el)) walk(c, depth + 1);
  };
  const root = xmlPartOf(batch.pkg, presentation)?.root;
  if (root) walk(root, 0);
  return found;
}

/**
 * Keep sections in slide order: a block joins the section of the slide before it, or
 * `section` when given (`placeBlockInSection`). `moveSlide` and `moveSlides` share this rule.
 */
export function placeBlock(batch: OpBatch, block: readonly string[], section?: string): void {
  const info = presentationInfo(batch.pkg);
  const ids = info.slides.map((s) => s.nodeId);
  const byNode = new Map(info.slides.map((s) => [s.nodeId, s.sldId]));
  const sldIdOf = (nodeId: string) => byNode.get(nodeId) ?? '';
  const first = ids.indexOf(block[0]);
  const before = first > 0 ? ids[first - 1] : undefined;
  const after = ids[first + block.length];
  placeBlockInSection(
    batch,
    info.presentationPart,
    block.map(sldIdOf),
    before === undefined ? undefined : sldIdOf(before),
    after === undefined ? undefined : sldIdOf(after),
    section
  );
}

/**
 * Whether the slides `block` (slide node ids, in slide order) are already the consecutive
 * slides of section `section`, in that order: a move that keeps their place needs no ops.
 */
export function blockIsInSection(
  batch: OpBatch,
  block: readonly string[],
  section: string
): boolean {
  if (block.length === 0) return false;
  const info = presentationInfo(batch.pkg);
  const { lists } = sections(batch, info.presentationPart);
  const list = lists.find((l) => {
    const el = sectionOfList(batch, info.presentationPart, l);
    return el !== undefined && (attr(el, 'id') ?? '') === section;
  });
  if (!list) return false;
  const sldIds = block.map((nodeId) => {
    const id = info.slides.find((s) => s.nodeId === nodeId)?.sldId;
    return id === undefined ? undefined : canonicalSlideId(id);
  });
  const named = elements(list).map((e) => canonicalSlideId(attr(e, 'id') ?? ''));
  const start = named.indexOf(sldIds[0] ?? '');
  return start >= 0 && sldIds.every((id, i) => id !== undefined && named[start + i] === id);
}
