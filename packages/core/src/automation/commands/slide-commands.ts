// Slide commands: add, duplicate, delete, reorder, and the background. Each touches several
// parts (presentation, its relationships, content types, the slide, its notes) in ONE
// transaction. Speaker notes are in `notes-commands.ts`.

import { patternXml, pictureFillXml, pictureType } from './fill-commands.ts';
import {
  attr,
  child,
  CONTENT_TYPE,
  elements,
  find,
  NS,
  path,
  presentationInfo,
  relationshipsOf,
  relsPartName,
  REL,
  serializeXmlPart,
  shapeTreeOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
  type PresentationInfo,
  type SlideRef,
} from '../../store/index.ts';
import { partNameKey } from '@eigenpal/ooxml-core';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import {
  addOverride,
  addRelationship,
  addRelationships,
  freePartName,
  partNamer,
  removeOverrides,
  removeRelationships,
  stripedIds,
  unsignedIntId,
  xmlAttr,
  XML_DECL,
} from './package-edits.ts';
import { addImagePart, specFromXml } from './shape-commands.ts';
import {
  blockIsInSection,
  dropSlidesFromSections,
  creationIdsOf,
  freshCreationIdAmong,
  ownedParts,
  PartCopier,
  placeBlock,
  placeInSection,
  removeOrphans,
} from './slide-structure.ts';
import { fillXml, gradientXml } from './style-commands.ts';
import { writeCopiedTableMarkers } from './table-repair.ts';
import type { SlideBackground } from '../../contracts/neutral.ts';

const NS_DECL = `xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"`;
export const SP_TREE_HEAD =
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/>' +
  '<a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
export { NS_DECL };

/** `ST_SlideId`: slide ids run from 256 to 2147483647. */
const MIN_SLIDE_ID = 256;
const MAX_SLIDE_ID = 2_147_483_647;

export function slideRef(pkg: PptxPackage, index: number): SlideRef {
  const ref = presentationInfo(pkg).slides[index];
  if (!ref) throw new CommandRefusal('errors.unsupported', `no slide ${index}`);
  return ref;
}

/**
 * Slide ids that no slide uses, one for each call. In a collaboration session, the free ids of the actor's
 * stripe, so two peers that add a slide at once do not share one. Otherwise one above the
 * largest, so the id of a deleted slide does not come back and an old reference to it (an
 * automation handle) cannot reach the new slide; when the largest id is at the top of the
 * range, the smallest free id. Refuses when the range is full.
 */
function freeSlideIds(list: Element, batch: OpBatch): () => number {
  const used = new Set<number>();
  let max = MIN_SLIDE_ID - 1;
  for (const s of elements(list)) {
    const id = unsignedIntId(attr(s, 'id'), MIN_SLIDE_ID, MAX_SLIDE_ID);
    if (id === undefined) continue;
    used.add(id);
    if (id > max) max = id;
  }
  if (batch.allocationActor) {
    // A dense fallback could take a peer's id, so a full stripe refuses, as every mint does.
    return stripedIds(
      batch.allocationActor,
      used,
      MIN_SLIDE_ID,
      MAX_SLIDE_ID,
      batch.mintedOf('slideId')
    );
  }
  let low = MIN_SLIDE_ID;
  return () => {
    let id: number;
    if (max < MAX_SLIDE_ID) id = ++max;
    else {
      while (low <= MAX_SLIDE_ID && used.has(low)) low += 1;
      if (low > MAX_SLIDE_ID) throw new CommandRefusal('errors.tooLarge');
      id = low;
    }
    used.add(id);
    return id;
  };
}

/** `CT_Presentation` children that come after `p:sldIdLst` (ECMA-376 Part 1, §19.2.1.26). */
const AFTER_SLIDE_LIST: ReadonlySet<string> = new Set([
  'sldSz',
  'notesSz',
  'smartTags',
  'embeddedFontLst',
  'custShowLst',
  'photoAlbum',
  'custDataLst',
  'kinsoku',
  'defaultTextStyle',
  'modifyVerifier',
  'extLst',
]);

/**
 * A new, empty `p:sldIdLst` at its schema place, for a deck that has none: the element is
 * optional (`minOccurs="0"`) when the deck has no slides.
 */
function slideIdList(batch: OpBatch): Element | undefined {
  const presentation = presentationInfo(batch.pkg).presentationPart;
  const root = xmlPartOf(batch.pkg, presentation)?.root;
  if (!root) return undefined;
  const at = root.children.findIndex(
    (c) => c.type === 'element' && c.ns === NS.p && AFTER_SLIDE_LIST.has(c.local)
  );
  const id = batch.insert({
    op: 'insertNode',
    part: presentation,
    parent: root.id,
    index: at < 0 ? root.children.length : at,
    spec: { ns: NS.p, local: 'sldIdLst', prefix: 'p' },
  });
  return batch.element(presentation, id);
}

/**
 * Insert a `p:sldId` for a new slide part at `position` and relate it to the presentation.
 * Returns the new slide id.
 */
export function registerSlide(batch: OpBatch, partName: string, position: number): string {
  const presentation = presentationInfo(batch.pkg).presentationPart;
  const [nodeId] = registerSlides(batch, [partName], position);
  return attr(batch.element(presentation, nodeId!), 'id') ?? '';
}

/**
 * Insert a `p:sldId` for each new slide part, in order, at `position` (a count of the deck's
 * slides), and relate each to the presentation. The relationship ids and the slide ids are
 * chosen in one pass over the presentation. Returns the node ids of the new entries.
 */
export function registerSlides(
  batch: OpBatch,
  partNames: readonly string[],
  position: number
): string[] {
  const info = presentationInfo(batch.pkg);
  const presentation = info.presentationPart;
  const rIds = addRelationships(batch, presentation, REL.slide, partNames);
  for (const part of partNames) addOverride(batch, part, CONTENT_TYPE.slide);
  const list = path(xmlPartOf(batch.pkg, presentation)?.root, 'p:sldIdLst') ?? slideIdList(batch);
  if (!list) throw new CommandRefusal('errors.malformed');
  const nextId = freeSlideIds(list, batch);
  // `position` counts the deck's slides. `p:sldIdLst` can also hold an entry that names no
  // slide after concurrent edits, so the new entry goes before the entry of the slide at `position`.
  const entries = elements(list);
  const before = info.slides[position]?.nodeId;
  const at = before === undefined ? -1 : entries.findIndex((e) => e.id === before);
  const index = at >= 0 ? at : entries.length;
  return partNames.map((_, i) =>
    batch.insert({
      op: 'insertNode',
      part: presentation,
      parent: list.id,
      index: index + i,
      spec: {
        ns: NS.p,
        local: 'sldId',
        prefix: 'p',
        attrs: [
          { ns: '', local: 'id', value: String(nextId()) },
          { ns: NS.r, local: 'id', prefix: 'r', value: rIds[i]! },
        ],
      },
    })
  );
}

/** Footer placeholders: a new slide gets them only through Header and Footer. */
export const FOOTER_PLACEHOLDERS: ReadonlySet<string> = new Set(['dt', 'ftr', 'sldNum', 'hdr']);

/**
 * An empty slide placeholder for the layout placeholder `layoutEl`:
 * the same `p:ph` (ECMA-376 Part 1, §19.3.1.36), no position of its own, and one empty paragraph.
 */
export function emptyPlaceholderXml(layoutEl: Element, id: number): string {
  const ph = path(layoutEl, 'p:nvSpPr', 'p:nvPr', 'p:ph');
  const phAttrs = ['type', 'orient', 'sz', 'idx']
    .map((a) => {
      const value = attr(ph, a);
      return value !== undefined ? ` ${a}="${xmlAttr(value, 'placeholder')}"` : '';
    })
    .join('');
  const name = xmlAttr(
    attr(path(layoutEl, 'p:nvSpPr', 'p:cNvPr'), 'name') ?? `Placeholder ${id}`,
    'placeholder name'
  );
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${phAttrs}/></p:nvPr></p:nvSpPr>` +
    '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody></p:sp>'
  );
}

/**
 * Add a slide after `after`, with the content placeholders of its layout: `layoutPart` when
 * given, otherwise the layout of the slide before (`newSlideLayout` picks it).
 * It joins the section of the slide before it.
 */
export function addSlide(batch: OpBatch, after: number, layoutPart?: string): number {
  const info = presentationInfo(batch.pkg);
  const source = info.slides[Math.max(0, Math.min(after, info.slides.length - 1))];
  const sourceLayout = source?.layoutPart;
  if (layoutPart !== undefined && !xmlPartOf(batch.pkg, layoutPart))
    throw new CommandRefusal('errors.layoutNotFound');
  if (!layoutPart && !sourceLayout) throw new CommandRefusal('errors.malformed');
  layoutPart ??= sourceLayout!;
  let id = 1;
  const placeholders: string[] = [];
  for (const el of elements(shapeTreeOf(batch.pkg, layoutPart))) {
    const ph = path(el, 'p:nvSpPr', 'p:nvPr', 'p:ph');
    if (!ph || FOOTER_PLACEHOLDERS.has(attr(ph, 'type') ?? 'obj')) continue;
    id += 1;
    placeholders.push(emptyPlaceholderXml(el, id));
  }
  const partName = freePartName(batch, '/ppt/slides/slide', '.xml');
  batch.push({
    op: 'addPart',
    part: partName,
    xml: `${XML_DECL}<p:sld ${NS_DECL}><p:cSld><p:spTree>${SP_TREE_HEAD}${placeholders.join('')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
  });
  addRelationship(batch, partName, REL.slideLayout, layoutPart);
  const sldId = registerSlide(batch, partName, after + 1);
  if (source) placeInSection(batch, info.presentationPart, sldId, source.sldId, 'after');
  return after + 1;
}

/** What the copies of one command share: the deck before them, their names, their ids. */
interface CopyRun {
  readonly info: PresentationInfo;
  /** A free part name for `prefix<n>suffix`. */
  readonly name: (prefix: string, suffix: string) => string;
  /** The `p14:creationId` values in use. */
  readonly creationIds: Set<number>;
}

/**
 * Copy the parts of slide `ref` (the slide, its notes, and the parts it owns) under new names,
 * with no entry in the slide list yet. Returns the new slide part.
 */
function copySlideParts(batch: OpBatch, ref: SlideRef, run: CopyRun): string {
  const partName = run.name('/ppt/slides/slide', '.xml');
  const sourceXml = xmlPartOf(batch.pkg, ref.partName);
  if (!sourceXml) throw new CommandRefusal('errors.malformed');
  batch.push({ op: 'addPart', part: partName, xml: serializeXmlPart(sourceXml) });
  writeCopiedTableMarkers(batch, partName);
  let notesName: string | undefined;
  if (ref.notesPart && xmlPartOf(batch.pkg, ref.notesPart))
    notesName = run.name('/ppt/notesSlides/notesSlide', '.xml');
  const copier = new PartCopier(batch, run.info.presentationPart);
  batch.push({
    op: 'addPart',
    part: relsPartName(partName),
    xml: copier.relsXml(ref.partName, partName, (type) =>
      type === REL.notesSlide ? (notesName ? { part: notesName } : { drop: true }) : undefined
    ),
  });
  if (notesName && ref.notesPart) {
    batch.push({
      op: 'addPart',
      part: notesName,
      xml: serializeXmlPart(xmlPartOf(batch.pkg, ref.notesPart)!),
    });
    batch.push({
      op: 'addPart',
      part: relsPartName(notesName),
      xml: copier.relsXml(ref.notesPart, notesName, (type) =>
        type === REL.slide ? { part: partName } : undefined
      ),
    });
    addOverride(batch, notesName, CONTENT_TYPE.notesSlide);
  }
  copier.finish();
  freshCreationIdAmong(batch, partName, run.creationIds);
  return partName;
}

/**
 * Duplicate slide `index` (and its notes) right after it, in the same section. The copy gets
 * its own comments, charts, diagrams, and embedded objects, so editing one slide never
 * changes the other; pictures, media, and the layout stay shared. A slide holding parts the
 * engine cannot copy (threaded comments, embedded controls) is refused.
 */
export function duplicateSlide(batch: OpBatch, index: number): number {
  const info = presentationInfo(batch.pkg);
  const ref = slideRef(batch.pkg, index);
  const partName = copySlideParts(batch, ref, {
    info,
    name: (prefix, suffix) => freePartName(batch, prefix, suffix),
    creationIds: creationIdsOf(
      batch,
      info.slides.map((s) => s.partName)
    ),
  });
  const id = registerSlide(batch, partName, index + 1);
  placeInSection(batch, info.presentationPart, id, ref.sldId, 'after');
  return index + 1;
}

/**
 * Duplicate the slides at `indices` (ascending), as `duplicateSlide` copies each one, and put
 * the copies in order right after the last of them, in its section. The deck is read once and
 * the slide list written once for all the copies, so a copy of every slide of a long deck
 * costs what it copies. Returns the `p:sldId` node ids of the copies, in order.
 */
export function duplicateSlidesAt(batch: OpBatch, indices: readonly number[]): string[] {
  const info = presentationInfo(batch.pkg);
  const refs = indices.map((i) => slideRef(batch.pkg, i));
  const run: CopyRun = {
    info,
    name: partNamer(batch),
    creationIds: creationIdsOf(
      batch,
      info.slides.map((s) => s.partName)
    ),
  };
  const parts = refs.map((ref) => copySlideParts(batch, ref, run));
  const copies = registerSlides(batch, parts, indices[indices.length - 1]! + 1);
  placeBlock(batch, copies);
  return copies;
}

/** Elements that reference a relationship by id: hyperlink actions and custom show entries. */
const HYPERLINKS = new Set(['hlinkClick', 'hlinkMouseOver', 'hlinkHover']);

/** Elements of a part with an `r:*` attribute equal to `rId`. */
function referencesTo(root: Element, rId: string): Element[] {
  const out: Element[] = [];
  const walk = (el: Element, depth: number): void => {
    if (depth > 256) return;
    if (el.attrs.some((a) => a.ns === NS.r && a.value === rId)) out.push(el);
    for (const c of elements(el)) walk(c, depth + 1);
  };
  walk(root, 0);
  return out;
}

/**
 * Remove every relationship from a remaining part to a removed one, with what uses it. A slide
 * jump (`a:hlinkClick` to a deleted slide) loses its action, and a custom show loses the slide,
 * A relationship that something else still uses is left alone.
 */
function dropReferences(batch: OpBatch, removed: ReadonlySet<string>, presentation: string): void {
  for (const source of [...batch.pkg.parts.keys()]) {
    if (removed.has(partNameKey(source)) || source.endsWith('.rels')) continue;
    const xml = xmlPartOf(batch.pkg, source);
    const dropped: string[] = [];
    for (const rel of relationshipsOf(batch.pkg, source)) {
      if (rel.external || !rel.target || !removed.has(partNameKey(rel.target))) continue;
      const users = xml ? referencesTo(xml.root, rel.id) : [];
      const known = users.every(
        (u) =>
          (u.ns === NS.a && HYPERLINKS.has(u.local)) ||
          (u.ns === NS.p && u.local === 'sld' && source === presentation)
      );
      if (!known) continue;
      for (const u of users) batch.push({ op: 'removeNode', part: source, node: u.id });
      dropped.push(rel.id);
    }
    if (dropped.length) removeRelationships(batch, source, dropped);
  }
}

/**
 * Delete slide `index`, its relationships, content type, and notes, and every reference other
 * parts make to them. The parts only it used (comments, charts, media) go too. The last slide
 * stays.
 */
export function deleteSlide(batch: OpBatch, index: number): void {
  deleteSlidesAt(batch, [index]);
}

/**
 * Delete the slides at `indices` (distinct), as `deleteSlide` deletes each one. Each pass over
 * the package (references, content types, orphans) runs once for all of them, so deleting most
 * slides of a long deck costs one pass, not one for each slide. At least one slide stays.
 */
export function deleteSlidesAt(batch: OpBatch, indices: readonly number[]): void {
  const info = presentationInfo(batch.pkg);
  if (info.slides.length - new Set(indices).size < 1)
    throw new CommandRefusal('disabled.lastSlide');
  const refs = indices.map((i) => slideRef(batch.pkg, i));
  for (const ref of refs)
    batch.push({ op: 'removeNode', part: info.presentationPart, node: ref.nodeId });
  const sldIds = refs.map((r) => r.sldId).filter((id) => id !== '');
  if (sldIds.length) dropSlidesFromSections(batch, info.presentationPart, sldIds);
  const parts = refs
    .flatMap((ref) => [ref.partName, ref.notesPart])
    .filter((p): p is string => !!p && batch.pkg.parts.has(p));
  const owned = ownedParts(batch.pkg, parts);
  // Custom shows name the slide through the presentation's relationship, so it goes last.
  dropReferences(batch, new Set(parts.map(partNameKey)), info.presentationPart);
  removeRelationships(
    batch,
    info.presentationPart,
    refs.map((r) => r.rId)
  );
  for (const part of parts) {
    batch.push({ op: 'removePart', part });
    if (batch.pkg.parts.has(relsPartName(part)))
      batch.push({ op: 'removePart', part: relsPartName(part) });
  }
  removeOverrides(batch, parts);
  removeOrphans(batch, owned);
}

/**
 * Move slide `from` to position `to`. It joins the section of the slide before it, or
 * `section`, by the rule of `moveSlides` (`placeBlock`), so the two commands agree. Behavior
 * change: a slide moved up to the first place of a section joined that section before (the
 * section of the slide after it); now it joins the section before it unless `section` names
 * that section. A move that changes nothing emits no ops, so it adds no undo step.
 */
export function moveSlide(batch: OpBatch, from: number, to: number, section?: string): void {
  const info = presentationInfo(batch.pkg);
  const ref = slideRef(batch.pkg, from);
  const list = path(xmlPartOf(batch.pkg, info.presentationPart)?.root, 'p:sldIdLst');
  if (!list) throw new CommandRefusal('errors.malformed');
  const target = Math.max(0, Math.min(to, info.slides.length - 1));
  const anchor = info.slides[target];
  if (target === from && (section === undefined || blockIsInSection(batch, [ref.nodeId], section)))
    return;
  const index = list.children.findIndex((c) => c.id === anchor.nodeId);
  batch.push({
    op: 'moveNode',
    part: info.presentationPart,
    node: ref.nodeId,
    parent: list.id,
    index,
  });
  placeBlock(batch, [ref.nodeId], section);
}

export { find, relationshipsOf };

function pictureBackground(
  batch: OpBatch,
  part: string,
  p: Extract<SlideBackground, { kind: 'picture' }>
): string {
  const type = pictureType(p.bytes);
  return pictureFillXml(addImagePart(batch, part, p.bytes, type), p.tile === true);
}

/** Set or remove the background of a slide (`p:cSld/p:bg`). */
export function setSlideBackground(
  batch: OpBatch,
  index: number,
  background: SlideBackground | null
): void {
  const part = slideRef(batch.pkg, index).partName;
  const cSld = find(xmlPartOf(batch.pkg, part)?.root, 'p:cSld');
  if (!cSld) throw new CommandRefusal('errors.malformed');
  const old = child(cSld, 'p:bg');
  if (old) batch.push({ op: 'removeNode', part, node: old.id });
  if (!background) return;
  const fill =
    background.kind === 'solid'
      ? fillXml(background.color)
      : background.kind === 'gradient'
        ? gradientXml(background.gradient)
        : background.kind === 'pattern'
          ? patternXml(background.pattern)
          : background.kind === 'picture'
            ? pictureBackground(batch, part, background)
            : null;
  if (!fill) throw new CommandRefusal('errors.unsupported', 'background');
  // `p:bg` is the first child of `p:cSld`.
  batch.insert({
    op: 'insertNode',
    part,
    parent: find(xmlPartOf(batch.pkg, part)?.root, 'p:cSld')!.id,
    index: 0,
    spec: specFromXml(`<p:bg><p:bgPr>${fill}<a:effectLst/></p:bgPr></p:bg>`),
  });
}
