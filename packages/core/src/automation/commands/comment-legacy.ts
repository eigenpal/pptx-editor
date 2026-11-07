// Legacy comments (ECMA-376 Part 1, 19.4): the author list (`p:cmAuthorLst`, the presentation's
// `commentAuthors` part) and one comment list per slide (`p:cmLst`). A comment is `p:cm authorId
// dt idx` with `p:pos` (1/576 inch) and `p:text`. An index is unique per author in the
// presentation, and `p:cmAuthor/@lastIdx` records the author's last one; it never goes down. A
// reply names its parent in `p15:threadingInfo/p15:parentCm` (the `p15` extension). The format
// has no resolved state.

import {
  attr,
  child,
  elements,
  NS,
  presentationInfo,
  xmlPartOf,
  type Element,
} from '../../store/index.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import {
  addOverride,
  addRelationship,
  freeIds,
  freePartName,
  MAX_MINTED_ID,
  stripedIds,
  unsignedIntId,
  xmlAttr,
  xmlText,
  XML_DECL,
} from './package-edits.ts';
import { ACTOR_ID_STRIPE } from './actor-scoped-ids.ts';
import { editable } from './slide-structure.ts';
import {
  canonicalAuthor,
  canonicalIdx,
  commentTree,
  COMMENTS,
  descendant,
  findCommentAuthor,
  legacyAuthors,
  legacyKey,
  legacyKeyFields,
  LEGACY_POS_EMU,
  presentationTarget,
  slideCommentParts,
} from './comment-read.ts';
import { fragment, freeName } from './comment-write.ts';

const isP = (el: Element | undefined, local: string): el is Element =>
  el?.ns === NS.p && el.local === local;

/** The author list, created when the deck has none. */
function authorList(batch: OpBatch): { part: string; root: Element } {
  let part = presentationTarget(batch.pkg, COMMENTS.legacyAuthorsRel);
  if (!part) {
    part = freeName(batch, '/ppt/commentAuthors.xml', '/ppt/commentAuthors', '.xml');
    batch.push({ op: 'addPart', part, xml: `${XML_DECL}<p:cmAuthorLst xmlns:p="${NS.p}"/>` });
    addOverride(batch, part, COMMENTS.legacyAuthorsType);
    const presentation = presentationInfo(batch.pkg).presentationPart;
    addRelationship(batch, presentation, COMMENTS.legacyAuthorsRel, part);
  }
  const root = editable(batch, part)?.root;
  if (!isP(root, 'cmAuthorLst')) throw new CommandRefusal('errors.malformed', 'commentAuthors');
  return { part, root };
}

/** The id of the author `name` with `initials`, added to the list when it is new. */
function authorId(
  batch: OpBatch,
  author: { readonly name: string; readonly initials: string }
): string {
  const known = unsignedIntId(findCommentAuthor(batch.pkg, 'legacy', author));
  if (known !== undefined) return String(known);
  const { part, root } = authorList(batch);
  // Ids stay unique over every list, as a save joins the lists (`comment-repair.ts`).
  const used = new Set([...legacyAuthors(batch.pkg).keys()].map(Number));
  let colors = -1;
  for (const a of elements(root)) {
    if (!isP(a, 'cmAuthor')) continue;
    const id = unsignedIntId(attr(a, 'id'));
    if (id === undefined) continue;
    used.add(id);
    colors = Math.max(colors, unsignedIntId(attr(a, 'clrIdx')) ?? 0);
  }
  // In a collaboration session the id comes from the actor's stripe, so two peers who add a
  // new author at once do not give two names one id.
  const id = batch.allocationActor
    ? stripedIds(batch.allocationActor, used, 0, MAX_MINTED_ID, batch.mintedOf('commentAuthor'))()
    : freeIds(used, 0, MAX_MINTED_ID)();
  // Attributes in name order, as a collaboration room writes back a part kept as bytes.
  batch.push({
    op: 'insertNode',
    part,
    parent: root.id,
    index: root.children.length,
    spec: fragment(
      `<p:cmAuthor clrIdx="${Math.min(colors + 1, MAX_MINTED_ID)}" id="${id}" ` +
        `initials="${xmlAttr(author.initials, 'initials')}" lastIdx="0" ` +
        `name="${xmlAttr(author.name, 'author')}"/>`
    ),
  });
  return String(id);
}

/** A new index for author `id`, recorded in its `lastIdx`. */
function nextIndex(batch: OpBatch, id: string): number {
  const used = new Set<number>();
  for (const s of presentationInfo(batch.pkg).slides)
    for (const list of slideCommentParts(batch.pkg, s.partName).legacy)
      for (const cm of elements(commentTree(batch.pkg, list)?.root))
        if (isP(cm, 'cm') && canonicalAuthor(attr(cm, 'authorId')) === id) {
          const n = unsignedIntId(attr(cm, 'idx'));
          if (n !== undefined) used.add(n);
        }
  const { part, root } = authorList(batch);
  const author = elements(root).find(
    (a) => isP(a, 'cmAuthor') && canonicalAuthor(attr(a, 'id')) === id
  );
  if (!author) throw new CommandRefusal('errors.malformed', 'commentAuthors');
  const floor = unsignedIntId(attr(author, 'lastIdx')) ?? 0;
  const idx = batch.allocationActor
    ? stripedIds(
        batch.allocationActor,
        used,
        Math.min(floor + 1, MAX_MINTED_ID - ACTOR_ID_STRIPE + 1),
        MAX_MINTED_ID,
        batch.mintedOf('commentIdx')
      )()
    : freeIds(used, 1, MAX_MINTED_ID, floor)();
  if (idx > floor)
    batch.push({
      op: 'setAttribute',
      part,
      node: author.id,
      name: { ns: '', local: 'lastIdx' },
      value: String(idx),
    });
  return idx;
}

/** The slide's legacy comment list, created with its relationship when it has none. */
function listOf(batch: OpBatch, slidePart: string): { part: string; root: Element } {
  let part = slideCommentParts(batch.pkg, slidePart).legacy[0];
  if (!part) {
    part = freePartName(batch, '/ppt/comments/comment', '.xml');
    batch.push({ op: 'addPart', part, xml: `${XML_DECL}<p:cmLst xmlns:p="${NS.p}"/>` });
    addOverride(batch, part, COMMENTS.legacyType);
    addRelationship(batch, slidePart, COMMENTS.legacyRel, part);
  }
  const root = editable(batch, part)?.root;
  if (!isP(root, 'cmLst')) throw new CommandRefusal('errors.malformed', 'comments');
  return { part, root };
}

const commentXml = (
  author: string,
  date: string,
  idx: number,
  pos: { readonly x: string; readonly y: string },
  text: string,
  parent?: { readonly author: string; readonly idx: string }
): string =>
  `<p:cm authorId="${author}" dt="${date}" idx="${idx}"><p:pos x="${pos.x}" y="${pos.y}"/>` +
  `<p:text>${xmlText(text, 'comment')}</p:text>` +
  (parent
    ? `<p:extLst><p:ext uri="${COMMENTS.threadingUri}"><p15:threadingInfo xmlns:p15="${NS.p15}" ` +
      `timeZoneBias="0"><p15:parentCm authorId="${xmlAttr(parent.author, 'author id')}" ` +
      `idx="${xmlAttr(parent.idx, 'index')}"/></p15:threadingInfo></p:ext></p:extLst>`
    : '') +
  '</p:cm>';

/** The `p:pos` of a new thread: the top right corner of element `nodeId`, else the corner. */
function positionOf(batch: OpBatch, slidePart: string, nodeId: string | undefined) {
  const tree = child(child(xmlPartOf(batch.pkg, slidePart)?.root, 'p:cSld'), 'p:spTree');
  const el = nodeId ? elements(tree).find((e) => e.id === nodeId) : undefined;
  const xfrm = el
    ? (descendant(el, NS.a, 'xfrm', 2) ?? descendant(el, NS.p, 'xfrm', 1))
    : undefined;
  const off = child(xfrm, 'a:off');
  const ext = child(xfrm, 'a:ext');
  const x = Number(attr(off, 'x')) + Number(attr(ext, 'cx'));
  const y = Number(attr(off, 'y'));
  if (!Number.isFinite(x) || !Number.isFinite(y) || x < 0 || y < 0) return { x: '10', y: '10' };
  return { x: String(Math.round(x / LEGACY_POS_EMU)), y: String(Math.round(y / LEGACY_POS_EMU)) };
}

/** Add a legacy comment on `slidePart`, placed at `nodeId` when given. Returns its id key. */
export function addLegacyComment(
  batch: OpBatch,
  slidePart: string,
  text: string,
  author: { readonly name: string; readonly initials: string },
  date: string,
  nodeId: string | undefined
): string {
  const id = authorId(batch, author);
  const idx = nextIndex(batch, id);
  const pos = positionOf(batch, slidePart, nodeId);
  const { part, root } = listOf(batch, slidePart);
  batch.push({
    op: 'insertNode',
    part,
    parent: root.id,
    index: root.children.length,
    spec: fragment(commentXml(id, date, idx, pos, text)),
  });
  return legacyKey(id, String(idx));
}

/** The parent key (`author:idx`) of a legacy reply, or undefined for a thread. */
function parentKey(cm: Element): string | undefined {
  const parent = descendant(child(cm, 'p:extLst'), NS.p15, 'parentCm', 4);
  return parent
    ? `${canonicalAuthor(attr(parent, 'authorId'))}:${canonicalIdx(attr(parent, 'idx'))}`
    : undefined;
}

const keyOf = (cm: Element): string =>
  `${canonicalAuthor(attr(cm, 'authorId'))}:${canonicalIdx(attr(cm, 'idx'))}`;

/**
 * A legacy comment of a slide by its id key (`cm-<author>-<idx>`, each field encoded, and
 * `-<n>` for the n-th repeat of an author and index): its list, the comment, and the thread it
 * is in (itself for a thread). A key without a place names a comment only while the slide holds
 * its author and index once.
 */
export function findLegacy(
  batch: OpBatch,
  slidePart: string,
  key: string
): { part: string; cm: Element; thread: Element; placed: boolean; group: number } {
  const fields = legacyKeyFields(key);
  if (!fields) throw new CommandRefusal('errors.noComment');
  const wanted = `${fields.authorId}:${fields.idx}`;
  const matches = (root: Element | undefined): Element[] =>
    isP(root, 'cmLst') ? elements(root).filter((cm) => isP(cm, 'cm') && keyOf(cm) === wanted) : [];
  // Read first, so a list without the comment is not rewritten as an XML part.
  const lists = slideCommentParts(batch.pkg, slidePart).legacy.map((part) => ({
    part,
    hits: matches(commentTree(batch.pkg, part)?.root).length,
  }));
  const total = lists.reduce((n, l) => n + l.hits, 0);
  if (fields.ordinal === undefined && total !== 1) throw new CommandRefusal('errors.noComment');
  let left = fields.ordinal ?? 1;
  for (const { part, hits } of lists) {
    if (left > hits) {
      left -= hits;
      continue;
    }
    const root = editable(batch, part)?.root;
    const cm = matches(root)[left - 1];
    if (!cm) break;
    const all = elements(root).filter((c) => isP(c, 'cm'));
    const parent = parentKey(cm);
    const thread = (parent && all.find((c) => keyOf(c) === parent && !parentKey(c))) || cm;
    return { part, cm, thread, placed: fields.ordinal !== undefined, group: total };
  }
  throw new CommandRefusal('errors.noComment');
}

/** Reply to `thread` of list `part`, at its place. Returns the reply's id key. */
export function replyLegacy(
  batch: OpBatch,
  part: string,
  thread: Element,
  text: string,
  author: { readonly name: string; readonly initials: string },
  date: string
): string {
  const id = authorId(batch, author);
  const idx = nextIndex(batch, id);
  const pos = child(thread, 'p:pos');
  const root = xmlPartOf(batch.pkg, part)!.root;
  const parent = {
    author: canonicalAuthor(attr(thread, 'authorId')),
    idx: canonicalIdx(attr(thread, 'idx')),
  };
  // After the thread's last reply, so file order keeps the conversation together.
  const thread0 = keyOf(thread);
  const kids = root.children;
  let at = kids.findIndex((c) => c.type === 'element' && c.id === thread.id);
  for (let i = at + 1; i < kids.length; i += 1) {
    const c = kids[i]!;
    if (c.type === 'element' && isP(c, 'cm') && parentKey(c) === thread0) at = i;
  }
  batch.push({
    op: 'insertNode',
    part,
    parent: root.id,
    index: at < 0 ? kids.length : at + 1,
    spec: fragment(
      commentXml(
        id,
        date,
        idx,
        {
          x: xmlAttr(attr(pos, 'x') ?? '10', 'position'),
          y: xmlAttr(attr(pos, 'y') ?? '10', 'position'),
        },
        text,
        parent
      )
    ),
  });
  return legacyKey(id, String(idx));
}

/** Replace the text of a legacy comment. */
export function editLegacy(batch: OpBatch, part: string, cm: Element, text: string): void {
  const node = child(cm, 'p:text');
  if (!node) throw new CommandRefusal('errors.malformed', 'comment');
  batch.push({ op: 'setText', part, node: node.id, text });
}

/**
 * Delete a reply, or a thread with its replies. An empty list stays while the deck is open, so a
 * comment that a collaborator adds at the same time is kept; the save drops empty lists.
 */
export function deleteLegacy(batch: OpBatch, part: string, cm: Element): void {
  const own = keyOf(cm);
  const thread = !parentKey(cm);
  const root = xmlPartOf(batch.pkg, part)!.root;
  for (const c of elements(root))
    if (c.id === cm.id || (thread && isP(c, 'cm') && parentKey(c) === own))
      batch.push({ op: 'removeNode', part, node: c.id });
}
