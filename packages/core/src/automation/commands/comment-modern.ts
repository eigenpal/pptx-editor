// Threaded comments (the `p188` extension): the author list (`p188:authorLst`, the presentation's
// `authors` part), one comment list per slide (`p188:cmLst`) that the slide names by
// relationship and by `p188:commentRel`, and the comments, replies, text, and status in it.
// A comment points at its slide through a moniker list: `pc:sldMkLst` names the slide by its
// `p14:creationId` and `p:sldId/@id`, and `ac:deMkLst` adds a drawing element of the slide by its
// `p:cNvPr/@id` and `a16:creationId`.

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
  MAX_UNSIGNED_INT,
  stripedIds,
  unsignedIntId,
  xmlAttr,
  xmlText,
  XML_DECL,
} from './package-edits.ts';
import { editable } from './slide-structure.ts';
import {
  COMMENTS,
  commentTree,
  modernText,
  findCommentAuthor,
  modernKeyParts,
  modernMark,
  modernTextEditable,
  presentationTarget,
  slideCommentParts,
} from './comment-read.ts';
import { fragment, freeName, newGuid, slideExtList } from './comment-write.ts';

const P188 = COMMENTS.p188;
const DECLS = `xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p188="${P188}"`;
const CREATION_URI = '{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}';

/** The moniker element of each drawing element kind. */
const MONIKERS: Readonly<Record<string, string>> = {
  sp: 'spMk',
  pic: 'picMk',
  cxnSp: 'cxnSpMk',
  grpSp: 'grpSpMk',
  graphicFrame: 'graphicFrameMk',
};

const is = (el: Element | undefined, local: string): el is Element =>
  el?.ns === P188 && el.local === local;

/** The id of the author `name` with `initials`, added to the author list when it is new. */
export function modernAuthorId(
  batch: OpBatch,
  author: { readonly name: string; readonly initials: string }
): string {
  const known = findCommentAuthor(batch.pkg, 'modern', author);
  if (known) return known;
  let part = presentationTarget(batch.pkg, COMMENTS.modernAuthorsRel);
  if (!part) {
    part = freeName(batch, '/ppt/authors.xml', '/ppt/authors', '.xml');
    batch.push({ op: 'addPart', part, xml: `${XML_DECL}<p188:authorLst ${DECLS}/>` });
    addOverride(batch, part, COMMENTS.modernAuthorsType);
    const presentation = presentationInfo(batch.pkg).presentationPart;
    addRelationship(batch, presentation, COMMENTS.modernAuthorsRel, part);
  }
  const list = editable(batch, part)?.root;
  if (!is(list, 'authorLst')) throw new CommandRefusal('errors.malformed', 'authors');
  const id = newGuid();
  const name = xmlAttr(author.name, 'author');
  // Attributes in name order, as a collaboration room writes back a part kept as bytes.
  batch.push({
    op: 'insertNode',
    part,
    parent: list.id,
    index: list.children.length,
    spec: fragment(
      `<p188:author id="${id}" initials="${xmlAttr(author.initials, 'initials')}" ` +
        `name="${name}" providerId="None" userId="${name}"/>`
    ),
  });
  return id;
}

/** The slide's `p14:creationId`, added when the slide has none: the moniker names it. */
function slideCreationId(batch: OpBatch, slidePart: string): string {
  const read = (part: string): string | undefined => {
    const ext = child(xmlPartOf(batch.pkg, part)?.root, 'p:extLst');
    for (const e of elements(ext))
      for (const c of elements(e))
        if (c.ns === NS.p14 && c.local === 'creationId') return attr(c, 'val');
    return undefined;
  };
  const own = unsignedIntId(read(slidePart), 1);
  if (own !== undefined) return String(own);
  const used = new Set<number>();
  for (const s of presentationInfo(batch.pkg).slides) {
    const n = unsignedIntId(read(s.partName));
    if (n !== undefined) used.add(n);
  }
  const value = batch.allocationActor
    ? stripedIds(batch.allocationActor, used, 1, MAX_UNSIGNED_INT, batch.mintedOf('creationId'))()
    : freeIds(used, 1, MAX_UNSIGNED_INT)();
  const extLst = slideExtList(batch, slidePart);
  batch.push({
    op: 'insertNode',
    part: slidePart,
    parent: extLst.id,
    index: 0,
    spec: fragment(
      `<p:ext uri="${CREATION_URI}"><p14:creationId xmlns:p14="${NS.p14}" val="${value}"/></p:ext>`
    ),
  });
  return String(value);
}

/** The slide's threaded comment list, created with its relationship when it has none. */
function listOf(batch: OpBatch, slidePart: string): { part: string; root: Element } {
  const existing = slideCommentParts(batch.pkg, slidePart).modern[0];
  if (existing) {
    const root = editable(batch, existing)?.root;
    if (!is(root, 'cmLst')) throw new CommandRefusal('errors.malformed', 'comments');
    return { part: existing, root };
  }
  const ref = presentationInfo(batch.pkg).slides.find((s) => s.partName === slidePart)!;
  const creation = slideCreationId(batch, slidePart);
  const hex = (n: string) => Number(n).toString(16).toUpperCase();
  const part = freeName(
    batch,
    `/ppt/comments/modernComment_${hex(ref.sldId)}_${hex(creation)}.xml`,
    '/ppt/comments/modernComment',
    '.xml'
  );
  batch.push({ op: 'addPart', part, xml: `${XML_DECL}<p188:cmLst ${DECLS}/>` });
  addOverride(batch, part, COMMENTS.modernType);
  const rId = addRelationship(batch, slidePart, COMMENTS.modernRel, part);
  const extLst = slideExtList(batch, slidePart);
  batch.push({
    op: 'insertNode',
    part: slidePart,
    parent: extLst.id,
    index: elements(extLst).length,
    spec: fragment(
      `<p:ext uri="${COMMENTS.commentRelUri}">` +
        `<p188:commentRel xmlns:p188="${P188}" r:id="${xmlAttr(rId, 'relationship')}"/></p:ext>`
    ),
  });
  return { part, root: xmlPartOf(batch.pkg, part)!.root };
}

/** `p188:txBody` with one paragraph per line. */
function bodyXml(text: string): string {
  return `<p188:txBody><a:bodyPr/><a:lstStyle/>${paragraphsXml(text)}</p188:txBody>`;
}

const paragraphsXml = (text: string): string =>
  text
    .split('\n')
    .map((line) =>
      line ? `<a:p><a:r><a:t>${xmlText(line, 'comment')}</a:t></a:r></a:p>` : '<a:p/>'
    )
    .join('');

/** The moniker list of a new comment: the slide, and the element `target` when given. */
function monikerXml(
  batch: OpBatch,
  slidePart: string,
  creation: string,
  target: { readonly kind: string; readonly id: string; readonly creationId: string } | undefined
): string {
  const ref = presentationInfo(batch.pkg).slides.find((s) => s.partName === slidePart)!;
  const slide = `<pc:docMk/><pc:sldMk cId="${creation}" sldId="${xmlAttr(ref.sldId, 'slide id')}"/>`;
  if (!target) return `<pc:sldMkLst xmlns:pc="${COMMENTS.pc}">${slide}</pc:sldMkLst>`;
  return (
    `<ac:deMkLst xmlns:ac="${COMMENTS.ac}" xmlns:pc="${COMMENTS.pc}">${slide}` +
    // Attributes in name order, as a collaboration room writes back a part kept as bytes.
    `<ac:${MONIKERS[target.kind]} creationId="${target.creationId}" ` +
    `id="${xmlAttr(target.id, 'shape id')}"/></ac:deMkLst>`
  );
}

/**
 * The drawing element `nodeId` of the slide, as a moniker names it: its kind, `p:cNvPr/@id`, and
 * `a16:creationId`, which is added when the element has none. Only top-level elements.
 */
function elementTarget(
  batch: OpBatch,
  slidePart: string,
  nodeId: string
): { kind: string; id: string; creationId: string } | undefined {
  const root = xmlPartOf(batch.pkg, slidePart)?.root;
  const tree = child(child(root, 'p:cSld'), 'p:spTree');
  const el = elements(tree).find((e) => e.id === nodeId);
  if (!el || !MONIKERS[el.local]) return undefined;
  const nv = elements(el).find((c) => c.ns === NS.p && c.local.startsWith('nv'));
  const cNvPr = child(nv, 'p:cNvPr');
  const id = attr(cNvPr, 'id');
  if (!cNvPr || id === undefined || unsignedIntId(id) === undefined) return undefined;
  const extLst = child(cNvPr, 'a:extLst');
  for (const ext of elements(extLst))
    for (const c of elements(ext))
      if (
        c.ns === NS.a16 &&
        c.local === 'creationId' &&
        /^\{[0-9A-Fa-f-]{36}\}$/.test(attr(c, 'id') ?? '')
      )
        return { kind: el.local, id, creationId: attr(c, 'id')!.toUpperCase() };
  const creationId = newGuid();
  const ext =
    `<a:ext uri="${COMMENTS.creationIdUri}">` +
    `<a16:creationId xmlns:a16="${NS.a16}" id="${creationId}"/></a:ext>`;
  if (extLst)
    batch.push({
      op: 'insertNode',
      part: slidePart,
      parent: extLst.id,
      index: extLst.children.length,
      spec: fragment(ext),
    });
  else
    batch.push({
      op: 'insertNode',
      part: slidePart,
      parent: cNvPr.id,
      index: cNvPr.children.length,
      spec: fragment(`<a:extLst>${ext}</a:extLst>`),
    });
  return { kind: el.local, id, creationId };
}

/** Add a threaded comment on `slidePart`, pointing at `nodeId` when given. Returns its GUID. */
export function addModernComment(
  batch: OpBatch,
  slidePart: string,
  text: string,
  author: { readonly name: string; readonly initials: string },
  date: string,
  nodeId: string | undefined
): string {
  const authorId = modernAuthorId(batch, author);
  const creation = slideCreationId(batch, slidePart);
  const target = nodeId ? elementTarget(batch, slidePart, nodeId) : undefined;
  const { part, root } = listOf(batch, slidePart);
  const id = newGuid();
  batch.push({
    op: 'insertNode',
    part,
    parent: root.id,
    index: root.children.length,
    spec: fragment(
      `<p188:cm authorId="${authorId}" created="${date}" id="${id}">` +
        monikerXml(batch, slidePart, creation, target) +
        `${bodyXml(text)}</p188:cm>`
    ),
  });
  return id;
}

type ModernHit = { thread: Element; reply?: Element };

/** Every threaded comment and reply of a list with GUID `wanted`, each thread before its replies. */
function modernHits(root: Element | undefined, wanted: string): ModernHit[] {
  if (!is(root, 'cmLst')) return [];
  const out: ModernHit[] = [];
  for (const cm of elements(root)) {
    // The reader skips a thread without an id, with its replies: so do the ordinals.
    if (!is(cm, 'cm') || !attr(cm, 'id')) continue;
    if (attr(cm, 'id')?.toUpperCase() === wanted) out.push({ thread: cm });
    for (const list of elements(cm))
      if (is(list, 'replyLst'))
        for (const r of elements(list))
          // A reply without an id is not shown either.
          if (is(r, 'reply') && attr(r, 'id') && attr(r, 'id')!.toUpperCase() === wanted)
            out.push({ thread: cm, reply: r });
  }
  return out;
}

/**
 * A threaded comment or reply of a slide by its id key (`modernId`): the GUID of a comment that
 * the slide holds once, the GUID and mark of a comment that no other comment repeats, or the
 * GUID, mark, and place of a twin. Returns its list, its thread, the reply, whether the key holds
 * a place, and how many comments have the GUID and mark.
 */
export function findModern(
  batch: OpBatch,
  slidePart: string,
  key: string
): {
  part: string;
  thread: Element;
  reply?: Element;
  placed: boolean;
  group: number;
  twinTexts: () => string[];
} {
  const parsed = modernKeyParts(key);
  if (!parsed) throw new CommandRefusal('errors.noComment');
  const { guid, mark, ordinal } = parsed;
  const markOf = (hit: ModernHit): string => modernMark(hit.reply ?? hit.thread);
  // Read first, so a list without the comment is not rewritten as an XML part.
  const lists = slideCommentParts(batch.pkg, slidePart).modern.map((part) => ({
    part,
    hits: modernHits(commentTree(batch.pkg, part)?.root, guid),
  }));
  const total = lists.reduce((n, l) => n + l.hits.length, 0);
  // A plain GUID names a comment only while the slide holds that GUID once.
  if (mark === undefined && total !== 1) throw new CommandRefusal('errors.noComment');
  const inGroup = (h: ModernHit): boolean => mark === undefined || markOf(h) === mark;
  const group = lists.reduce((n, l) => n + l.hits.filter(inGroup).length, 0);
  // A GUID and mark without a place name a comment only while no twin shares them.
  if (ordinal === undefined && group !== 1) throw new CommandRefusal('errors.noComment');
  let left = ordinal ?? 1;
  for (const { part, hits } of lists) {
    const index = hits.findIndex((h) => inGroup(h) && --left === 0);
    if (index < 0) continue;
    const hit = modernHits(editable(batch, part)?.root, guid)[index];
    // The texts of the other comments of the group, as the reader shows them.
    const twinTexts = (): string[] => {
      const all = lists.flatMap((l) => l.hits.filter(inGroup));
      const at = (ordinal ?? 1) - 1;
      return all.filter((_, i) => i !== at).map((h) => modernText(h.reply ?? h.thread));
    };
    if (hit) return { part, ...hit, placed: ordinal !== undefined, group, twinTexts };
    break;
  }
  throw new CommandRefusal('errors.noComment');
}

/** Add a reply to the thread `thread` of list `part`. Returns its GUID. */
export function replyModern(
  batch: OpBatch,
  part: string,
  thread: Element,
  text: string,
  author: { readonly name: string; readonly initials: string },
  date: string
): string {
  const authorId = modernAuthorId(batch, author);
  const cm = batch.element(part, thread.id);
  const id = newGuid();
  const reply = `<p188:reply authorId="${authorId}" created="${date}" id="${id}">${bodyXml(text)}</p188:reply>`;
  const list = child(cm, { ns: P188, local: 'replyLst' });
  if (list) {
    batch.push({
      op: 'insertNode',
      part,
      parent: list.id,
      index: list.children.length,
      spec: fragment(reply),
    });
    return id;
  }
  // `p188:replyLst` sits after the monikers and `p188:pos`, before `p188:txBody`.
  const kids = elements(cm);
  const at = kids.findIndex((c) => is(c, 'txBody') || is(c, 'extLst'));
  const index = at < 0 ? cm.children.length : cm.children.indexOf(kids[at]!);
  batch.push({
    op: 'insertNode',
    part,
    parent: cm.id,
    index,
    spec: fragment(`<p188:replyLst>${reply}</p188:replyLst>`),
  });
  return id;
}

/** Replace the text of a threaded comment or reply. A body that holds more than text is kept. */
export function editModern(batch: OpBatch, part: string, el: Element, text: string): void {
  if (!modernTextEditable(el)) throw new CommandRefusal('disabled.commentNotEditable');
  const body = child(el, { ns: P188, local: 'txBody' })!;
  for (const p of elements(body))
    if (p.ns === NS.a && p.local === 'p') batch.push({ op: 'removeNode', part, node: p.id });
  const fresh = batch.element(part, body.id);
  for (const spec of text.split('\n').map((line) => fragment(paragraphsXml(line))))
    batch.push({ op: 'insertNode', part, parent: fresh.id, index: fresh.children.length, spec });
}

/**
 * Delete a reply, or a thread with its replies. An empty list stays while the deck is open, so a
 * comment that a collaborator adds at the same time is kept; the save drops empty lists.
 */
export function deleteModern(
  batch: OpBatch,
  part: string,
  thread: Element,
  reply: Element | undefined
): void {
  if (reply) {
    const list = elements(thread).find((c) => is(c, 'replyLst') && c.children.includes(reply))!;
    const only = elements(list).length === 1;
    batch.push({ op: 'removeNode', part, node: only ? list.id : reply.id });
    return;
  }
  batch.push({ op: 'removeNode', part, node: thread.id });
}

/** Set the status of a thread: `resolved`, or active (the default, no attribute). */
export function setModernStatus(
  batch: OpBatch,
  part: string,
  thread: Element,
  resolved: boolean
): void {
  const status = attr(thread, 'status');
  const isResolved = status === 'resolved' || status === 'closed';
  if (isResolved === resolved) return;
  batch.push({
    op: 'setAttribute',
    part,
    node: thread.id,
    name: { ns: '', local: 'status' },
    value: resolved ? 'resolved' : null,
  });
}
