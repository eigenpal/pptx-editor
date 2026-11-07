// Comment structure as a save writes it. A deck holds one author list per format, one comment
// list per slide and format, one reply list per threaded comment, and a threaded comment names
// the slide that holds it. Two collaborators who add the first author, the first comment of a
// slide, or the first reply of a thread at the same time each add one, and a file can hold
// such copies too. The save joins the copies into the first one, drops a list that holds no
// comment, and points each threaded comment at its own slide, so the saved file holds each list
// once. The store does not change.

import {
  attr,
  child,
  elements,
  NS,
  presentationInfo,
  relationshipsOf,
  relsPartName,
  specOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../../store/index.ts';
import { OpBatch } from './op-batch.ts';
import { removeOverride } from './package-edits.ts';
import { COMMENTS, presentationTargets, slideCommentParts } from './comment-read.ts';
import { dropCommentList, fragment } from './comment-write.ts';

const P188 = COMMENTS.p188;

/** The key that names one author or one comment once in a list. */
type KeyOf = (el: Element) => string | undefined;

const authorKey: KeyOf = (a) =>
  a.ns === P188 ? attr(a, 'id')?.toUpperCase() : attr(a, 'id')?.replace(/^0+(?=\d)/, '');

/** Move the children of `from` that `into` does not hold (by `key`) to the end of `into`. */
function join(batch: OpBatch, into: string, from: string, key: KeyOf): void {
  const target = xmlPartOf(batch.pkg, into)?.root;
  const source = xmlPartOf(batch.pkg, from)?.root;
  if (!target || !source || target.local !== source.local || target.ns !== source.ns) return;
  const have = new Set(elements(target).map(key));
  for (const el of elements(source)) {
    const k = key(el);
    if (k !== undefined && have.has(k)) continue;
    have.add(k);
    const root = xmlPartOf(batch.pkg, into)!.root;
    batch.push({
      op: 'insertNode',
      part: into,
      parent: root.id,
      index: root.children.length,
      spec: specOf(el),
    });
  }
}

/** Join the author lists of one format into the first, and drop the others. */
function joinAuthors(batch: OpBatch, type: string): void {
  const parts = presentationTargets(batch.pkg, type).filter((p) => xmlPartOf(batch.pkg, p));
  if (parts.length < 2) return;
  const [keep, ...rest] = parts;
  const presentation = presentationInfo(batch.pkg).presentationPart;
  for (const part of rest) {
    join(batch, keep!, part, authorKey);
    const rels = xmlPartOf(batch.pkg, relsPartName(presentation))?.root;
    const ids = new Set(
      relationshipsOf(batch.pkg, presentation)
        .filter((r) => r.type === type && r.target === part)
        .map((r) => r.id)
    );
    for (const r of elements(rels))
      if (ids.has(attr(r, 'Id') ?? ''))
        batch.push({ op: 'removeNode', part: relsPartName(presentation), node: r.id });
    removeOverride(batch, part);
    if (xmlPartOf(batch.pkg, relsPartName(part)))
      batch.push({ op: 'removePart', part: relsPartName(part) });
    batch.push({ op: 'removePart', part });
  }
}

/** Join the comment lists of one slide and format into the first. */
function joinLists(
  batch: OpBatch,
  slidePart: string,
  format: 'legacy' | 'modern'
): string | undefined {
  const parts = slideCommentParts(batch.pkg, slidePart)[format].filter((p) =>
    xmlPartOf(batch.pkg, p)
  );
  const [keep, ...rest] = parts;
  const key: KeyOf =
    format === 'modern'
      ? (cm) => attr(cm, 'id')?.toUpperCase()
      : (cm) => `${attr(cm, 'authorId')}:${attr(cm, 'idx')}`;
  for (const part of rest) {
    join(batch, keep!, part, key);
    dropCommentList(batch, slidePart, part);
  }
  // A delete keeps an emptied list while the deck is open; the saved file never holds one.
  if (keep && elements(xmlPartOf(batch.pkg, keep)?.root).length === 0) {
    dropCommentList(batch, slidePart, keep);
    return undefined;
  }
  return keep;
}

/** The slide's `p188:commentRel` extensions, by `p:ext`. */
function commentRels(batch: OpBatch, slidePart: string): Element[] {
  const extLst = child(xmlPartOf(batch.pkg, slidePart)?.root, 'p:extLst');
  return elements(extLst).filter((ext) =>
    elements(ext).some((c) => c.ns === P188 && c.local === 'commentRel')
  );
}

/** One `p188:commentRel` that names the relationship to `list`. */
function pointAt(batch: OpBatch, slidePart: string, list: string): void {
  const rId = relationshipsOf(batch.pkg, slidePart).find(
    (r) => r.type === COMMENTS.modernRel && r.target === list
  )?.id;
  if (!rId) return;
  const [first, ...extra] = commentRels(batch, slidePart);
  for (const ext of extra) batch.push({ op: 'removeNode', part: slidePart, node: ext.id });
  const rel = first && elements(first).find((c) => c.local === 'commentRel')!;
  const name = { ns: NS.r, local: 'id', prefix: 'r' };
  if (rel && attr(rel, { ns: NS.r, local: 'id' }) !== rId)
    batch.push({ op: 'setAttribute', part: slidePart, node: rel.id, name, value: rId });
  if (rel) return;
  const root = xmlPartOf(batch.pkg, slidePart)!.root;
  let extLst = child(root, 'p:extLst');
  if (!extLst) {
    batch.push({
      op: 'insertNode',
      part: slidePart,
      parent: root.id,
      index: root.children.length,
      spec: { ns: NS.p, local: 'extLst', prefix: 'p' },
    });
    extLst = child(xmlPartOf(batch.pkg, slidePart)!.root, 'p:extLst')!;
  }
  batch.push({
    op: 'insertNode',
    part: slidePart,
    parent: extLst.id,
    index: extLst.children.length,
    spec: fragment(
      `<p:ext uri="${COMMENTS.commentRelUri}"><p188:commentRel xmlns:p188="${P188}" r:id="${rId}"/></p:ext>`
    ),
  });
}

/** The slide's `p14:creationId/@val`, if it has one. */
function creationOf(batch: OpBatch, slidePart: string): string | undefined {
  const extLst = child(xmlPartOf(batch.pkg, slidePart)?.root, 'p:extLst');
  for (const ext of elements(extLst))
    for (const c of elements(ext))
      if (c.ns === NS.p14 && c.local === 'creationId') return attr(c, 'val');
  return undefined;
}

/** Each threaded comment of `list`: one reply list, and slide monikers that name this slide. */
function fixThreads(batch: OpBatch, list: string, sldId: string, creation: string | undefined) {
  for (const cm of elements(xmlPartOf(batch.pkg, list)?.root)) {
    if (cm.ns !== P188 || cm.local !== 'cm') continue;
    const lists = elements(cm).filter((c) => c.ns === P188 && c.local === 'replyLst');
    for (const extra of lists.slice(1)) {
      for (const reply of elements(extra)) {
        const into = batch.element(list, lists[0]!.id);
        batch.push({
          op: 'moveNode',
          part: list,
          node: reply.id,
          parent: into.id,
          index: into.children.length,
        });
      }
      batch.push({ op: 'removeNode', part: list, node: extra.id });
    }
    for (const monikers of elements(cm)) {
      if (!monikers.local.endsWith('MkLst')) continue;
      const slide = elements(monikers).find((m) => m.ns === COMMENTS.pc && m.local === 'sldMk');
      if (!slide) continue;
      const set = (name: string, value: string | undefined): void => {
        if (value !== undefined && attr(slide, name) !== value)
          batch.push({
            op: 'setAttribute',
            part: list,
            node: slide.id,
            name: { ns: '', local: name },
            value,
          });
      };
      set('sldId', sldId);
      set('cId', creation);
    }
  }
}

/** The package with its comment structure as a save writes it; the same package when it is. */
export function repairCommentsForSave(pkg: PptxPackage): PptxPackage {
  const slides = presentationInfo(pkg).slides;
  const commented = slides.filter((s) =>
    relationshipsOf(pkg, s.partName).some(
      (r) => r.type === COMMENTS.modernRel || r.type === COMMENTS.legacyRel
    )
  );
  const twice = [COMMENTS.modernAuthorsRel, COMMENTS.legacyAuthorsRel].some(
    (type) => presentationTargets(pkg, type).length > 1
  );
  if (commented.length === 0 && !twice) return pkg;
  const batch = new OpBatch(pkg);
  joinAuthors(batch, COMMENTS.modernAuthorsRel);
  joinAuthors(batch, COMMENTS.legacyAuthorsRel);
  for (const slide of commented) {
    joinLists(batch, slide.partName, 'legacy');
    const list = joinLists(batch, slide.partName, 'modern');
    if (!list) continue;
    pointAt(batch, slide.partName, list);
    fixThreads(batch, list, slide.sldId, creationOf(batch, slide.partName));
  }
  return batch.ops.length ? batch.pkg : pkg;
}
