// Comments of imported slides. A comment list (`p:cmLst`, ECMA-376 Part 1, 19.4.3) names its
// authors by `p:cm/@authorId`, an id of the presentation's comment author list (`p:cmAuthorLst`,
// 19.4.2). A copied list therefore takes the destination's ids: an author of the source with
// the same name and initials as a destination author becomes that author, any other is added
// with the next free id. Comment indexes are then renumbered so each stays unique per author.

import {
  attr,
  contentTypeOf,
  elements,
  NS,
  presentationInfo,
  relationshipsOf,
  relsPartName,
  xmlPartOf,
  type Element,
  type PptxOp,
  type PptxPackage,
} from '../../store/index.ts';
import { partNameKey } from '@eigenpal/ooxml-core';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { addOverride, addRelationship, freePartName, stripedIds } from './package-edits.ts';
import { editable, renumberComments, treeOf } from './slide-structure.ts';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const PML = 'application/vnd.openxmlformats-officedocument.presentationml.';

/** The relationship from a slide to its comment list. */
export const REL_COMMENTS = `${R}comments`;
/** The content type of a comment list. */
export const CT_COMMENTS = `${PML}comments+xml`;
const REL_AUTHORS = `${R}commentAuthors`;
const CT_AUTHORS = `${PML}commentAuthors+xml`;

/** `p:cmAuthor/@id` is an `xsd:unsignedInt`. */
const MAX_AUTHOR_ID = 4_294_967_295;
/** The most authors one import adds. */
const MAX_AUTHORS = 10_000;

interface Author {
  readonly name: string;
  readonly initials: string;
}

/** The comment author list part of a package, if it has one. */
function authorsPart(pkg: PptxPackage): string | undefined {
  const presentation = presentationInfo(pkg).presentationPart;
  return relationshipsOf(pkg, presentation).find((r) => r.type === REL_AUTHORS && !r.external)
    ?.target;
}

/** The authors of a list, by id. */
function authorsOf(root: Element | undefined): Map<string, Author> {
  const out = new Map<string, Author>();
  if (root?.ns !== NS.p || root.local !== 'cmAuthorLst') return out;
  for (const a of elements(root)) {
    const id = attr(a, 'id');
    if (a.local !== 'cmAuthor' || id === undefined || out.has(id)) continue;
    out.set(id, { name: attr(a, 'name') ?? '', initials: attr(a, 'initials') ?? '' });
  }
  return out;
}

/** `/ppt/commentAuthors.xml` when it is free, as most packages name it, else a numbered name. */
function freeAuthorsName(batch: OpBatch): string {
  const plain = '/ppt/commentAuthors.xml';
  const taken = (name: string): boolean => {
    const key = partNameKey(name);
    return (
      batch.reservedPartNames.has(key) ||
      [...batch.pkg.parts.keys()].some((n) => partNameKey(n) === key)
    );
  };
  if (!taken(plain) && !taken(relsPartName(plain))) return plain;
  return freePartName(batch, '/ppt/commentAuthors', '.xml');
}

const keyOf = (a: Author): string => `${a.name}\u0000${a.initials}`;

/**
 * Give the copied comment lists `copies` the destination's author ids, adding authors the
 * destination does not have, then renumber their indexes. A comment whose author the source
 * does not list is dropped: no author id of the destination can stand for it.
 */
export function mergeImportedComments(
  batch: OpBatch,
  source: PptxPackage,
  copies: readonly string[]
): void {
  if (copies.length === 0) return;
  const presentation = presentationInfo(batch.pkg).presentationPart;
  const sourceAuthors = authorsOf(treeOf(source, authorsPart(source))?.root);

  let part = authorsPart(batch.pkg);
  if (!part || contentTypeOf(batch.pkg, part) !== CT_AUTHORS || !treeOf(batch.pkg, part)) {
    if (part) throw new CommandRefusal('errors.malformed', 'commentAuthors');
    part = freeAuthorsName(batch);
    batch.push({ op: 'addPart', part, xml: `<p:cmAuthorLst xmlns:p="${NS.p}"/>` });
    addOverride(batch, part, CT_AUTHORS);
    addRelationship(batch, presentation, REL_AUTHORS, part);
  }
  const list = editable(batch, part)?.root;
  if (list?.ns !== NS.p || list.local !== 'cmAuthorLst')
    throw new CommandRefusal('errors.malformed', 'commentAuthors');

  const byKey = new Map<string, string>();
  const used = new Set<number>();
  let top = -1;
  let colors = -1;
  for (const a of elements(list)) {
    if (a.local !== 'cmAuthor') continue;
    // `xsd:unsignedInt` (19.4.2): another id is no id a comment can name.
    const raw = attr(a, 'id') ?? '';
    const id = /^\d{1,10}$/.test(raw) ? Number(raw) : NaN;
    if (!(id <= MAX_AUTHOR_ID)) continue;
    used.add(id);
    top = Math.max(top, id);
    colors = Math.max(colors, Number(attr(a, 'clrIdx')) || 0);
    const key = keyOf({ name: attr(a, 'name') ?? '', initials: attr(a, 'initials') ?? '' });
    if (!byKey.has(key)) byKey.set(key, String(id));
  }

  // In a collaboration session a new author takes an id of the actor's stripe: the author list is
  // an XML tree that merges, so two peers who import at once must not add two authors with one id.
  const striped = batch.allocationActor
    ? stripedIds(batch.allocationActor, used, 0, MAX_AUTHOR_ID, batch.mintedOf('commentAuthor'))
    : undefined;
  // Source author id → destination author id.
  const ids = new Map<string, string>();
  const added: PptxOp[] = [];
  const destinationId = (sourceId: string): string | undefined => {
    const known = ids.get(sourceId);
    if (known !== undefined) return known;
    const author = sourceAuthors.get(sourceId);
    if (!author) return undefined;
    let id = byKey.get(keyOf(author));
    if (id === undefined) {
      if (added.length >= MAX_AUTHORS || top >= MAX_AUTHOR_ID)
        throw new CommandRefusal('errors.tooLarge', 'comment authors');
      top = striped ? striped() : top + 1;
      colors += 1;
      id = String(top);
      byKey.set(keyOf(author), id);
      added.push({
        op: 'insertNode',
        part: part!,
        parent: list.id,
        index: list.children.length + added.length,
        spec: {
          ns: NS.p,
          local: 'cmAuthor',
          prefix: 'p',
          // In name order, as a collaboration room writes attributes back, so a replica that
          // made the author and one that received it save the same bytes: the comment author
          // list is a part the reader keeps as bytes.
          attrs: [
            { ns: '', local: 'clrIdx', value: String(colors) },
            { ns: '', local: 'id', value: id },
            { ns: '', local: 'initials', value: author.initials },
            { ns: '', local: 'lastIdx', value: '0' },
            { ns: '', local: 'name', value: author.name },
          ],
        },
      });
    }
    ids.set(sourceId, id);
    return id;
  };

  const edits: PptxOp[] = [];
  // The source author of each element whose author id changes, for the renumbering.
  const original = new Map<string, string>();
  const setAuthor = (copy: string, el: Element, value: string): void => {
    original.set(`${copy}\u0000${el.id}`, attr(el, 'authorId') ?? '');
    if (attr(el, 'authorId') !== value)
      edits.push({
        op: 'setAttribute',
        part: copy,
        node: el.id,
        name: { ns: '', local: 'authorId' },
        value,
      });
  };
  for (const copy of copies) {
    const root = xmlPartOf(batch.pkg, copy)?.root;
    if (root?.ns !== NS.p || root.local !== 'cmLst')
      throw new CommandRefusal('errors.malformed', 'comments');
    for (const cm of elements(root)) {
      if (cm.local !== 'cm') continue;
      const id = destinationId(attr(cm, 'authorId') ?? '');
      if (id === undefined) {
        edits.push({ op: 'removeNode', part: copy, node: cm.id });
        continue;
      }
      setAuthor(copy, cm, id);
      // A reply names its parent comment by author and index (`parentCm` of the threading
      // extension, the only element of that name in a comment).
      const walk = (el: Element, depth: number): void => {
        if (depth > 16) return;
        if (el.local === 'parentCm') {
          const parent = destinationId(attr(el, 'authorId') ?? '');
          if (parent !== undefined) setAuthor(copy, el, parent);
          else original.set(`${copy}\u0000${el.id}`, attr(el, 'authorId') ?? '');
        }
        for (const c of elements(el)) walk(c, depth + 1);
      };
      walk(cm, 0);
    }
  }
  batch.pushAll(added);
  batch.pushAll(edits);
  // A reply whose parent was dropped, or never came along, loses its threading extension.
  renumberComments(batch, copies, presentation, {
    authorOf: (copy, el) => original.get(`${copy}\u0000${el.id}`) ?? attr(el, 'authorId'),
    dropOrphanReplies: true,
  });
}
