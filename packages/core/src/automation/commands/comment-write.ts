// What both comment writers share: the checks of text, author, and date, the comment parts of a
// slide, and the removal of a comment list that holds no comment any more. The writers are
// `comment-modern.ts` (threaded comments) and `comment-legacy.ts` (ECMA-376 Part 1, 19.4).

import {
  attr,
  child,
  elements,
  NS,
  parseXmlPart,
  presentationInfo,
  relationshipsOf,
  relsPartName,
  xmlPartOf,
  type Element,
  type NodeSpec,
} from '../../store/index.ts';
import { specOf } from '../../store/index.ts';
import type { CommentAuthorInput } from '../../contracts/neutral.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { freePartName, removeOverride, xmlText } from './package-edits.ts';
import { partNameKey } from '@eigenpal/ooxml-core';
import { cleanText } from './text-units.ts';
import { COMMENTS } from './comment-read.ts';

/** The longest comment text, in UTF-16 code units. */
export const MAX_COMMENT_TEXT = 10_000;
const MAX_AUTHOR = 255;
const MAX_INITIALS = 8;

/** Comment text as the writers store it: line feeds, no control characters, 1 to 10,000 long. */
export function commentText(raw: unknown): string {
  if (typeof raw !== 'string') throw new CommandRefusal('errors.commentText');
  const text = cleanText(raw)
    .replace(/\u000B/g, '\n')
    .trim();
  if (!text || text.length > MAX_COMMENT_TEXT) throw new CommandRefusal('errors.commentText');
  xmlText(text, 'comment');
  return text;
}

/** The author of a new comment: the command's, else the editor's. Refused without a name. */
export function commentAuthor(
  input: CommentAuthorInput,
  fallback: string | undefined
): { readonly name: string; readonly initials: string } {
  const raw = typeof input.author === 'string' && input.author.trim() ? input.author : fallback;
  const name = typeof raw === 'string' ? cleanText(raw).replace(/\s+/g, ' ').trim() : '';
  if (!name) throw new CommandRefusal('disabled.noCommentAuthor');
  if (name.length > MAX_AUTHOR) throw new CommandRefusal('errors.commentAuthor');
  const given =
    typeof input.initials === 'string' ? cleanText(input.initials).replace(/\s+/g, '') : '';
  const initials = (given || initialsOf(name)).slice(0, MAX_INITIALS);
  xmlText(name, 'author');
  xmlText(initials, 'initials');
  return { name, initials };
}

/** The first letters of the first two words, upper case: "Ada Lovelace" is "AL". */
export function initialsOf(name: string): string {
  return name
    .split(' ')
    .filter(Boolean)
    .slice(0, 2)
    .map((w) => String.fromCodePoint(w.codePointAt(0)!).toUpperCase())
    .join('');
}

/** `xsd:dateTime`, with optional fraction and zone. */
const DATE_TIME = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,9})?(?:Z|[+-]\d{2}:\d{2})?$/;

/** The date of a new comment: the command's, else now in UTC, as the file stores it. */
export function commentDate(raw: string | undefined, now: () => Date = () => new Date()): string {
  if (raw !== undefined) {
    if (typeof raw !== 'string' || !DATE_TIME.test(raw) || Number.isNaN(Date.parse(raw)))
      throw new CommandRefusal('errors.commentDate');
    return raw;
  }
  // Without a zone, as presentation files store comment dates: UTC with milliseconds.
  return now().toISOString().replace(/Z$/, '');
}

/** A random version 4 GUID in braces, upper case (`ST_Guid`). */
export function newGuid(): string {
  const b = globalThis.crypto.getRandomValues(new Uint8Array(16));
  b[6] = (b[6]! & 0x0f) | 0x40;
  b[8] = (b[8]! & 0x3f) | 0x80;
  let hex = '';
  for (const x of b) hex += x.toString(16).padStart(2, '0');
  hex = hex.toUpperCase();
  return `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}}`;
}

/** Every prefix that a comment fragment uses, bound for parsing. */
const WRAP =
  `xmlns:a="${NS.a}" xmlns:p="${NS.p}" xmlns:r="${NS.r}" xmlns:p188="${COMMENTS.p188}" ` +
  `xmlns:p15="${NS.p15}"`;

/** Parse an XML fragment into a node spec. The fragment's own namespace declarations stay. */
export function fragment(xml: string): NodeSpec {
  const r = parseXmlPart('/fragment.xml', `<w ${WRAP}>${xml}</w>`);
  const first = r.ok ? elements(r.part.root)[0] : undefined;
  if (!first) throw new CommandRefusal('errors.unsupported', 'comment template');
  return specOf(first);
}

/** The slide of a comment id (`<slide part>#<comment>`), checked against the deck. */
export function splitCommentId(
  batch: OpBatch,
  commentId: unknown
): { readonly slidePart: string; readonly key: string } {
  if (typeof commentId !== 'string') throw new CommandRefusal('errors.noComment');
  // A part name holds no `#`, so the first one ends it; a comment key from a file can hold more.
  const hash = commentId.indexOf('#');
  const slidePart = commentId.slice(0, hash);
  const key = commentId.slice(hash + 1);
  if (
    hash <= 0 ||
    !key ||
    !presentationInfo(batch.pkg).slides.some((s) => s.partName === slidePart)
  )
    throw new CommandRefusal('errors.noComment');
  return { slidePart, key };
}

/** The `p:extLst` of a slide, created as its last child when it has none. */
export function slideExtList(batch: OpBatch, slidePart: string): Element {
  const root = xmlPartOf(batch.pkg, slidePart)?.root;
  if (!root) throw new CommandRefusal('errors.malformed', slidePart);
  const own = child(root, 'p:extLst');
  if (own) return own;
  const id = batch.insert({
    op: 'insertNode',
    part: slidePart,
    parent: root.id,
    index: root.children.length,
    spec: { ns: NS.p, local: 'extLst', prefix: 'p' },
  });
  return batch.element(slidePart, id);
}

/**
 * Remove an empty comment list: its part, its content-type override, the slide's relationship
 * to it, and for a threaded list the slide's `p188:commentRel`.
 */
export function dropCommentList(batch: OpBatch, slidePart: string, listPart: string): void {
  const rels = xmlPartOf(batch.pkg, relsPartName(slidePart));
  const ids = new Set(
    relationshipsOf(batch.pkg, slidePart)
      .filter((r) => !r.external && r.target === listPart)
      .map((r) => r.id)
  );
  for (const r of elements(rels?.root))
    if (ids.has(attr(r, 'Id') ?? ''))
      batch.push({ op: 'removeNode', part: relsPartName(slidePart), node: r.id });
  const extLst = child(xmlPartOf(batch.pkg, slidePart)?.root, 'p:extLst');
  for (const ext of elements(extLst)) {
    const rel = elements(ext).find((c) => c.ns === COMMENTS.p188 && c.local === 'commentRel');
    if (rel && ids.has(attr(rel, { ns: NS.r, local: 'id' }) ?? ''))
      batch.push({ op: 'removeNode', part: slidePart, node: ext.id });
  }
  const left = child(xmlPartOf(batch.pkg, slidePart)?.root, 'p:extLst');
  if (left && elements(left).length === 0)
    batch.push({ op: 'removeNode', part: slidePart, node: left.id });
  // Another slide can still name the part (a file can share one): it then stays.
  const shared = presentationInfo(batch.pkg).slides.some((s) =>
    relationshipsOf(batch.pkg, s.partName).some((r) => !r.external && r.target === listPart)
  );
  if (shared) return;
  removeOverride(batch, listPart);
  if (xmlPartOf(batch.pkg, relsPartName(listPart)))
    batch.push({ op: 'removePart', part: relsPartName(listPart) });
  batch.push({ op: 'removePart', part: listPart });
}

/** `preferred` when it is free, else the first free name `prefix<n>suffix`. */
export function freeName(
  batch: OpBatch,
  preferred: string,
  prefix: string,
  suffix: string
): string {
  const key = partNameKey(preferred);
  const taken =
    batch.reservedPartNames.has(key) ||
    [...batch.pkg.parts.keys()].some(
      (n) => partNameKey(n) === key || partNameKey(n) === partNameKey(relsPartName(preferred))
    );
  // In a collaboration session a name from the actor's stripe: two peers who add the first
  // author list at once must not write one part, or one of their authors is lost.
  return taken || batch.allocationActor ? freePartName(batch, prefix, suffix) : preferred;
}
