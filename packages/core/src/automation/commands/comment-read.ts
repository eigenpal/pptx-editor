// The comment threads of a slide, read from both comment formats:
//
// - Legacy (ECMA-376 Part 1, 19.4): `p:cmAuthorLst` (the presentation's `commentAuthors` part)
//   names authors by `p:cmAuthor/@id`; each slide relates to one `p:cmLst` of `p:cm authorId dt
//   idx` with `p:pos` and `p:text`. A reply names its parent in `p15:threadingInfo/p15:parentCm`.
//   `p:pos` counts in 1/576 inch.
// - Threaded (the `p188` extension): `p188:authorLst` (the presentation's `authors` part) names
//   authors by GUID; each slide relates to a `p188:cmLst` of `p188:cm id authorId created
//   status`, each with an anchor moniker list (`pc:sldMkLst`, or `ac:deMkLst` for an element),
//   an optional `p188:pos`, `p188:replyLst/p188:reply`, and a `p188:txBody`.
//
// Comment parts are kept as bytes by the reader; this module parses them with the same bounded
// parser, once per part object. Every value comes from the file and is shown as text.

import {
  attr,
  child,
  elements,
  NS,
  parseXmlPart,
  presentationInfo,
  relationshipsOf,
  xmlPartOf,
  type Element,
  type PackagePart,
  type PptxPackage,
  type XmlPart,
} from '../../store/index.ts';
import type { CommentAnchor, CommentThread, SlideComment } from '../../contracts/neutral.ts';
import { unsignedIntId } from './package-edits.ts';

const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const MS = 'http://schemas.microsoft.com/office/2018/10/relationships/';
const PML = 'application/vnd.openxmlformats-officedocument.presentationml.';

/** Namespaces and part types of both comment formats. */
export const COMMENTS = Object.freeze({
  p188: 'http://schemas.microsoft.com/office/powerpoint/2018/8/main',
  pc: 'http://schemas.microsoft.com/office/powerpoint/2013/main/command',
  ac: 'http://schemas.microsoft.com/office/drawing/2013/main/command',
  legacyRel: `${OD}comments`,
  legacyAuthorsRel: `${OD}commentAuthors`,
  legacyType: `${PML}comments+xml`,
  legacyAuthorsType: `${PML}commentAuthors+xml`,
  modernRel: `${MS}comments`,
  modernAuthorsRel: `${MS}authors`,
  modernType: 'application/vnd.ms-powerpoint.comments+xml',
  modernAuthorsType: 'application/vnd.ms-powerpoint.authors+xml',
  /** The `p:ext` of a slide that names its threaded comment part (`p188:commentRel`). */
  commentRelUri: '{6950BFC3-D8DA-4A85-94F7-54DA5524770B}',
  /** The `p:ext` of a legacy reply (`p15:threadingInfo`). */
  threadingUri: '{C676402C-5697-4E1C-873F-D02D1690AC5C}',
  /** The `a:ext` of a drawing element's `a16:creationId`. */
  creationIdUri: '{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}',
});

/** EMU per legacy `p:pos` unit (1/576 inch). */
export const LEGACY_POS_EMU = 1587.5;
/** How deep the readers walk into a comment element. */
const MAX_DEPTH = 32;

/** The parsed tree of a part, once per part object: comment parts are kept as bytes. */
const trees = new WeakMap<PackagePart, XmlPart | null>();

export function commentTree(pkg: PptxPackage, name: string | undefined): XmlPart | undefined {
  const part = name ? pkg.parts.get(name) : undefined;
  if (!part) return undefined;
  if (part.kind === 'xml') return part.xml;
  let tree = trees.get(part);
  if (tree === undefined) {
    const parsed = parseXmlPart(part.partName, decode(part.bytes));
    tree = parsed.ok ? parsed.part : null;
    trees.set(part, tree);
  }
  return tree ?? undefined;
}

function decode(bytes: Uint8Array): string {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes);
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes);
  const text = new TextDecoder('utf-8').decode(bytes);
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

/** An author of either list. */
export interface CommentAuthor {
  readonly name: string;
  readonly initials: string;
}

/** The internal target of the presentation's first relationship of `type`. */
export function presentationTarget(pkg: PptxPackage, type: string): string | undefined {
  const presentation = presentationInfo(pkg).presentationPart;
  return relationshipsOf(pkg, presentation).find((r) => r.type === type && !r.external)?.target;
}

/**
 * The internal targets of the presentation's relationships of `type`, in order, once each. A
 * file has one author list; two collaborators who add the first comment at once give the
 * presentation two until a save joins them (`comment-repair.ts`).
 */
export function presentationTargets(pkg: PptxPackage, type: string): string[] {
  const presentation = presentationInfo(pkg).presentationPart;
  const out: string[] = [];
  for (const r of relationshipsOf(pkg, presentation))
    if (r.type === type && !r.external && r.target && !out.includes(r.target)) out.push(r.target);
  return out;
}

/** Legacy authors by canonical id (`xsd:unsignedInt`, so "007" is author 7). */
export function legacyAuthors(pkg: PptxPackage): Map<string, CommentAuthor> {
  const out = new Map<string, CommentAuthor>();
  for (const part of presentationTargets(pkg, COMMENTS.legacyAuthorsRel)) {
    const root = commentTree(pkg, part)?.root;
    if (root?.ns !== NS.p || root.local !== 'cmAuthorLst') continue;
    for (const a of elements(root)) {
      const id = unsignedIntId(attr(a, 'id'));
      if (a.local !== 'cmAuthor' || id === undefined || out.has(String(id))) continue;
      out.set(String(id), { name: attr(a, 'name') ?? '', initials: attr(a, 'initials') ?? '' });
    }
  }
  return out;
}

/** Threaded authors by GUID, upper case. */
export function modernAuthors(pkg: PptxPackage): Map<string, CommentAuthor> {
  const out = new Map<string, CommentAuthor>();
  for (const part of presentationTargets(pkg, COMMENTS.modernAuthorsRel)) {
    const root = commentTree(pkg, part)?.root;
    if (root?.ns !== COMMENTS.p188 || root.local !== 'authorLst') continue;
    for (const a of elements(root)) {
      const id = attr(a, 'id')?.toUpperCase();
      if (a.local !== 'author' || !id || out.has(id)) continue;
      out.set(id, { name: attr(a, 'name') ?? '', initials: attr(a, 'initials') ?? '' });
    }
  }
  return out;
}

/**
 * The id of the author `name` with `initials` in any author list of `format`, as the file writes
 * it. A peer that reads only the first list would add itself again after a join.
 */
export function findCommentAuthor(
  pkg: PptxPackage,
  format: 'legacy' | 'modern',
  author: { readonly name: string; readonly initials: string }
): string | undefined {
  const rel = format === 'legacy' ? COMMENTS.legacyAuthorsRel : COMMENTS.modernAuthorsRel;
  const local = format === 'legacy' ? 'cmAuthor' : 'author';
  for (const part of presentationTargets(pkg, rel))
    for (const a of elements(commentTree(pkg, part)?.root))
      if (
        a.local === local &&
        attr(a, 'name') === author.name &&
        attr(a, 'initials') === author.initials &&
        attr(a, 'id')
      )
        return attr(a, 'id');
  return undefined;
}

/** The comment parts of a slide: its legacy list and its threaded lists. */
export function slideCommentParts(
  pkg: PptxPackage,
  slidePart: string
): { readonly legacy: readonly string[]; readonly modern: readonly string[] } {
  const legacy: string[] = [];
  const modern: string[] = [];
  for (const r of relationshipsOf(pkg, slidePart)) {
    if (r.external || !r.target || !pkg.parts.has(r.target)) continue;
    if (r.type === COMMENTS.legacyRel && !legacy.includes(r.target)) legacy.push(r.target);
    if (r.type === COMMENTS.modernRel && !modern.includes(r.target)) modern.push(r.target);
  }
  return { legacy, modern };
}

/**
 * Whether any slide of the deck has a comment list of `format` that holds a comment. A list that
 * a delete emptied stays until the save, and does not count.
 */
export function deckUses(pkg: PptxPackage, format: 'legacy' | 'modern'): boolean {
  return presentationInfo(pkg).slides.some((s) =>
    slideCommentParts(pkg, s.partName)[format].some((part) => {
      const root = commentTree(pkg, part)?.root;
      return !root || elements(root).length > 0;
    })
  );
}

/** A legacy id field, with `-`, `#`, and `%` encoded, so the id key splits back into its fields. */
const idField = (raw: string): string => encodeURIComponent(raw).replace(/-/g, '%2D');

/** A field of a comment identity: encoded, so it holds no `~` and no `#`. */
const identityField = (raw: string): string => encodeURIComponent(raw).replace(/~/g, '%7E');

/** The key of a legacy comment id (`cm-<author>-<idx>`), with each field encoded. */
export const legacyKey = (authorId: string, idx: string): string =>
  `cm-${idField(authorId)}-${idField(idx)}`;

/** The author and index of a legacy id key, or undefined when the key is not one. */
export function legacyKeyFields(
  key: string
): { readonly authorId: string; readonly idx: string } | undefined {
  const m = /^cm-([^-]*)-([^-]*)$/.exec(key);
  if (!m) return undefined;
  try {
    return { authorId: decodeURIComponent(m[1]!), idx: decodeURIComponent(m[2]!) };
  } catch {
    return undefined;
  }
}

/** The comment id of a legacy comment: its author and index. */
export const legacyId = (slidePart: string, authorId: string, idx: string): string =>
  `${slidePart}#${legacyKey(authorId, idx)}`;

/**
 * The mark of a threaded comment or reply: its author id and creation time, as 8 hex digits
 * (FNV-1a). A delete or an edit of another comment never changes it.
 */
export function modernMark(el: Element): string {
  return fnv(`${(attr(el, 'authorId') ?? '').toUpperCase()}\u0000${attr(el, 'created') ?? ''}`);
}

/** FNV-1a of `text`, as 8 upper-case hex digits. */
function fnv(text: string): string {
  let hash = 0x811c9dc5;
  for (let i = 0; i < text.length; i += 1)
    hash = Math.imul(hash ^ text.charCodeAt(i), 0x01000193) >>> 0;
  return hash.toString(16).toUpperCase().padStart(8, '0');
}

/**
 * The comment id of a threaded comment or reply: its GUID. A GUID that the slide repeats (a file
 * can) or that holds `~` (no GUID does) also takes the comment's mark (`modernMark`) and its
 * ordinal among the repeats with that mark: `G~<mark>~1`. A delete of another repeat then moves
 * no id onto another comment. `modernKeyParts` reads the key back.
 */
export const modernId = (
  slidePart: string,
  guid: string,
  marked?: { readonly mark: string; readonly ordinal: number }
): string =>
  `${slidePart}#${guid.toUpperCase()}${marked ? `~${marked.mark}~${marked.ordinal}` : ''}`;

/** The parts of a threaded comment id key (`modernId`), or undefined for a key that is not one. */
export interface ModernKey {
  readonly guid: string;
  /** The mark of a repeated GUID, absent for a GUID that the slide holds once. */
  readonly mark?: string;
  readonly ordinal: number;
}

/** The parts of a threaded comment id key, in its canonical form only. */
export function modernKeyParts(key: string): ModernKey | undefined {
  const upper = key.toUpperCase();
  const m = /^(.+)~([0-9A-F]{8})~([1-9]\d{0,5})$/.exec(upper);
  if (m) return { guid: m[1]!, mark: m[2]!, ordinal: Number(m[3]) };
  return upper && !upper.includes('~') ? { guid: upper, ordinal: 1 } : undefined;
}

/** The first descendant with this namespace and local name, at most `depth` levels down. */
export function descendant(
  el: Element | undefined,
  ns: string,
  local: string,
  depth = MAX_DEPTH
): Element | undefined {
  if (!el || depth < 0) return undefined;
  for (const c of elements(el)) {
    if (c.ns === ns && c.local === local) return c;
    const hit = descendant(c, ns, local, depth - 1);
    if (hit) return hit;
  }
  return undefined;
}

/** The text of a `p188:txBody`: its paragraphs, one line feed apart; `a:br` is a line feed. */
export function bodyText(body: Element | undefined): string {
  const lines: string[] = [];
  for (const p of elements(body)) {
    if (p.ns !== NS.a || p.local !== 'p') continue;
    let line = '';
    for (const r of elements(p)) {
      if (r.local === 'br') line += '\n';
      else if (r.local === 'r' || r.local === 'fld') line += textOfT(r);
    }
    lines.push(line);
  }
  return lines.join('\n');
}

function textOfT(run: Element): string {
  const t = elements(run).find((c) => c.local === 't');
  let out = '';
  for (const c of t?.children ?? []) if (c.type === 'text') out += c.value;
  return out;
}

/** Elements a plain comment text body holds; any other content (a mention) is not edited. */
const PLAIN = new Set(['bodyPr', 'lstStyle', 'p', 'pPr', 'r', 'rPr', 't', 'br', 'endParaRPr']);

function plainBody(el: Element, depth = 0): boolean {
  if (depth > 0 && (el.ns !== NS.a || !PLAIN.has(el.local))) return false;
  if (
    el.local === 'rPr' ||
    el.local === 'pPr' ||
    el.local === 'endParaRPr' ||
    el.local === 'bodyPr'
  )
    return true;
  return depth > 4 ? false : elements(el).every((c) => plainBody(c, depth + 1));
}

/**
 * Whether `editComment` can replace the text of a threaded comment or reply: a body of plain
 * text and no extension list. The reader and the writer share this rule.
 */
export function modernTextEditable(el: Element): boolean {
  const body = child(el, { ns: COMMENTS.p188, local: 'txBody' });
  return !!body && plainBody(body) && !child(el, { ns: COMMENTS.p188, local: 'extLst' });
}

/** The legacy author id of an element, canonical, or the raw value when it is not a number. */
export const canonicalAuthor = (raw: string | undefined): string =>
  raw === undefined ? '' : String(unsignedIntId(raw) ?? raw);

/** The index of a legacy comment, canonical. */
export const canonicalIdx = (raw: string | undefined): string =>
  raw === undefined ? '' : String(unsignedIntId(raw) ?? raw);

/** The `a16:creationId` and `p:cNvPr/@id` of every drawing element of a slide, by node id. */
function shapeKeys(root: Element | undefined): {
  byCreation: Map<string, string>;
  byId: Map<string, string>;
} {
  const byCreation = new Map<string, string>();
  const byId = new Map<string, string>();
  const walk = (el: Element, depth: number): void => {
    if (depth > MAX_DEPTH) return;
    for (const c of elements(el)) {
      const nv = elements(c).find((e) => e.ns === NS.p && e.local.startsWith('nv'));
      const cNvPr = child(nv, 'p:cNvPr');
      if (cNvPr) {
        const id = attr(cNvPr, 'id');
        if (id !== undefined && !byId.has(id)) byId.set(id, c.id);
        const creation = attr(descendant(cNvPr, NS.a16, 'creationId', 3), 'id')?.toUpperCase();
        if (creation && !byCreation.has(creation)) byCreation.set(creation, c.id);
      }
      if (c.local === 'spTree' || c.local === 'grpSp' || c.local === 'cSld') walk(c, depth + 1);
    }
  };
  if (root) walk(root, 0);
  return { byCreation, byId };
}

/** The anchor of a threaded comment: an element moniker, a position, or the slide. */
function modernAnchor(cm: Element, keys: ReturnType<typeof shapeKeys>): CommentAnchor {
  for (const list of elements(cm)) {
    if (!list.local.endsWith('MkLst')) continue;
    // The innermost moniker of a drawing element names it by creation id and drawing id.
    const marks = elements(list).filter((m) => m.ns === COMMENTS.ac && m.local.endsWith('Mk'));
    const mark = marks[marks.length - 1];
    if (!mark) continue;
    const creation = attr(mark, 'creationId')?.toUpperCase();
    const node =
      (creation ? keys.byCreation.get(creation) : undefined) ??
      (attr(mark, 'id') !== undefined ? keys.byId.get(attr(mark, 'id')!) : undefined);
    if (node) return { kind: 'shape', nodeId: node };
  }
  return pointOf(child(cm, { ns: COMMENTS.p188, local: 'pos' }), 1);
}

function pointOf(pos: Element | undefined, scale: number): CommentAnchor {
  const x = Number(attr(pos, 'x'));
  const y = Number(attr(pos, 'y'));
  if (!pos || !Number.isFinite(x) || !Number.isFinite(y)) return { kind: 'slide' };
  return { kind: 'point', x: Math.round(x * scale), y: Math.round(y * scale) };
}

const optionalDate = (date: string | undefined) => (date !== undefined ? { date } : {});

/**
 * The ids of a slide's comments: how many threaded comments use each GUID (upper case), how many
 * use each GUID and mark (twins), and each GUID, mark, and text; and how many of each the reader
 * has met so far. Legacy comments count by author and index.
 */
interface GuidCount {
  readonly uses: Map<string, number>;
  readonly twins: Map<string, number>;
  readonly texts: Map<string, number>;
  readonly seen: Map<string, number>;
  readonly seenText: Map<string, number>;
  readonly legacy: Map<string, number>;
  readonly seenLegacy: Map<string, number>;
}

const bump = (map: Map<string, number>, key: string): number => {
  const n = (map.get(key) ?? 0) + 1;
  map.set(key, n);
  return n;
};

const modernText = (el: Element): string =>
  bodyText(child(el, { ns: COMMENTS.p188, local: 'txBody' }));

/** The threads and replies the reader shows: a thread with an id, a reply with an id. */
function eachModern(root: Element | undefined, visit: (el: Element) => void): void {
  if (root?.ns !== COMMENTS.p188 || root.local !== 'cmLst') return;
  for (const cm of elements(root)) {
    if (cm.ns !== COMMENTS.p188 || cm.local !== 'cm' || !attr(cm, 'id')) continue;
    visit(cm);
    for (const list of elements(cm))
      if (list.ns === COMMENTS.p188 && list.local === 'replyLst')
        for (const r of elements(list))
          if (r.ns === COMMENTS.p188 && r.local === 'reply' && attr(r, 'id')) visit(r);
  }
}

function modernComment(
  slidePart: string,
  el: Element,
  authors: Map<string, CommentAuthor>,
  count: GuidCount
): SlideComment {
  const author = authors.get(attr(el, 'authorId')?.toUpperCase() ?? '');
  const guid = (attr(el, 'id') ?? '').toUpperCase();
  const mark = modernMark(el);
  const key = `${guid}\u0000${mark}`;
  const ordinal = bump(count.seen, key);
  const repeated = (count.uses.get(guid) ?? 0) > 1 || guid.includes('~');
  const text = modernText(el);
  // The GUID and the mark, which edits, resolves, and deletes of other comments keep. Twins (one
  // GUID, author, and date) also take a hash of their text, and twins with one text their place
  // among each other: nothing else in the file tells them apart. `CommentState` carries the
  // identity of a twin across its own edits and the deletes of other twins.
  const hash = fnv(text);
  const same = `${key}\u0000${hash}`;
  const place = bump(count.seenText, same);
  const twins = (count.twins.get(key) ?? 0) > 1;
  const suffix = twins ? `~${hash}${(count.texts.get(same) ?? 0) > 1 ? `~${place}` : ''}` : '';
  const identity = `${slidePart}#m~${identityField(guid)}~${mark}${suffix}`;
  return {
    id: modernId(slidePart, guid, repeated ? { mark, ordinal } : undefined),
    identity,
    author: author?.name ?? '',
    initials: author?.initials ?? '',
    ...optionalDate(attr(el, 'created')),
    text,
    textEditable: modernTextEditable(el),
  };
}

/**
 * The group of a comment `identity`: the part that names its GUID and mark, or its legacy author
 * and index. Comments of one group are twins.
 */
export function identityGroup(identity: string): string {
  const cut = identity.lastIndexOf('#');
  return (
    identity.slice(0, cut + 1) +
    identity
      .slice(cut + 1)
      .split('~')
      .slice(0, 3)
      .join('~')
  );
}

/** The threaded comments of one list. */
function modernThreads(
  pkg: PptxPackage,
  slidePart: string,
  listPart: string,
  keys: ReturnType<typeof shapeKeys>,
  count: GuidCount
): CommentThread[] {
  const root = commentTree(pkg, listPart)?.root;
  if (root?.ns !== COMMENTS.p188 || root.local !== 'cmLst') return [];
  const authors = modernAuthors(pkg);
  const out: CommentThread[] = [];
  for (const cm of elements(root)) {
    if (cm.ns !== COMMENTS.p188 || cm.local !== 'cm' || !attr(cm, 'id')) continue;
    // The thread before its replies: `findModern` counts repeats in this order.
    const head = modernComment(slidePart, cm, authors, count);
    // Every reply list: two collaborators who give a thread its first reply at once give it two
    // until a save joins them.
    const replies = elements(cm)
      .filter((l) => l.ns === COMMENTS.p188 && l.local === 'replyLst')
      .flatMap((l) => elements(l))
      .filter((r) => r.ns === COMMENTS.p188 && r.local === 'reply' && attr(r, 'id'))
      .map((r) => modernComment(slidePart, r, authors, count));
    const status = attr(cm, 'status');
    out.push({
      ...head,
      format: 'modern',
      resolved: status === 'resolved' || status === 'closed',
      anchor: modernAnchor(cm, keys),
      replies,
    });
  }
  return out;
}

/** The legacy comments of one list, with replies under their parent. */
function legacyThreads(
  pkg: PptxPackage,
  slidePart: string,
  listPart: string,
  count: GuidCount
): CommentThread[] {
  const root = commentTree(pkg, listPart)?.root;
  if (root?.ns !== NS.p || root.local !== 'cmLst') return [];
  const authors = legacyAuthors(pkg);
  const threads: { thread: CommentThread; replies: SlideComment[] }[] = [];
  const byKey = new Map<string, { replies: SlideComment[] }>();
  const pending: { parent: string; reply: SlideComment }[] = [];
  for (const cm of elements(root)) {
    if (cm.ns !== NS.p || cm.local !== 'cm') continue;
    const authorId = canonicalAuthor(attr(cm, 'authorId'));
    const idx = canonicalIdx(attr(cm, 'idx'));
    const author = authors.get(authorId);
    // A file can repeat an author and index: each repeat also takes its place.
    const key = `${authorId}\u0000${idx}`;
    const place = bump(count.seenLegacy, key);
    const base = `${slidePart}#l~${identityField(authorId)}~${identityField(idx)}`;
    const comment: SlideComment = {
      id: legacyId(slidePart, authorId, idx),
      identity: (count.legacy.get(key) ?? 0) > 1 ? `${base}~${place}` : base,
      author: author?.name ?? '',
      initials: author?.initials ?? '',
      ...optionalDate(attr(cm, 'dt')),
      text: textOfPlain(child(cm, 'p:text')),
      textEditable: !!child(cm, 'p:text'),
    };
    const parent = descendant(child(cm, 'p:extLst'), NS.p15, 'parentCm', 4);
    if (parent) {
      const key = `${canonicalAuthor(attr(parent, 'authorId'))}:${canonicalIdx(attr(parent, 'idx'))}`;
      pending.push({ parent: key, reply: comment });
      continue;
    }
    const replies: SlideComment[] = [];
    const thread: CommentThread = {
      ...comment,
      format: 'legacy',
      resolved: false,
      anchor: pointOf(child(cm, 'p:pos'), LEGACY_POS_EMU),
      replies,
    };
    threads.push({ thread, replies });
    if (!byKey.has(`${authorId}:${idx}`)) byKey.set(`${authorId}:${idx}`, { replies });
  }
  // A reply whose parent is not in the list shows as a thread of its own.
  for (const { parent, reply } of pending) {
    const owner = byKey.get(parent);
    if (owner) owner.replies.push(reply);
    else
      threads.push({
        thread: {
          ...reply,
          format: 'legacy',
          resolved: false,
          anchor: { kind: 'slide' },
          replies: [],
        },
        replies: [],
      });
  }
  return threads.map((t) => t.thread);
}

function textOfPlain(el: Element | undefined): string {
  let out = '';
  for (const c of el?.children ?? []) if (c.type === 'text') out += c.value;
  return out;
}

/** Every comment thread of a slide: its legacy list first, then its threaded lists. */
export function readSlideComments(pkg: PptxPackage, slidePart: string): CommentThread[] {
  const { legacy, modern } = slideCommentParts(pkg, slidePart);
  if (legacy.length === 0 && modern.length === 0) return [];
  const keys = shapeKeys(xmlPartOf(pkg, slidePart)?.root);
  const count: GuidCount = {
    uses: new Map(),
    twins: new Map(),
    texts: new Map(),
    seen: new Map(),
    seenText: new Map(),
    legacy: new Map(),
    seenLegacy: new Map(),
  };
  for (const part of modern)
    eachModern(commentTree(pkg, part)?.root, (el) => {
      const guid = (attr(el, 'id') ?? '').toUpperCase();
      const key = `${guid}\u0000${modernMark(el)}`;
      bump(count.uses, guid);
      bump(count.twins, key);
      bump(count.texts, `${key}\u0000${fnv(modernText(el))}`);
    });
  for (const part of legacy) {
    const root = commentTree(pkg, part)?.root;
    if (root?.ns !== NS.p || root.local !== 'cmLst') continue;
    for (const cm of elements(root))
      if (cm.ns === NS.p && cm.local === 'cm')
        bump(
          count.legacy,
          `${canonicalAuthor(attr(cm, 'authorId'))}\u0000${canonicalIdx(attr(cm, 'idx'))}`
        );
  }
  return [
    ...legacy.flatMap((part) => legacyThreads(pkg, slidePart, part, count)),
    ...modern.flatMap((part) => modernThreads(pkg, slidePart, part, keys, count)),
  ];
}
