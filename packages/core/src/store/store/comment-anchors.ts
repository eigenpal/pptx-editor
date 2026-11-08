// Threaded comments that point at text follow their text through text edits.
//
// A threaded comment can point at a range of a shape's text: a text moniker list
// (`ac:txMkLst`) names the shape (`ac:spMk`) and then the range, `ac:txMk cp len`, as the first
// character and the number of characters. Paragraphs count one character between them, as the
// editor counts them (`automation/commands/comment-text-range.ts`). An edit before the range
// moves it, an edit in it grows or shrinks it, and an edit that removes all of its text leaves
// it with no length: the reader then points the thread at the shape.
//
// Each local transaction that changes the text of a shape with such a comment is checked here.
// The changed regions between the old and the new text of the shape (`diffHunks`) tell where
// each end of the range went. Text that an edit adds at either end of the range stays out of
// it. The editor names where the edit started (`TextEditAt`), and an insertion or deletion that
// fits in more than one place goes to the place nearest it, as in `paragraph-ranges.ts`.
//
// The fixes are ops of the same transaction, so undo restores the range with the text, and
// collaborators receive them with the edit. A moved range drops its `ac:context`, the check of
// the text that it was made on: the check no longer describes the text.

import { NS } from '../package/namespaces.ts';
import {
  decodeXml,
  relationshipsOf,
  type PackagePart,
  type PptxPackage,
} from '../package/package.ts';
import { attr, child, elements } from '../package/query.ts';
import { diffResult, type TextHunk } from '../package/text-diff.ts';
import { parseXmlPart, type Element } from '../package/tree.ts';
import { serializeXmlPart } from '../package/serialize.ts';
import type { PptxOp } from './ops.ts';
import { slid, type TextEditAt } from './paragraph-ranges.ts';

const P188 = 'http://schemas.microsoft.com/office/powerpoint/2018/8/main';
const AC = 'http://schemas.microsoft.com/office/drawing/2013/main/command';
const COMMENTS_REL = 'http://schemas.microsoft.com/office/2018/10/relationships/comments';
const MAX_DEPTH = 64;

/** The tree of a comment part, once per part object: the reader keeps comment parts as bytes. */
const trees = new WeakMap<PackagePart, Element | null>();

function treeOf(part: PackagePart | undefined): Element | undefined {
  if (!part) return undefined;
  if (part.kind === 'xml') return part.xml.root;
  let root = trees.get(part);
  if (root === undefined) {
    const parsed = parseXmlPart(part.partName, decodeXml(part.bytes));
    root = parsed.ok ? parsed.part.root : null;
    trees.set(part, root);
  }
  return root ?? undefined;
}

/** A text moniker of a thread: the moniker list's shape keys and the `ac:txMk` element. */
interface TextMark {
  /** The thread's GUID, upper case. */
  readonly thread: string;
  readonly id: string | undefined;
  readonly creationId: string | undefined;
  readonly txMk: Element;
}

/** The text monikers of a threaded comment list, in document order. */
function textMarks(root: Element | undefined): TextMark[] {
  if (root?.ns !== P188 || root.local !== 'cmLst') return [];
  const out: TextMark[] = [];
  for (const cm of elements(root)) {
    const thread = attr(cm, 'id')?.toUpperCase();
    if (cm.ns !== P188 || cm.local !== 'cm' || !thread) continue;
    for (const list of elements(cm)) {
      if (list.ns !== AC || list.local !== 'txMkLst') continue;
      const marks = elements(list).filter((m) => m.ns === AC && m.local.endsWith('Mk'));
      const txMk = marks[marks.length - 1];
      const shape = marks[marks.length - 2];
      if (txMk?.local !== 'txMk' || shape?.local !== 'spMk') continue;
      out.push({
        thread,
        id: attr(shape, 'id'),
        creationId: attr(shape, 'creationId')?.toUpperCase(),
        txMk,
      });
    }
  }
  return out;
}

/** A count of `xsd:unsignedInt`, or undefined. */
function count(value: string | undefined): number | undefined {
  return value !== undefined && /^\s*\d{1,9}\s*$/.test(value) ? Number(value.trim()) : undefined;
}

/** The shapes of a slide by `p:cNvPr/@id` and by `a16:creationId`, at any depth of groups. */
function shapesOf(root: Element): { byId: Map<string, Element>; byCreation: Map<string, Element> } {
  const byId = new Map<string, Element>();
  const byCreation = new Map<string, Element>();
  const walk = (el: Element, depth: number): void => {
    if (depth > MAX_DEPTH) return;
    for (const c of elements(el)) {
      if (c.kind === 'sp') {
        const cNvPr = child(child(c, 'p:nvSpPr'), 'p:cNvPr');
        const id = attr(cNvPr, 'id');
        if (id !== undefined && !byId.has(id)) byId.set(id, c);
        for (const ext of elements(child(cNvPr, 'a:extLst')))
          for (const e of elements(ext))
            if (e.ns === NS.a16 && e.local === 'creationId') {
              const key = attr(e, 'id')?.toUpperCase();
              if (key && !byCreation.has(key)) byCreation.set(key, c);
            }
      } else if (c.kind === 'spTree' || c.kind === 'grpSp' || c.local === 'cSld')
        walk(c, depth + 1);
    }
  };
  walk(root, 0);
  return { byId, byCreation };
}

/**
 * The text of paragraph `p` of slide `part`, as the editor counts it, so that a range moves in
 * the units it was written in. The store has no math or field model; the editor names its own
 * (`DocumentStoreOptions.paragraphText`).
 */
export type ParagraphText = (p: Element, pkg: PptxPackage, part: string) => string;

/**
 * The text of a paragraph without the editor's count: each run and field as stored, and one
 * character for each line break.
 */
export const storedParagraphText: ParagraphText = (p) => {
  let text = '';
  for (const c of elements(p)) {
    if (c.kind === 'break') text += '\u000b';
    else if (c.kind === 'run' || c.kind === 'field')
      for (const t of elements(c)) if (t.kind === 'text') text += textOfNode(t);
  }
  return text;
};

/** The paragraph texts of a shape of slide `part` in `pkg`. */
function paragraphTexts(
  shape: Element,
  pkg: PptxPackage,
  part: string,
  textOf: ParagraphText
): string[] {
  const out: string[] = [];
  for (const p of elements(child(shape, 'p:txBody')))
    if (p.kind === 'paragraph') out.push(textOf(p, pkg, part));
  return out;
}

function textOfNode(el: Element): string {
  let out = '';
  for (const c of el.children) if (c.type === 'text') out += c.value;
  return out;
}

/** Where an old boundary `p` went. `after`: text inserted at `p` goes before it. */
function mapped(p: number, hunks: readonly TextHunk[], after: boolean): number {
  let delta = 0;
  for (const h of hunks) {
    if (p < h.aStart) break;
    if (p === h.aStart && h.aStart === h.aEnd) return after ? h.bEnd : h.bStart;
    if (p === h.aStart) return h.bStart;
    if (p < h.aEnd) return after ? h.bStart : h.bEnd;
    delta = h.bEnd - h.aEnd;
  }
  return p + delta;
}

/** The point of an edit in old paragraph texts joined by line feeds. */
function charAt(texts: readonly string[], at: { paragraph: number; offset: number }): number {
  let c = 0;
  for (let i = 0; i < at.paragraph && i < texts.length; i += 1) c += texts[i]!.length + 1;
  return c + Math.max(0, Math.min(at.offset, texts[at.paragraph]?.length ?? 0));
}

/**
 * The ops that keep the text ranges of the threaded comments on slide `part` on their text,
 * after an edit changed the slide's root from `before` to `after` (in `prior` and `next`). `at`
 * is where a text edit started, and `textOf` counts the text as the editor does. Empty when
 * nothing needs to change.
 */
export function commentAnchorFixes(
  prior: PptxPackage,
  next: PptxPackage,
  part: string,
  before: Element,
  after: Element,
  at?: TextEditAt,
  textOf: ParagraphText = storedParagraphText
): PptxOp[] {
  if (after.ns !== NS.p || after.local !== 'sld') return [];
  const lists = relationshipsOf(next, part).filter((r) => r.type === COMMENTS_REL && !r.external);
  if (lists.length === 0) return [];
  const ops: PptxOp[] = [];
  let old: ReturnType<typeof shapesOf> | undefined;
  let now: ReturnType<typeof shapesOf> | undefined;
  const changes = new Map<string, { hunks: TextHunk[]; exact: boolean } | null>();
  for (const rel of lists) {
    const name = rel.target;
    if (!name) continue;
    const listPart = next.parts.get(name);
    const marks = textMarks(treeOf(listPart));
    if (marks.length === 0) continue;
    // The range of a thread as it was before the edit: a range that the same transaction added
    // or changed names the new text already.
    const was = new Map(
      textMarks(treeOf(prior.parts.get(name))).map((m) => [
        m.thread,
        `${count(attr(m.txMk, 'cp'))} ${count(attr(m.txMk, 'len'))}`,
      ])
    );
    const edits: { mark: TextMark; cp: number; len: number }[] = [];
    for (const mark of marks) {
      const cp = count(attr(mark.txMk, 'cp'));
      const len = count(attr(mark.txMk, 'len'));
      if (cp === undefined || len === undefined) continue;
      if (was.get(mark.thread) !== `${cp} ${len}`) continue;
      const key = `${mark.creationId ?? ''} ${mark.id ?? ''}`;
      let change = changes.get(key);
      if (change === undefined) {
        old ??= shapesOf(before);
        now ??= shapesOf(after);
        const find = (s: ReturnType<typeof shapesOf>) =>
          (mark.creationId ? s.byCreation.get(mark.creationId) : undefined) ??
          (mark.id !== undefined ? s.byId.get(mark.id) : undefined);
        const a = find(old);
        const b = find(now);
        // An element that the edit did not touch keeps its object.
        const was = a && b && a !== b ? paragraphTexts(a, prior, part, textOf) : undefined;
        const aText = was?.join('\n');
        const bText = was && paragraphTexts(b!, next, part, textOf).join('\n');
        if (!was || bText === undefined || aText === bText) change = null;
        else {
          const hint = at?.part === part && at.node === b!.id ? charAt(was, at) : undefined;
          const r = diffResult(aText!, bText);
          change = {
            hunks: slid(aText!, bText, r.hunks, r.exact ? hint : undefined),
            exact: r.exact,
          };
        }
        changes.set(key, change);
      }
      if (!change) continue;
      const start = mapped(cp, change.hunks, true);
      let end = Math.max(start, mapped(cp + len, change.hunks, false));
      // Too many changes for a character diff: one region covers the whole changed middle, and
      // a range in it cannot tell where its text went. It points at the shape.
      if (!change.exact && change.hunks.some((h) => cp < h.aEnd && cp + len > h.aStart))
        end = start;
      if (start === cp && end - start === len) continue;
      edits.push({ mark, cp: start, len: end - start });
    }
    if (edits.length > 0) for (const op of writes(name, listPart!, edits)) ops.push(op);
  }
  return ops;
}

/**
 * The ops that write `edits` into the comment part `name`: attribute writes into a part the
 * store holds as a tree, or the part again, as XML, when it holds the part as bytes.
 */
function writes(
  name: string,
  listPart: PackagePart,
  edits: readonly { mark: TextMark; cp: number; len: number }[]
): PptxOp[] {
  const attrOp = (node: string, local: string, value: number): PptxOp => ({
    op: 'setAttribute',
    part: name,
    node,
    name: { ns: '', local },
    value: String(value),
  });
  if (listPart.kind === 'xml') {
    const out: PptxOp[] = [];
    for (const { mark, cp, len } of edits) {
      out.push(attrOp(mark.txMk.id, 'cp', cp), attrOp(mark.txMk.id, 'len', len));
      for (const c of elements(mark.txMk))
        if (c.ns === AC && c.local === 'context')
          out.push({ op: 'removeNode', part: name, node: c.id });
    }
    return out;
  }
  // A part kept as bytes: a fresh tree, changed and written back whole.
  const parsed = parseXmlPart(name, decodeXml(listPart.bytes));
  if (!parsed.ok) return [];
  const fresh = textMarks(parsed.part.root);
  for (const { mark, cp, len } of edits) {
    const twin = fresh.find(
      (m) => m.thread === mark.thread && attr(m.txMk, 'cp') === attr(mark.txMk, 'cp')
    );
    if (!twin) continue;
    const el = writable(twin.txMk);
    el.attrs = el.attrs.map((a) =>
      a.ns === '' && (a.local === 'cp' || a.local === 'len')
        ? { ...a, value: String(a.local === 'cp' ? cp : len) }
        : a
    );
    el.children = el.children.filter(
      (c) => c.type !== 'element' || c.ns !== AC || c.local !== 'context'
    );
  }
  return [
    { op: 'removePart', part: name },
    { op: 'addPart', part: name, xml: serializeXmlPart(parsed.part) },
  ];
}

/** An element of a tree that this module parsed and that nothing else holds. */
const writable = (el: Element) => el as { -readonly [K in keyof Element]: Element[K] };
