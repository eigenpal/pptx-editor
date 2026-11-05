// The speaker notes field as DOM: an editable element (`contenteditable="plaintext-only"`) that
// shows the notes runs with their formatting and maps DOM points to plain-text offsets, the
// offsets `setNotes` and `formatNotes` take. Framework-free, so the React and Vue adapters share
// it. Paragraphs are `\n` text between runs; nothing here parses HTML, and every text goes in
// through `textContent` or text nodes.

import type { NotesFormat, NotesParagraph, NotesRun } from '../contracts/index.ts';
import { cleanText } from '../automation/commands/text-units.ts';

const FORMAT_CLASS: Readonly<Record<NotesFormat, string>> = {
  bold: 'pptx-notes__run--bold',
  italic: 'pptx-notes__run--italic',
  underline: 'pptx-notes__run--underline',
  strike: 'pptx-notes__run--strike',
};
const FORMATS = Object.keys(FORMAT_CLASS) as NotesFormat[];

/**
 * Draw `paragraphs` into `root`: one span per run, `\n` between paragraphs, and an end marker
 * after a final empty line so the caret can stand there. @internal
 */
export function renderNotes(root: HTMLElement, paragraphs: readonly NotesParagraph[]): void {
  const doc = root.ownerDocument;
  // A fragment, not an argument list: the run count comes from the file.
  const nodes = doc.createDocumentFragment();
  paragraphs.forEach((p, i) => {
    if (i > 0) nodes.append(doc.createTextNode('\n'));
    for (const r of p.runs) {
      if (!r.text) continue;
      const span = doc.createElement('span');
      const classes = FORMATS.filter((f) => r[f]).map((f) => FORMAT_CLASS[f]);
      if (classes.length) span.className = classes.join(' ');
      span.textContent = r.text;
      nodes.append(span);
    }
  });
  if (paragraphs.length > 1 && !paragraphs[paragraphs.length - 1]!.runs.some((r) => r.text)) {
    const end = doc.createElement('br');
    end.dataset.end = '';
    nodes.append(end);
  }
  root.replaceChildren(nodes);
}

/** A walk of the field: each text node and each line break, with its plain-text start. */
interface Segment {
  readonly node: Node;
  readonly start: number;
  readonly length: number;
}

const BLOCKS: ReadonlySet<string> = new Set(['DIV', 'P']);

/** Whether `node` is the last node of `root` in document order. */
function isLastNode(root: Node, node: Node): boolean {
  for (let n: Node | null = node; n && n !== root; n = n.parentNode)
    if (n.nextSibling) return false;
  return true;
}

/** The field's text and segments. Browsers may add `<br>` or `<div>` lines while typing. */
function walk(root: HTMLElement): { text: string; segments: Segment[]; ends: Map<Node, number> } {
  let text = '';
  const segments: Segment[] = [];
  const ends = new Map<Node, number>();
  const visit = (node: Node, depth: number): void => {
    if (depth > 64) return;
    if (node.nodeType === 3) {
      const data = (node.nodeValue ?? '').replace(/\r/g, '');
      segments.push({ node, start: text.length, length: data.length });
      text += data;
      return;
    }
    if (node.nodeType !== 1) return;
    const el = node as HTMLElement;
    if (el.tagName === 'BR') {
      // A browser keeps a `<br>` as the placeholder of an emptied line; at the very end it is
      // no line of its own.
      if (!('end' in el.dataset) && !isLastNode(root, el)) {
        segments.push({ node, start: text.length, length: 1 });
        text += '\n';
      }
      return;
    }
    if (el !== root && BLOCKS.has(el.tagName) && text && !text.endsWith('\n')) text += '\n';
    for (const c of el.childNodes) visit(c, depth + 1);
    ends.set(el, text.length);
  };
  visit(root, 0);
  return { text, segments, ends };
}

/** The plain text of the field (paragraphs and line breaks as `\n`). @internal */
export function notesFieldText(root: HTMLElement): string {
  return walk(root).text;
}

/** The plain-text offset of a DOM point inside `root`. @internal */
export function notesOffsetOf(root: HTMLElement, node: Node, offset: number): number {
  const { text, segments, ends } = walk(root);
  if (node.nodeType === 3) {
    const s = segments.find((g) => g.node === node);
    return s ? s.start + Math.min(offset, s.length) : text.length;
  }
  // An element point: the start of its child `offset`, or the element's end.
  const next = node.childNodes[offset];
  if (!next) return ends.get(node) ?? text.length;
  const first = (n: Node): number | undefined => {
    const s = segments.find((g) => g.node === n || n.contains(g.node));
    return s?.start;
  };
  return first(next) ?? ends.get(next) ?? ends.get(node) ?? text.length;
}

/** The DOM point of a plain-text offset inside `root`. @internal */
export function notesPointAt(root: HTMLElement, offset: number): { node: Node; offset: number } {
  const { segments } = walk(root);
  for (const s of segments) {
    if (offset > s.start + s.length) continue;
    if (s.node.nodeType === 3) return { node: s.node, offset: offset - s.start };
    // A `<br>`: before it, or after it.
    const parent = s.node.parentNode!;
    const index = [...parent.childNodes].indexOf(s.node as ChildNode);
    return { node: parent, offset: offset > s.start ? index + 1 : index };
  }
  return { node: root, offset: root.childNodes.length };
}

/** The selected plain-text range inside `root`, or null when the selection is elsewhere. @internal */
export function notesSelection(root: HTMLElement): { start: number; end: number } | null {
  const sel = root.ownerDocument.getSelection();
  if (!sel || sel.rangeCount === 0) return null;
  const r = sel.getRangeAt(0);
  if (!root.contains(r.startContainer) || !root.contains(r.endContainer)) return null;
  const a = notesOffsetOf(root, r.startContainer, r.startOffset);
  const b = notesOffsetOf(root, r.endContainer, r.endOffset);
  return { start: Math.min(a, b), end: Math.max(a, b) };
}

/**
 * Offset `offset` of the field's text `text`, counted in the text `setNotes` writes: `setNotes`
 * drops characters a slide cannot hold (`cleanText`, such as pasted C0 controls), so a later
 * offset moves back by the characters dropped before it. @internal
 */
export function notesCommittedOffset(text: string, offset: number): number {
  return cleanText(text.slice(0, offset)).length;
}

/** Select the plain-text range `[start, end)` inside `root`. @internal */
export function selectNotes(root: HTMLElement, start: number, end = start): void {
  const sel = root.ownerDocument.getSelection();
  if (!sel) return;
  const a = notesPointAt(root, start);
  const b = notesPointAt(root, end);
  const range = root.ownerDocument.createRange();
  range.setStart(a.node, a.offset);
  range.setEnd(b.node, b.offset);
  sel.removeAllRanges();
  sel.addRange(range);
}

/**
 * Replace the selection with `text`, as typing does, and put the caret after it. For the keys
 * the field handles itself (Enter) so every browser writes the same `\n`. @internal
 */
export function insertNotesText(root: HTMLElement, text: string): void {
  const range = notesSelection(root);
  const at = range ?? { start: notesFieldText(root).length, end: notesFieldText(root).length };
  const sel = root.ownerDocument.getSelection()!;
  selectNotes(root, at.start, at.end);
  const r = sel.getRangeAt(0);
  r.deleteContents();
  const node = root.ownerDocument.createTextNode(text);
  r.insertNode(node);
  // A final line break needs something after it to show the new line. `insertNode` splits the
  // text node it lands in, so an empty text node can follow: test for content, not for a node.
  const rest = root.ownerDocument.createRange();
  rest.setStartAfter(node);
  rest.setEnd(root, root.childNodes.length);
  const follows = rest.toString() !== '' || !!rest.cloneContents().querySelector('br');
  if (text.endsWith('\n') && !follows) {
    const end = root.ownerDocument.createElement('br');
    end.dataset.end = '';
    node.after(end);
  }
  root.normalize();
  selectNotes(root, at.start + text.length);
}

/**
 * The formats that the selected text has throughout: what the notes B, I, and U buttons show
 * pressed. A caret reads the character before it, as typing would take; at the start of a
 * paragraph it reads the character after it, the first one the paragraph holds. @internal
 */
export function notesFormatsAt(
  root: HTMLElement,
  range: { start: number; end: number }
): Record<NotesFormat, boolean> {
  const arm = armed.get(root);
  if (arm && range.start === range.end && range.start === arm.at) return { ...arm.formats };
  const { segments, text } = walk(root);
  const startOfParagraph = range.start === 0 || text[range.start - 1] === '\n';
  const [from, to] =
    range.start !== range.end
      ? [range.start, range.end]
      : startOfParagraph
        ? [range.start, range.start + 1]
        : [range.start - 1, range.start];
  // Paragraph breaks between runs carry no format of their own.
  const covered = segments.filter(
    (s) =>
      s.node.nodeType === 3 &&
      /[^\n]/.test(s.node.nodeValue ?? '') &&
      s.start < to &&
      s.start + s.length > from
  );
  const out = { bold: false, italic: false, underline: false, strike: false };
  if (!covered.length) return out;
  for (const f of FORMATS)
    out[f] = covered.every(
      (s) => !!s.node.parentElement?.closest(`.${FORMAT_CLASS[f]}`) && root.contains(s.node)
    );
  return out;
}

type NotesFormats = Record<NotesFormat, boolean>;

/** The formats of a text node of the field: the format classes of the spans around it. */
function formatsOfNode(node: Node): NotesFormats {
  const out = { bold: false, italic: false, underline: false, strike: false };
  for (const f of FORMATS) out[f] = !!node.parentElement?.closest(`.${FORMAT_CLASS[f]}`);
  return out;
}

const NONE: NotesFormats = { bold: false, italic: false, underline: false, strike: false };

/**
 * The field's text as formatted paragraphs, as the field shows them while the user types: text
 * typed into a bold run shows bold, so it is bold here. A commit passes them to `setNotes`
 * (`paragraphs`), so what was shown is what is stored. @internal
 */
export function notesFieldParagraphs(root: HTMLElement): NotesParagraph[] {
  const { text, segments } = walk(root);
  const marks: (NotesFormats | undefined)[] = [];
  for (const s of segments) {
    if (s.node.nodeType !== 3) continue;
    const formats = formatsOfNode(s.node);
    for (let i = s.start; i < s.start + s.length; i += 1) marks[i] = formats;
  }
  const paragraphs: NotesParagraph[] = [];
  let runs: NotesRun[] = [];
  for (let i = 0; i <= text.length; i += 1) {
    if (i === text.length || text[i] === '\n') {
      paragraphs.push({ runs });
      runs = [];
      continue;
    }
    const f = marks[i] ?? NONE;
    const last = runs[runs.length - 1];
    if (last && FORMATS.every((k) => last[k] === f[k]))
      runs[runs.length - 1] = { ...last, text: last.text + text[i] };
    else runs.push({ text: text[i]!, ...f });
  }
  return paragraphs;
}

/** A format armed at a caret: the formats the next typed text takes, and where. */
const armed = new WeakMap<HTMLElement, { readonly at: number; readonly formats: NotesFormats }>();

/**
 * Turn `format` on or off for the next text typed at the caret, as the canvas does with no
 * text selected: the B, I, and U buttons show it pressed, and `typeArmedNotes` writes the next
 * typed text with it. A caret move drops it. Returns false without a caret in the field.
 * @internal
 */
export function armNotesFormat(root: HTMLElement, format: NotesFormat): boolean {
  const range = notesSelection(root);
  if (!range || range.start !== range.end) return false;
  const now = notesFormatsAt(root, range);
  armed.set(root, { at: range.start, formats: { ...now, [format]: !now[format] } });
  return true;
}

/** Drop a format armed at a caret that moved away from it. @internal */
export function followNotesArm(root: HTMLElement): void {
  const arm = armed.get(root);
  const range = arm ? notesSelection(root) : null;
  if (arm && (!range || range.start !== arm.at || range.end !== arm.at)) armed.delete(root);
}

/**
 * Write `data`, typed at the caret, with the formats armed there (`armNotesFormat`), in a run
 * of its own, and put the caret after it, so the next keys type into that run. Returns false,
 * and writes nothing, when no format is armed at the caret: the browser types as usual.
 * @internal
 */
export function typeArmedNotes(root: HTMLElement, data: string): boolean {
  followNotesArm(root);
  const arm = armed.get(root);
  if (!arm || !data) return false;
  armed.delete(root);
  const doc = root.ownerDocument;
  const span = doc.createElement('span');
  const classes = FORMATS.filter((f) => arm.formats[f]).map((f) => FORMAT_CLASS[f]);
  if (classes.length) span.className = classes.join(' ');
  span.textContent = data;
  const point = notesPointAt(root, arm.at);
  // The run goes between the field's own runs: a run that holds the caret splits there.
  let top: Node | null = point.node === root ? null : point.node;
  while (top && top.parentNode !== root) top = top.parentNode;
  if (!top) root.insertBefore(span, root.childNodes[point.offset] ?? null);
  else {
    const rest = doc.createRange();
    rest.setStart(point.node, point.offset);
    rest.setEndAfter(top);
    const tail = rest.extractContents();
    root.insertBefore(span, top.nextSibling);
    root.insertBefore(tail, span.nextSibling);
  }
  const sel = doc.getSelection();
  const text = span.firstChild!;
  sel?.removeAllRanges();
  const caret = doc.createRange();
  caret.setStart(text, data.length);
  sel?.addRange(caret);
  return true;
}

/**
 * Write `data`, typed at the start of a paragraph, into the paragraph's first run, so it shows
 * the format that the run has, as the B, I, and U buttons say (`notesFormatsAt`). A browser
 * puts it into the line break before the run otherwise, where it shows no format. Returns
 * false, and writes nothing, when the caret is not at such a place. @internal
 */
export function typeAtParagraphStart(root: HTMLElement, data: string): boolean {
  const range = notesSelection(root);
  const sel = root.ownerDocument.getSelection();
  if (!data || !range || range.start !== range.end || !sel?.anchorNode) return false;
  const { text, segments } = walk(root);
  const at = range.start;
  if ((at > 0 && text[at - 1] !== '\n') || at >= text.length || text[at] === '\n') return false;
  const next = segments.find(
    (s) => s.node.nodeType === 3 && s.start <= at && at < s.start + s.length
  );
  if (!next || next.node === sel.anchorNode || !next.node.parentElement?.closest('span'))
    return false;
  const node = next.node as Text;
  const offset = at - next.start;
  node.insertData(offset, data);
  const caret = root.ownerDocument.createRange();
  caret.setStart(node, offset + data.length);
  sel.removeAllRanges();
  sel.addRange(caret);
  return true;
}
