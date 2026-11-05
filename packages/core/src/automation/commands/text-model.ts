// Text editing on the canonical tree. A text body is a list of paragraphs; a paragraph is a list
// of inline items (runs, breaks, fields) addressed by UTF-16 offset. Every edit here compiles to
// store ops through an OpBatch, so typing is as serializable and undoable as any other command.
//
// Fields (`a:fld`) are atomic: an edit never splits one. A point inside a field snaps to its end.
// A field is as long as the text layout shows: a `slidenum` field on a slide shows the slide
// number (the cascade's rule), so offsets here and caret offsets from layout agree.

import {
  attr,
  child,
  elements,
  NS,
  presentationInfo,
  specOf,
  textOf,
  type Element,
  type NodeSpec,
} from '../../store/index.ts';
import type { TextPoint } from '../../contracts/neutral.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { caretWordRange } from '@eigenpal/ooxml-core';

/** Slide numbers by part, per batch: slide order does not change inside a text edit. */
const slideNumbers = new WeakMap<OpBatch, Map<string, number | undefined>>();

/**
 * The number a `slidenum` field shows in `part` (1-based slide index, as the cascade resolves
 * it), or undefined for a part that is not a slide (notes, layouts) where the stored text shows.
 */
export function slideNumberOf(batch: OpBatch, part: string): number | undefined {
  let byPart = slideNumbers.get(batch);
  if (!byPart) slideNumbers.set(batch, (byPart = new Map()));
  if (!byPart.has(part)) {
    const index = presentationInfo(batch.pkg).slides.findIndex((s) => s.partName === part);
    byPart.set(part, index < 0 ? undefined : index + 1);
  }
  return byPart.get(part);
}

/** The text an inline item shows: its `a:t`, or the slide number for a `slidenum` field. */
function slotText(el: Element, slideNumber: number | undefined): string {
  if (el.kind === 'field' && slideNumber !== undefined && attr(el, 'type') === 'slidenum')
    return String(slideNumber);
  return textOf(child(el, 'a:t'));
}

const isHigh = (c: number): boolean => c >= 0xd800 && c <= 0xdbff;
const isLow = (c: number): boolean => c >= 0xdc00 && c <= 0xdfff;

/** True when `offset` falls between the two halves of a surrogate pair in `text`. */
export function splitsPair(text: string, offset: number): boolean {
  return (
    offset > 0 &&
    offset < text.length &&
    isHigh(text.charCodeAt(offset - 1)) &&
    isLow(text.charCodeAt(offset))
  );
}

/** One inline item with its offset range in the paragraph. */
export interface Slot {
  readonly el: Element;
  readonly kind: 'run' | 'break' | 'field';
  readonly start: number;
  readonly end: number;
  readonly text: string;
}

const A = NS.a;
const aSpec = (
  local: string,
  attrs: Record<string, string> = {},
  children: NodeSpec['children'] = []
): NodeSpec => ({
  ns: A,
  local,
  prefix: 'a',
  attrs: Object.entries(attrs).map(([k, v]) => ({ ns: '', local: k, value: v })),
  children,
});

/** The paragraphs of a text body. */
export function paragraphsOf(txBody: Element): Element[] {
  return elements(txBody).filter((e) => e.kind === 'paragraph');
}

/**
 * Inline items of a paragraph with offsets. Pass the slide number for a paragraph on a slide,
 * so a `slidenum` field is as long as the number layout shows.
 */
export function slotsOf(p: Element, slideNumber?: number): Slot[] {
  const out: Slot[] = [];
  let at = 0;
  for (const el of elements(p)) {
    if (el.kind === 'run' || el.kind === 'field') {
      const text = slotText(el, slideNumber);
      out.push({
        el,
        kind: el.kind === 'run' ? 'run' : 'field',
        start: at,
        end: at + text.length,
        text,
      });
      at += text.length;
    } else if (el.kind === 'break') {
      out.push({ el, kind: 'break', start: at, end: at + 1, text: '\n' });
      at += 1;
    }
  }
  return out;
}

export function paragraphLength(p: Element, slideNumber?: number): number {
  const slots = slotsOf(p, slideNumber);
  return slots.length ? slots[slots.length - 1].end : 0;
}

/** Plain text of a paragraph. */
export function paragraphText(p: Element, slideNumber?: number): string {
  return slotsOf(p, slideNumber)
    .map((s) => s.text)
    .join('');
}

/** Snap a point out of the middle of a field or of a surrogate pair, to its end. */
export function snapPoint(p: Element, offset: number, slideNumber?: number): number {
  for (const s of slotsOf(p, slideNumber)) {
    if (offset <= s.start || offset >= s.end) continue;
    if (s.kind === 'field') return s.end;
    if (splitsPair(s.text, offset - s.start)) return offset + 1;
  }
  return Math.max(0, Math.min(offset, paragraphLength(p, slideNumber)));
}

/**
 * Widen `[from, to)` in one paragraph so it never cuts a field or a surrogate pair: a deleted
 * field goes whole, and a deleted emoji goes with both of its halves.
 */
function widenRange(
  p: Element,
  from: number,
  to: number,
  slideNumber: number | undefined
): [number, number] {
  let a = from;
  let b = to;
  for (const s of slotsOf(p, slideNumber)) {
    if (s.kind === 'field') {
      if (a > s.start && a < s.end) a = s.start;
      if (b > s.start && b < s.end) b = s.end;
    } else if (s.kind === 'run') {
      if (a > s.start && a < s.end && splitsPair(s.text, a - s.start)) a -= 1;
      if (b > s.start && b < s.end && splitsPair(s.text, b - s.start)) b += 1;
    }
  }
  return [a, b];
}

/** Compare two points. */
export function comparePoints(a: TextPoint, b: TextPoint): number {
  return a.paragraph - b.paragraph || a.offset - b.offset;
}

/** The run properties spec a new run at a point inherits. */
function rPrSpecAt(p: Element, offset: number, slideNumber: number | undefined): NodeSpec {
  const slots = slotsOf(p, slideNumber).filter((s) => s.kind === 'run');
  const before = [...slots].reverse().find((s) => s.start < offset) ?? slots[0];
  const source = before ? child(before.el, 'a:rPr') : child(p, 'a:endParaRPr');
  if (!source) return aSpec('rPr', { lang: 'en-US', dirty: '0' });
  const spec = specOf(source);
  return { ...spec, local: 'rPr' };
}

/** Index in the paragraph's children where inline content at `offset` starts. */
function childIndexForOffset(p: Element, offset: number, slideNumber: number | undefined): number {
  const slots = slotsOf(p, slideNumber);
  let index = 0;
  for (let i = 0; i < p.children.length; i += 1) {
    const c = p.children[i];
    if (c.type !== 'element') continue;
    if (c.local === 'pPr') index = i + 1;
    const slot = slots.find((s) => s.el.id === c.id);
    if (slot && slot.end <= offset) index = i + 1;
  }
  return index;
}

/**
 * Ensure a run boundary at `offset` by splitting the run that spans it. Returns nothing; the
 * paragraph must be re-read from the batch afterwards.
 */
export function splitRunAt(batch: OpBatch, part: string, pId: string, offset: number): void {
  const p = batch.element(part, pId);
  const slot = slotsOf(p, slideNumberOf(batch, part)).find(
    (s) => s.kind === 'run' && offset > s.start && offset < s.end
  );
  if (!slot) return;
  const t = child(slot.el, 'a:t');
  if (!t) return;
  const cut = offset - slot.start;
  batch.push({ op: 'setText', part, node: t.id, text: slot.text.slice(0, cut) });
  const rPr = child(slot.el, 'a:rPr');
  const index = p.children.findIndex((c) => c.id === slot.el.id) + 1;
  batch.insert({
    op: 'insertNode',
    part,
    parent: pId,
    index,
    spec: aSpec('r', {}, [
      ...(rPr ? [specOf(rPr)] : []),
      aSpec('t', {}, [{ text: slot.text.slice(cut) }]),
    ]),
  });
}

/** Insert text at a point. Returns the point after the insertion. */
export function insertText(
  batch: OpBatch,
  part: string,
  txBodyId: string,
  at: TextPoint,
  text: string
): TextPoint {
  const txBody = batch.element(part, txBodyId);
  const p = paragraphsOf(txBody)[at.paragraph];
  if (!p) throw new CommandRefusal('errors.unsupported', 'no paragraph');
  const n = slideNumberOf(batch, part);
  const offset = snapPoint(p, at.offset, n);
  const slots = slotsOf(p, n);
  // Prefer an empty run at the caret (a format armed for this text), then the run that ends
  // at the caret, then the run that contains it.
  const host =
    slots.find((s) => s.kind === 'run' && s.start === offset && s.end === offset) ??
    slots.find((s) => s.kind === 'run' && offset > s.start && offset <= s.end) ??
    slots.find((s) => s.kind === 'run' && offset === s.start && s.start === 0);
  if (host) {
    const t = child(host.el, 'a:t');
    if (t) {
      const local = offset - host.start;
      batch.push({
        op: 'setText',
        part,
        node: t.id,
        text: host.text.slice(0, local) + text + host.text.slice(local),
      });
      return { paragraph: at.paragraph, offset: offset + text.length };
    }
  }
  batch.insert({
    op: 'insertNode',
    part,
    parent: p.id,
    index: childIndexForOffset(p, offset, n),
    spec: aSpec('r', {}, [rPrSpecAt(p, offset, n), aSpec('t', {}, [{ text }])]),
  });
  return { paragraph: at.paragraph, offset: offset + text.length };
}

/** Delete `[from, to)` within one paragraph. */
function deleteInParagraph(
  batch: OpBatch,
  part: string,
  pId: string,
  from: number,
  to: number
): void {
  if (to <= from) return;
  const p = batch.element(part, pId);
  for (const s of slotsOf(p, slideNumberOf(batch, part))) {
    if (s.end <= from || s.start >= to) continue;
    if (s.kind !== 'run' || (s.start >= from && s.end <= to)) {
      batch.push({ op: 'removeNode', part, node: s.el.id });
      continue;
    }
    const t = child(s.el, 'a:t');
    if (!t) continue;
    const a = Math.max(from, s.start) - s.start;
    const b = Math.min(to, s.end) - s.start;
    batch.push({ op: 'setText', part, node: t.id, text: s.text.slice(0, a) + s.text.slice(b) });
  }
}

/**
 * Delete the range between two points (any order). A range end inside a field or a surrogate
 * pair widens to take all of it. Returns the collapsed point.
 */
export function deleteRange(
  batch: OpBatch,
  part: string,
  txBodyId: string,
  a: TextPoint,
  b: TextPoint
): TextPoint {
  const [lo, hi] = comparePoints(a, b) <= 0 ? [a, b] : [b, a];
  const paragraphs = paragraphsOf(batch.element(part, txBodyId));
  const first = paragraphs[lo.paragraph];
  const last = paragraphs[hi.paragraph];
  if (!first || !last) throw new CommandRefusal('errors.unsupported', 'no paragraph');
  const n = slideNumberOf(batch, part);
  const from = { paragraph: lo.paragraph, offset: widenRange(first, lo.offset, lo.offset, n)[0] };
  const to = { paragraph: hi.paragraph, offset: widenRange(last, hi.offset, hi.offset, n)[1] };
  if (first.id === last.id) {
    deleteInParagraph(batch, part, first.id, from.offset, to.offset);
    return from;
  }
  deleteInParagraph(batch, part, first.id, from.offset, paragraphLength(first, n));
  deleteInParagraph(batch, part, last.id, 0, to.offset);
  // Move what remains of the last paragraph into the first, then drop the paragraphs between.
  const lastNow = batch.element(part, last.id);
  for (const s of slotsOf(lastNow, n)) {
    const firstNow = batch.element(part, first.id);
    const endIndex = firstNow.children.findIndex(
      (c) => c.type === 'element' && c.local === 'endParaRPr'
    );
    batch.push({
      op: 'moveNode',
      part,
      node: s.el.id,
      parent: first.id,
      index: endIndex < 0 ? firstNow.children.length : endIndex,
    });
  }
  for (let i = from.paragraph + 1; i <= to.paragraph; i += 1)
    batch.push({ op: 'removeNode', part, node: paragraphs[i].id });
  return from;
}

/** Split a paragraph at a point (Enter). Returns the start of the new paragraph. */
export function splitParagraph(
  batch: OpBatch,
  part: string,
  txBodyId: string,
  at: TextPoint
): TextPoint {
  const txBody = batch.element(part, txBodyId);
  const p = paragraphsOf(txBody)[at.paragraph];
  if (!p) throw new CommandRefusal('errors.unsupported', 'no paragraph');
  const n = slideNumberOf(batch, part);
  const offset = snapPoint(p, at.offset, n);
  splitRunAt(batch, part, p.id, offset);
  const pNow = batch.element(part, p.id);
  const pPr = child(pNow, 'a:pPr');
  const end = child(pNow, 'a:endParaRPr');
  const index = txBody.children.findIndex((c) => c.id === p.id) + 1;
  const newId = batch.insert({
    op: 'insertNode',
    part,
    parent: txBodyId,
    index,
    spec: aSpec('p', {}, [
      ...(pPr ? [specOf(pPr)] : []),
      ...(end ? [specOf(end)] : [{ ...rPrSpecAt(pNow, offset, n), local: 'endParaRPr' }]),
    ]),
  });
  const moving = slotsOf(batch.element(part, p.id), n).filter((s) => s.start >= offset);
  let target = pPr ? 1 : 0;
  for (const s of moving) {
    batch.push({ op: 'moveNode', part, node: s.el.id, parent: newId, index: target });
    target += 1;
  }
  return { paragraph: at.paragraph + 1, offset: 0 };
}

/** Insert a line break (Shift+Enter). */
export function insertBreak(
  batch: OpBatch,
  part: string,
  txBodyId: string,
  at: TextPoint
): TextPoint {
  const p = paragraphsOf(batch.element(part, txBodyId))[at.paragraph];
  if (!p) throw new CommandRefusal('errors.unsupported', 'no paragraph');
  const n = slideNumberOf(batch, part);
  const offset = snapPoint(p, at.offset, n);
  splitRunAt(batch, part, p.id, offset);
  const pNow = batch.element(part, p.id);
  batch.insert({
    op: 'insertNode',
    part,
    parent: p.id,
    index: childIndexForOffset(pNow, offset, n),
    spec: aSpec('br', {}, [rPrSpecAt(pNow, offset, n)]),
  });
  return { paragraph: at.paragraph, offset: offset + 1 };
}

/**
 * A paragraph's text for the caret word rule (`caretWordRange`, shared with DOCX Editor). A
 * field is an object character (U+FFFC) for each of its units, as DOCX Editor reads an object
 * in a run, so a field ends a word.
 */
export function wordTextOf(p: Element, slideNumber?: number): string {
  return slotsOf(p, slideNumber)
    .map((s) => (s.kind === 'field' ? '\uFFFC'.repeat(s.text.length) : s.text))
    .join('');
}

/** No inline stops: a slide paragraph has no inline control edges. */
export const NO_WORD_STOPS: ReadonlySet<number> = new Set();

/**
 * An empty run at the caret that carries a format for the text typed next ("arms" it, as
 * DOCX Editor does at a word edge). `insertText` types into it. An empty run that
 * is already at the caret is reused. The batch is marked `armedCaret`: the editor keeps it as
 * session state and never commits it alone (`caret-arm.ts`).
 */
function armedRun(batch: OpBatch, part: string, pId: string, offset: number): Element {
  batch.armedCaret = true;
  const n = slideNumberOf(batch, part);
  const empty = (): Element | undefined =>
    slotsOf(batch.element(part, pId), n).find(
      (s) => s.kind === 'run' && s.start === offset && s.end === offset
    )?.el;
  const found = empty();
  if (found) return found;
  splitRunAt(batch, part, pId, offset);
  const p = batch.element(part, pId);
  const id = batch.insert({
    op: 'insertNode',
    part,
    parent: pId,
    index: childIndexForOffset(p, offset, n),
    spec: aSpec('r', {}, [rPrSpecAt(p, offset, n), aSpec('t')]),
  });
  return batch.element(part, id);
}

/** The paragraph mark (`a:endParaRPr`, last in `CT_TextParagraph`), added when it is absent. */
function paragraphMark(batch: OpBatch, part: string, p: Element): Element {
  const found = child(p, 'a:endParaRPr');
  if (found) return found;
  const id = batch.insert({
    op: 'insertNode',
    part,
    parent: p.id,
    index: p.children.length,
    spec: aSpec('endParaRPr', { lang: 'en-US', dirty: '0' }),
  });
  return batch.element(part, id);
}

/**
 * The run elements a formatting command touches: the runs inside the range (split at its
 * edges), or every run and paragraph mark of the body when the range is absent (the shape is
 * selected). At a caret inside a word, the word; in an empty paragraph, its paragraph mark; at
 * any other caret, an empty run that formats the text typed next.
 */
export function runsForFormatting(
  batch: OpBatch,
  part: string,
  txBodyId: string,
  range?: { from: TextPoint; to: TextPoint },
  /** False for a pen (highlight): a caret inside a word arms the next typed text too. */
  wordAtCaret = true
): Element[] {
  const n = slideNumberOf(batch, part);
  if (range && comparePoints(range.from, range.to) === 0) {
    const at = range.from;
    const p = paragraphsOf(batch.element(part, txBodyId))[at.paragraph];
    if (!p) return [];
    const word = wordAtCaret ? caretWordRange(wordTextOf(p, n), at.offset, NO_WORD_STOPS) : null;
    if (word)
      return runsForFormatting(batch, part, txBodyId, {
        from: { paragraph: at.paragraph, offset: word.start },
        to: { paragraph: at.paragraph, offset: word.end },
      });
    // An empty paragraph keeps the format on its paragraph mark, which the text typed there
    // takes and which sets the empty line's height (`a:endParaRPr`). An empty run that
    // another tool wrote there (`<a:r><a:rPr/><a:t/></a:r>`) takes it too: typed text goes into
    // that run, and the format reads from it.
    if (paragraphLength(p, n) === 0)
      return [...elements(p).filter((c) => c.local === 'r'), paragraphMark(batch, part, p)];
    return [armedRun(batch, part, p.id, at.offset)];
  }
  const collapsed = !range;
  if (collapsed) {
    const out: Element[] = [];
    for (const p of paragraphsOf(batch.element(part, txBodyId))) {
      for (const s of slotsOf(p, n)) if (s.kind !== 'break' || child(s.el, 'a:rPr')) out.push(s.el);
      const end = child(p, 'a:endParaRPr');
      if (end) out.push(end);
    }
    return out;
  }
  const [from, to] =
    comparePoints(range.from, range.to) <= 0 ? [range.from, range.to] : [range.to, range.from];
  const ids = paragraphsOf(batch.element(part, txBodyId)).map((p) => p.id);
  // Range edges never cut a surrogate pair: a run split there would hold half a character.
  const edges = (p: Element, pi: number): [number, number] =>
    widenRange(
      p,
      pi === from.paragraph ? from.offset : 0,
      pi === to.paragraph ? to.offset : paragraphLength(p, n),
      n
    );
  for (let pi = from.paragraph; pi <= to.paragraph; pi += 1) {
    const p = batch.element(part, ids[pi]);
    const [a, b] = edges(p, pi);
    splitRunAt(batch, part, p.id, b);
    splitRunAt(batch, part, p.id, a);
  }
  const out: Element[] = [];
  for (let pi = from.paragraph; pi <= to.paragraph; pi += 1) {
    const p = batch.element(part, ids[pi]);
    const [a, b] = edges(p, pi);
    // A field inside the range takes the format too.
    for (const s of slotsOf(p, n))
      if ((s.kind === 'run' || s.kind === 'field') && s.start >= a && s.end <= b && s.end > s.start)
        out.push(s.el);
  }
  return out;
}

/** `a:rPr` children in schema order (CT_TextCharacterProperties). */
const RPR_ORDER = [
  'ln',
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
  'effectLst',
  'effectDag',
  'highlight',
  'uLnTx',
  'uLn',
  'uFillTx',
  'uFill',
  'latin',
  'ea',
  'cs',
  'sym',
  'hlinkClick',
  'hlinkMouseOver',
  'rtl',
  'extLst',
];

/** The properties element of a run-like node, created when missing. Returns its id. */
export function ensureRunProps(batch: OpBatch, part: string, runLike: Element): string {
  if (runLike.local === 'endParaRPr') return runLike.id;
  const existing = child(batch.element(part, runLike.id), 'a:rPr');
  if (existing) return existing.id;
  return batch.insert({
    op: 'insertNode',
    part,
    parent: runLike.id,
    index: 0,
    spec: aSpec('rPr', { lang: 'en-US', dirty: '0' }),
  });
}

/** Set or remove an attribute on a run's properties. */
export function setRunAttr(
  batch: OpBatch,
  part: string,
  runLike: Element,
  name: string,
  value: string | null
): void {
  const rPrId = ensureRunProps(batch, part, runLike);
  batch.push({ op: 'setAttribute', part, node: rPrId, name: { ns: '', local: name }, value });
}

/**
 * Replace the children of a run's properties that belong to one group (`fill`, `highlight`,
 * `latin`) with `spec`, inserted at its schema position. `spec` null removes them.
 */
export function setRunChild(
  batch: OpBatch,
  part: string,
  runLike: Element,
  group: readonly string[],
  spec: NodeSpec | null
): void {
  const rPrId = ensureRunProps(batch, part, runLike);
  for (const c of elements(batch.element(part, rPrId)))
    if (c.ns === A && group.includes(c.local)) batch.push({ op: 'removeNode', part, node: c.id });
  if (!spec) return;
  const rank = RPR_ORDER.indexOf(spec.local);
  const rPr = batch.element(part, rPrId);
  let index = 0;
  rPr.children.forEach((c, i) => {
    if (c.type === 'element' && RPR_ORDER.indexOf(c.local) <= rank) index = i + 1;
  });
  batch.insert({ op: 'insertNode', part, parent: rPrId, index, spec });
}

/** `a:pPr` of a paragraph, created when missing. Returns its id. */
export function ensureParagraphProps(batch: OpBatch, part: string, pId: string): string {
  const existing = child(batch.element(part, pId), 'a:pPr');
  if (existing) return existing.id;
  return batch.insert({ op: 'insertNode', part, parent: pId, index: 0, spec: aSpec('pPr') });
}

/** Paragraph ids a paragraph command touches: those in the range, or all when none. */
export function paragraphsForFormatting(
  batch: OpBatch,
  part: string,
  txBodyId: string,
  range?: { from: TextPoint; to: TextPoint }
): string[] {
  const all = paragraphsOf(batch.element(part, txBodyId)).map((p) => p.id);
  if (!range) return all;
  const lo = Math.min(range.from.paragraph, range.to.paragraph);
  const hi = Math.max(range.from.paragraph, range.to.paragraph);
  return all.slice(lo, hi + 1);
}

export { aSpec, attr };
