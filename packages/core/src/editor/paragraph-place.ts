// The place of an IME composition: its paragraph's element id, its offset, and the
// paragraph's exact text before the offset. An edit after the place, or in another paragraph
// or shape, keeps the place: a peer's Enter above the paragraph or a row inserted above its
// cell moves the paragraph's index, not its id. Any other edit (before or over the place, a
// split before it, a merge, a delete) ends the composition, and its text is refused. The
// place is never guessed (`composition.ts`). DOM-free.

import type { EditorSelection, TableCellAddress } from '../contracts/index.ts';
import type { OpBatch } from '../automation/commands/op-batch.ts';
import { paragraphText, paragraphsOf, slideNumberOf } from '../automation/commands/text-model.ts';
import { elements, type Element } from '../store/index.ts';
import { selectionBody } from './table-cells.ts';

type TextCaret = Extract<EditorSelection, { kind: 'text' }>;

/** A composition's place: its caret, its paragraph's element id, and the text before it. */
export interface TextPlace {
  readonly at: EditorSelection;
  readonly paragraph: string;
  readonly before: string;
}

/** Where a paragraph is in its shape or table: its cell, and its index in the text body. */
interface ParagraphAt {
  readonly cell?: TableCellAddress;
  readonly paragraph: number;
}

/** The paragraph with element id `id` under `el`, depth first; tables give its cell. */
function find(el: Element, id: string, cell?: TableCellAddress): ParagraphAt | null {
  const kids = elements(el);
  const index = kids.filter((k) => k.kind === 'paragraph').findIndex((p) => p.id === id);
  if (index >= 0) return { ...(cell ? { cell } : {}), paragraph: index };
  let row = 0;
  for (const kid of kids) {
    if (kid.kind === 'paragraph') continue;
    if (kid.local !== 'tr') {
      const found = find(kid, id, cell);
      if (found) return found;
      continue;
    }
    let column = 0;
    for (const tc of elements(kid)) {
      if (tc.local !== 'tc') continue;
      const found = find(tc, id, { row, column });
      if (found) return found;
      column += 1;
    }
    row += 1;
  }
  return null;
}

/**
 * The caret of `place` now: in the paragraph with its element id, wherever an edit moved that
 * paragraph, at its offset, when the paragraph's text before the offset is exactly
 * `place.before`. Null otherwise. Reads only.
 */
export function placeCaret(batch: OpBatch, part: string, place: TextPlace): TextCaret | null {
  const { at: sel, paragraph: id, before } = place;
  if (sel.kind !== 'text' || sel.diagramNode !== undefined) return null;
  const root = batch.element(part, sel.nodeId);
  const at = root ? find(root, id) : null;
  if (!at) return null;
  const point = { paragraph: at.paragraph, offset: before.length };
  const { cell: _cell, ...rest } = sel;
  const caret: TextCaret = {
    ...rest,
    ...(at.cell ? { cell: at.cell } : {}),
    anchor: point,
    focus: point,
  };
  const body = selectionBody(batch, part, caret);
  const p = body && paragraphsOf(body)[at.paragraph];
  if (p?.id !== id) return null;
  return paragraphText(p, slideNumberOf(batch, part)).startsWith(before) ? caret : null;
}

/**
 * The place of a composition over a selected range: the range, the element ids of its anchor
 * and focus paragraphs, and its exact text. Another selection (a shape) has no ids or text,
 * and holds only while the store revision is the same.
 */
export interface SelectionPlace {
  readonly at: EditorSelection;
  readonly ids: readonly [string, string] | null;
  readonly text: string;
  readonly revision: number;
}

type Point = { readonly paragraph: number; readonly offset: number };
const samePoint = (p: Point, q: Point) => p.paragraph === q.paragraph && p.offset === q.offset;

/** The place of the selection `sel` now. Reads only. */
export function selectionPlace(
  batch: OpBatch,
  part: string,
  sel: EditorSelection,
  revision: number
): SelectionPlace {
  const none = { at: sel, ids: null, text: '', revision };
  if (sel.kind !== 'text' || sel.diagramNode !== undefined) return none;
  const body = selectionBody(batch, part, sel);
  const paragraphs = body ? paragraphsOf(body) : [];
  const a = paragraphs[sel.anchor.paragraph];
  const f = paragraphs[sel.focus.paragraph];
  if (!a || !f) return none;
  const forward =
    sel.anchor.paragraph < sel.focus.paragraph ||
    (sel.anchor.paragraph === sel.focus.paragraph && sel.anchor.offset <= sel.focus.offset);
  const [start, end] = forward ? [sel.anchor, sel.focus] : [sel.focus, sel.anchor];
  const slide = slideNumberOf(batch, part);
  const texts = paragraphs
    .slice(start.paragraph, end.paragraph + 1)
    .map((p) => paragraphText(p, slide));
  const last = texts.length - 1;
  texts[last] = texts[last]!.slice(0, end.offset);
  const text = texts.join('\n').slice(start.offset);
  return { at: sel, ids: [a.id, f.id], text, revision };
}

/** Whether two selections are the same: the same body, cell, anchor, and focus. */
function sameSelection(a: EditorSelection, b: EditorSelection): boolean {
  if (a.kind !== 'text' || b.kind !== 'text') return JSON.stringify(a) === JSON.stringify(b);
  return (
    a.nodeId === b.nodeId &&
    a.diagramNode === b.diagramNode &&
    a.cell?.row === b.cell?.row &&
    a.cell?.column === b.cell?.column &&
    samePoint(a.anchor, b.anchor) &&
    samePoint(a.focus, b.focus)
  );
}

/**
 * Whether the selection place `now` is still `place`: the same range, in the same paragraphs
 * by element id, with the same text. A place with no text range needs the same revision too.
 */
export function holdsSelection(place: SelectionPlace, now: SelectionPlace): boolean {
  if (!sameSelection(place.at, now.at)) return false;
  if (!place.ids || !now.ids) return !place.ids && !now.ids && place.revision === now.revision;
  return place.ids[0] === now.ids[0] && place.ids[1] === now.ids[1] && place.text === now.text;
}
