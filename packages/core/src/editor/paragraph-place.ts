// The place of an IME composition through other edits (a peer's, or the API's): its
// paragraph's element id, its offset, and the paragraph's text when it was last seen. After an
// edit, the offset maps through a diff of the old and the new paragraph text. The caret that
// the edit left is never used: the composition's text goes in only at its own place
// (`composition.ts`). DOM-free.

import type { EditorSelection, TableCellAddress } from '../contracts/index.ts';
import type { OpBatch } from '../automation/commands/op-batch.ts';
import { paragraphText, paragraphsOf, slideNumberOf } from '../automation/commands/text-model.ts';
import { diffHunks, type TextHunk } from '../store/package/text-diff.ts';
import { elements, type Element } from '../store/index.ts';
import { selectionBody } from './table-cells.ts';

type TextCaret = Extract<EditorSelection, { kind: 'text' }>;

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

/** The caret `sel` at `offset` of the paragraph `at`. */
function caretAt(sel: TextCaret, at: ParagraphAt, offset: number): TextCaret {
  const point = { paragraph: at.paragraph, offset };
  const { cell: _cell, ...rest } = sel;
  return { ...rest, ...(at.cell ? { cell: at.cell } : {}), anchor: point, focus: point };
}

/**
 * The caret `sel` moved to the paragraph with element id `id` in the same shape or table, at
 * the same offset; null when that paragraph is gone. Reads only.
 */
export function caretInParagraph(
  batch: OpBatch,
  part: string,
  sel: TextCaret,
  id: string
): TextCaret | null {
  const root = batch.element(part, sel.nodeId);
  const at = root ? find(root, id) : null;
  return at ? caretAt(sel, at, sel.focus.offset) : null;
}

/**
 * `offset` of the text the `hunks` change, in the changed text. Text inserted at the offset
 * comes after it: the composition was there first. Null when the offset lies inside a
 * changed region: the text just before it is no longer the same text.
 */
export function mapOffset(hunks: readonly TextHunk[], offset: number): number | null {
  let by = 0;
  for (const h of hunks) {
    if (h.aStart < offset && h.aEnd <= offset) by += h.bEnd - h.bStart - (h.aEnd - h.aStart);
    else if (h.aStart >= offset) break;
    else return null;
  }
  return offset + by;
}

/** A composition's place: its caret, its paragraph's element id, and that paragraph's text. */
export interface ParagraphPlace {
  readonly caret: TextCaret;
  readonly id: string;
  readonly text: string;
}

/**
 * The place `place` after edits. When its paragraph is still there, the offset maps through
 * the diff of the paragraph's old and new text (`mapOffset`). When the text before the place
 * left with the rest of the paragraph (a split: an Enter before the place), the place goes
 * with it to the next paragraph. When the paragraph is gone (a merge: a Backspace at its
 * start), the place goes to the one paragraph of the same text body that holds the text
 * before it. Each move checks that the text before the place came with it. Null when no
 * place holds: the composition's text then goes nowhere. Reads only.
 */
export function placeAfterEdits(
  batch: OpBatch,
  part: string,
  place: ParagraphPlace
): ParagraphPlace | null {
  const { caret, text: old } = place;
  const offset = caret.focus.offset;
  const slide = slideNumberOf(batch, part);
  /** The paragraph `at`, and its text. */
  const paragraph = (at: ParagraphAt) => {
    const body = selectionBody(batch, part, caretAt(caret, at, 0));
    const p = body && paragraphsOf(body)[at.paragraph];
    return p ? { id: p.id, text: paragraphText(p, slide) } : null;
  };
  /** The place in paragraph `at`, when the `moved` text came there before it. */
  const moveTo = (at: ParagraphAt, moved: string): ParagraphPlace | null => {
    const p = paragraph(at);
    if (!p) return null;
    const there = mapOffset(diffHunks(old, p.text), offset);
    if (there === null || p.text.slice(there - moved.length, there) !== moved) return null;
    return { caret: caretAt(caret, at, there), id: p.id, text: p.text };
  };
  const root = batch.element(part, caret.nodeId);
  const at = root && find(root, place.id);
  if (at) {
    const now = paragraph(at)?.text ?? '';
    // A split before the place: the end of the paragraph, from before the place, left it.
    const split = diffHunks(old, now).find(
      (h) => h.aStart < offset && h.aEnd === old.length && h.bEnd === now.length
    );
    const moved = split && old.slice(split.aStart, offset);
    const next = moved ? moveTo({ ...at, paragraph: at.paragraph + 1 }, moved) : null;
    return next ?? moveTo(at, '');
  }
  // A merge: the text before the place is in exactly one paragraph of the same text body.
  const before = old.slice(0, offset);
  const body = before ? selectionBody(batch, part, caret) : undefined;
  if (!body) return null;
  const found = paragraphsOf(body)
    .map((_, paragraph) =>
      moveTo({ ...(caret.cell ? { cell: caret.cell } : {}), paragraph }, before)
    )
    .filter((p): p is ParagraphPlace => p !== null);
  return found.length === 1 ? found[0]! : null;
}
