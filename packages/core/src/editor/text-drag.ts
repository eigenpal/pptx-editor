// The unit of a text press and its drag, as DOCX Editor's `surface-pointer.ts` granularity: a
// first press selects by character, a second by word, a third by paragraph, a fourth by
// character again, and a drag after it extends by that unit from the first range. A drag
// after a double-click then takes whole words, and after a triple-click whole paragraphs.

import type { EditorSelection, TextPoint } from '../contracts/index.ts';
import type { LayoutItem } from '../layout/index.ts';
import { paragraphUnits, wordSelectionAt } from '../automation/commands/text-units.ts';

/** What one press selects. @internal */
export type TextDragUnit = 'character' | 'word' | 'paragraph';

/** A text drag: the item, the unit, and the range that the press selected. @internal */
export interface TextDrag {
  readonly id: string;
  readonly unit: TextDragUnit;
  /** The start of the pressed range (the anchor of a character drag). */
  readonly anchor: TextPoint;
  /** The end of the pressed range; the anchor for a character drag. */
  readonly anchorEnd: TextPoint;
}

const before = (a: TextPoint, b: TextPoint): boolean =>
  a.paragraph < b.paragraph || (a.paragraph === b.paragraph && a.offset < b.offset);

const UNITS: readonly TextDragUnit[] = ['character', 'word', 'paragraph'];

/** The unit for a press: character, word, paragraph, then round again, as DOCX Editor. */
export function unitOfPress(count: number): TextDragUnit {
  return UNITS[(Math.max(1, count) - 1) % UNITS.length]!;
}

/**
 * The selection covers every character of its text body: a shape selection, or a text range
 * from the first offset to the last. Undefined layout or item: false.
 * @internal
 */
export function coversWholeBody(sel: EditorSelection, item: LayoutItem | undefined): boolean {
  if (sel.kind === 'shapes') return true;
  if (sel.kind !== 'text' || item?.kind !== 'shape') return false;
  const lens = item.text?.paragraphLengths ?? [];
  const last = lens.length - 1;
  const [a, b] = before(sel.focus, sel.anchor) ? [sel.focus, sel.anchor] : [sel.anchor, sel.focus];
  return (
    last >= 0 &&
    a.paragraph === 0 &&
    a.offset === 0 &&
    b.paragraph === last &&
    b.offset === lens[last]
  );
}

/** The range that `unit` selects at `point` in `item`'s text, as `[from, to]`. @internal */
export function unitRange(
  item: LayoutItem,
  point: TextPoint,
  unit: TextDragUnit
): [TextPoint, TextPoint] {
  if (unit === 'character' || item.kind !== 'shape') return [point, point];
  const { paragraph } = point;
  if (unit === 'paragraph') {
    const length = item.text?.paragraphLengths?.[paragraph] ?? 0;
    return [
      { paragraph, offset: 0 },
      { paragraph, offset: length },
    ];
  }
  const p = item.element.text?.paragraphs[paragraph];
  if (!p) return [point, point];
  const [from, to] = wordSelectionAt(paragraphUnits(p), point.offset);
  return [
    { paragraph, offset: from },
    { paragraph, offset: to },
  ];
}

/**
 * The selection while the pointer is at `point`: the pressed range joined with the unit
 * under the pointer, with the focus on the pointer's side, so Shift+arrows go on from there.
 * @internal
 */
export function dragSelection(
  drag: TextDrag,
  item: LayoutItem,
  point: TextPoint
): { anchor: TextPoint; focus: TextPoint } {
  if (drag.unit === 'character') return { anchor: drag.anchor, focus: point };
  const [from, to] = unitRange(item, point, drag.unit);
  return before(from, drag.anchor)
    ? { anchor: drag.anchorEnd, focus: from }
    : { anchor: drag.anchor, focus: before(to, drag.anchorEnd) ? drag.anchorEnd : to };
}

/**
 * Counts the presses of a multi-click, as DOCX Editor's `surface-pointer.ts` does: within
 * 500 ms and 4 px of the last one. Counted here, as browsers disagree on a press's `detail`.
 * @internal
 */
export class PressCounter {
  #last = { count: 0, at: 0, x: 0, y: 0 };

  /** Count one press and return its number in the multi-click (1 for a single press). */
  count(e: { readonly clientX: number; readonly clientY: number }): number {
    const now = Date.now();
    const c = this.#last;
    const near = Math.abs(e.clientX - c.x) <= 4 && Math.abs(e.clientY - c.y) <= 4;
    const count = near && now - c.at >= 0 && now - c.at <= 500 ? c.count + 1 : 1;
    this.#last = { count, at: now, x: e.clientX, y: e.clientY };
    return count;
  }
}
