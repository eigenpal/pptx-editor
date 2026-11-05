// Keys in a table cell: Tab moves to the next cell and selects its text,
// Shift+Tab to the previous cell, and Tab in the last cell adds a row and moves into it. An
// arrow key at the edge of a cell's text moves into the next cell that way. Cells are visited in
// reading order: the table layout lists drawn cells row by row.

import type { EditorCommand, EditorSelection, TableCellAddress } from '../contracts/index.ts';
import { apply, invert, type SlideLayout } from '../layout/index.ts';
import { caretAt, pointAt, textFrameMatrix } from './surface-geometry.ts';
import { addressOf, cellItem, cellLayoutAt, tableItem, textItemOf } from './table-cells.ts';
import { parentGroupOf } from '../automation/commands/placement.ts';
import { comparePoints } from '../automation/commands/text-model.ts';
import { moveCaret } from './text-input.ts';

type TextSelection = Extract<EditorSelection, { kind: 'text' }>;

/** What a key does in a cell: an optional command, then the selection after it. */
export interface TableKeyAction {
  readonly command?: EditorCommand;
  /** The selection to show, read from the layout after the command. */
  readonly select: (layout: SlideLayout | undefined) => EditorSelection | undefined;
}

/** The whole text of a cell, selected (or a caret at its start or end). */
function cellSelection(
  layout: SlideLayout | undefined,
  nodeId: string,
  cell: TableCellAddress,
  mode: 'all' | 'start' | 'end'
): EditorSelection | undefined {
  const frame = tableItem(layout, nodeId);
  const item = frame ? cellItem(frame, cell) : undefined;
  if (!item) return undefined;
  const lens = item.text?.paragraphLengths ?? [0];
  const start = { paragraph: 0, offset: 0 };
  const end = { paragraph: Math.max(0, lens.length - 1), offset: lens[lens.length - 1] ?? 0 };
  return {
    kind: 'text',
    nodeId,
    cell,
    anchor: mode === 'end' ? end : start,
    focus: mode === 'start' ? start : end,
  };
}

/** A caret at the end of a table's first cell: Enter, F2, or a typed key on a selected table. */
export function firstCellCaret(
  layout: SlideLayout | undefined,
  nodeId: string
): EditorSelection | undefined {
  const first = tableItem(layout, nodeId)?.table.cells[0];
  return first ? cellSelection(layout, nodeId, addressOf(first), 'end') : undefined;
}

/** The grid step of an arrow key: [rows, columns]. */
const ARROWS: Readonly<Record<string, readonly [number, number]>> = {
  ArrowLeft: [0, -1],
  ArrowRight: [0, 1],
  ArrowUp: [-1, 0],
  ArrowDown: [1, 0],
};

/**
 * Keys with a cell range selected. Delete and Backspace clear the cells, Escape selects the
 * table, Tab and Shift+Tab select the text of the cell after (or before) the focus cell, an
 * arrow puts the caret in the cell beside the focus cell, and Shift+arrow moves the range's
 * focus. The focus stays in the table at its edges.
 */
function cellsKeyAction(
  layout: SlideLayout | undefined,
  sel: Extract<EditorSelection, { kind: 'cells' }>,
  key: string,
  shift: boolean
): TableKeyAction | null {
  if (key === 'Delete' || key === 'Backspace')
    return { command: { type: 'delete' }, select: () => sel };
  if (key === 'Escape') {
    // Escape selects the table, in the group the user entered, as Escape from text does.
    const group = parentGroupOf(layout?.items ?? [], sel.nodeId);
    return {
      select: () => ({ kind: 'shapes', nodeIds: [sel.nodeId], ...(group ? { group } : {}) }),
    };
  }
  const frame = tableItem(layout, sel.nodeId);
  const here = frame ? cellLayoutAt(frame.table, sel.focus) : undefined;
  if (!frame || !here) return null;
  const id = sel.nodeId;
  if (key === 'Tab') {
    const drawn = frame.table.cells;
    const next = drawn[drawn.indexOf(here) + (shift ? -1 : 1)] ?? here;
    return { select: (l) => cellSelection(l, id, addressOf(next), 'all') };
  }
  const step = ARROWS[key];
  if (!step) return null;
  const c = here.cell;
  const row = step[0] > 0 ? c.row + c.rowSpan : step[0] < 0 ? c.row - 1 : c.row;
  const column = step[1] > 0 ? c.col + c.gridSpan : step[1] < 0 ? c.col - 1 : c.col;
  const target = cellLayoutAt(frame.table, { row, column });
  if (!target) return { select: () => sel };
  if (shift) {
    const to = addressOf(target);
    // Back to one cell: a selection of its text, as a pointer drag back into the cell gives.
    if (to.row === sel.anchor.row && to.column === sel.anchor.column)
      return { select: (l) => cellSelection(l, id, to, 'all') };
    return { select: () => ({ ...sel, focus: to }) };
  }
  return { select: (l) => cellSelection(l, id, addressOf(target), 'start') };
}

/** The key's action in a table, or null to leave the key to text editing and the slide. */
export function tableKeyAction(
  layout: SlideLayout | undefined,
  selection: EditorSelection,
  key: string,
  shift: boolean,
  other: boolean
): TableKeyAction | null {
  if (other) return null;
  if (selection.kind === 'cells') return cellsKeyAction(layout, selection, key, shift);
  if (selection.kind !== 'text' || !selection.cell) return null;
  const sel: TextSelection = selection;
  const frame = tableItem(layout, sel.nodeId);
  const here = frame ? cellLayoutAt(frame.table, selection.cell) : undefined;
  if (!frame || !here) return null;
  const drawn = frame.table.cells;
  const i = drawn.indexOf(here);
  const id = sel.nodeId;
  if (key === 'Tab') {
    const next = drawn[i + (shift ? -1 : 1)];
    if (next) return { select: (l) => cellSelection(l, id, addressOf(next), 'all') };
    if (shift) return { select: () => sel };
    // Tab in the last cell adds a row below and moves into its first cell.
    const row = here.cell.row + here.cell.rowSpan;
    return {
      command: { type: 'insertRow', where: 'below' },
      select: (l) => cellSelection(l, id, { row, column: 0 }, 'start'),
    };
  }
  if (shift || comparePoints(sel.anchor, sel.focus) !== 0) return null;
  const item = textItemOf(layout, sel);
  if (!item) return null;
  if (key === 'ArrowRight' || key === 'ArrowLeft') {
    const moved = moveCaret(item, sel.focus, key);
    if (comparePoints(moved, sel.focus) !== 0) return null;
    const next = drawn[i + (key === 'ArrowRight' ? 1 : -1)];
    if (!next) return null;
    return {
      select: (l) => cellSelection(l, id, addressOf(next), key === 'ArrowRight' ? 'start' : 'end'),
    };
  }
  if (key === 'ArrowDown' || key === 'ArrowUp') {
    const moved = moveCaret(item, sel.focus, key);
    if (comparePoints(moved, sel.focus) !== 0) return null;
    const down = key === 'ArrowDown';
    const row = down ? here.cell.row + here.cell.rowSpan : here.cell.row - 1;
    const target = cellLayoutAt(frame.table, { row, column: here.cell.col });
    if (!target) return null;
    // Keep the caret's x position, on the first (or last) line of the cell it moves into.
    const caret = caretAt(item, sel.focus);
    return {
      select: (l) => {
        const f = tableItem(l, id);
        const to = f ? cellItem(f, addressOf(target)) : undefined;
        const m = to ? textFrameMatrix(to) : undefined;
        const lines = to?.text?.lines ?? [];
        const line = down ? lines[0] : lines[lines.length - 1];
        if (!to || !m || !caret || !line) return cellSelection(l, id, addressOf(target), 'start');
        const local = apply(invert(m), caret.x1, caret.y1);
        const s = apply(m, local.x, line.top + line.height / 2);
        const point = pointAt(to, s.x, s.y) ?? { paragraph: 0, offset: 0 };
        return { kind: 'text', nodeId: id, cell: addressOf(target), anchor: point, focus: point };
      },
    };
  }
  return null;
}

/** Run a table key: its command first, then the selection read from the new layout. */
export function applyTableKey(
  action: TableKeyAction,
  exec: (command: EditorCommand) => { readonly ok: boolean },
  layout: () => SlideLayout | undefined,
  select: (selection: EditorSelection) => void
): void {
  if (action.command && !exec(action.command).ok) return;
  const next = action.select(layout());
  if (next) select(next);
}
