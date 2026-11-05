// Pointer gestures on a table: a click in a cell places the caret there, a
// drag inside one cell selects text, a drag across cells selects a cell range, Shift+click
// extends to a range, and a drag on an inner grid line resizes the columns or rows beside it.
// Moved out of `surface.ts` (the 1000-line cap); the surface calls these before its own hits.

import type { SlideLayout } from '../layout/index.ts';
import type {
  EditorCommand,
  EditorExecOptions,
  EditorSelection,
  HistoryGroup,
  TableCellAddress,
  TextPoint,
} from '../contracts/index.ts';
import { apply, invert } from '../layout/index.ts';
import { pointAt } from './surface-geometry.ts';
import {
  cellAtPoint,
  cellItem,
  cellRangeOf,
  cellRangeRects,
  expandLayoutRange,
  gridLineAt,
  tableItem,
  type ShapeItem,
} from './table-cells.ts';

const EMU_PER_PX = 9_525;
/** The narrowest column or row a drag leaves, EMU (0.1"). */
const MIN_LINE = 91_440;

/** What table gestures need from the surface. */
export interface TableGestureHost {
  readonly selection: () => EditorSelection;
  readonly readOnly: () => boolean;
  readonly zoom: () => number;
  select(selection: EditorSelection): void;
  run(command: EditorCommand, options?: EditorExecOptions): void;
  beginHistoryGroup(): HistoryGroup;
}

type TextDrag = {
  readonly id: string;
  readonly cell: TableCellAddress;
  readonly anchor: TextPoint;
};
type GridDrag = {
  readonly id: string;
  readonly axis: 'column' | 'row';
  readonly index: number;
  readonly start: { x: number; y: number };
  readonly sizes: readonly number[];
  /** Every step of the drag runs with it: the drag is one undo step. */
  readonly historyGroup: HistoryGroup;
};

/** The table gestures of one surface. */
export class TableGestures {
  #text: TextDrag | null = null;
  #grid: GridDrag | null = null;

  constructor(readonly host: TableGestureHost) {}

  /** True while a table gesture runs. */
  get active(): boolean {
    return !!this.#text || !!this.#grid;
  }

  #tolerance(): number {
    return (4 * EMU_PER_PX) / this.host.zoom();
  }

  /** The table's selection state: the frame id when the selection is in or on it. */
  #selectedTable(): string | undefined {
    const sel = this.host.selection();
    if (sel.kind === 'text' && sel.cell) return sel.nodeId;
    if (sel.kind === 'cells') return sel.nodeId;
    if (sel.kind === 'shapes' && sel.nodeIds.length === 1) return sel.nodeIds[0];
    return undefined;
  }

  /**
   * A press on element `id`. Returns `text` when a caret or cell range was set (the surface
   * focuses its input), `grid` for a grid line drag, or null when the press is not for a table.
   */
  down(
    layout: SlideLayout,
    id: string,
    p: { x: number; y: number },
    shift: boolean,
    painting = false
  ): 'text' | 'grid' | null {
    const frame = tableItem(layout, id);
    if (!frame || this.host.readOnly()) return null;
    // The format painter never drags a grid line.
    const line =
      !painting && this.#selectedTable() === id
        ? gridLineAt(frame, p, this.#tolerance())
        : undefined;
    if (line) {
      const t = frame.table;
      const lines = line.axis === 'column' ? t.columnX : t.rowY;
      const sizes = lines.slice(1).map((v, i) => Math.round(v - lines[i]));
      this.#grid = {
        id,
        ...line,
        start: apply(invert(frame.matrix), p.x, p.y),
        sizes,
        historyGroup: this.host.beginHistoryGroup(),
      };
      return 'grid';
    }
    const cell = cellAtPoint(frame, p);
    const item = cell ? cellItem(frame, cell) : undefined;
    if (!cell || !item) return null;
    const point = pointAt(item, p.x, p.y) ?? { paragraph: 0, offset: 0 };
    const sel = this.host.selection();
    if (shift && sel.kind === 'text' && sel.nodeId === id && sel.cell) {
      const same = sel.cell.row === cell.row && sel.cell.column === cell.column;
      this.host.select(
        same
          ? { ...sel, focus: point }
          : { kind: 'cells', nodeId: id, anchor: sel.cell, focus: cell }
      );
      this.#text = { id, cell: sel.cell, anchor: sel.anchor };
      return 'text';
    }
    if (shift && sel.kind === 'cells' && sel.nodeId === id) {
      this.host.select({ ...sel, focus: cell });
      return 'text';
    }
    this.#text = { id, cell, anchor: point };
    this.host.select({ kind: 'text', nodeId: id, cell, anchor: point, focus: point });
    return 'text';
  }

  /** A pointer move while a table gesture runs. Returns true when it was one. */
  move(layout: SlideLayout, p: { x: number; y: number }): boolean {
    const g = this.#grid;
    if (g) {
      const frame = tableItem(layout, g.id);
      if (!frame) return true;
      const local = apply(invert(frame.matrix), p.x, p.y);
      const delta = g.axis === 'column' ? local.x - g.start.x : local.y - g.start.y;
      const sizes = [...g.sizes];
      const i = g.index - 1;
      if (g.axis === 'column' && i + 1 < sizes.length) {
        // An inner column line moves between its two columns; the table keeps its width. A
        // column narrower than the minimum keeps its width; it never jumps wider.
        const lo = Math.min(MIN_LINE, sizes[i]) - sizes[i];
        const hi = sizes[i + 1] - Math.min(MIN_LINE, sizes[i + 1]);
        const d = Math.max(lo, Math.min(hi, delta));
        sizes[i] += d;
        sizes[i + 1] -= d;
      } else sizes[i] = Math.max(Math.min(MIN_LINE, sizes[i]), sizes[i] + delta);
      // The sizes are in visual order; the grid of a right-to-left table runs the other way.
      const columns = frame.table.rtl ? [...sizes].reverse() : sizes;
      const command: EditorCommand =
        g.axis === 'column'
          ? { type: 'setTableGrid', columns: columns.map(Math.round) }
          : { type: 'setTableGrid', rows: sizes.map(Math.round) };
      this.host.run(command, { historyGroup: g.historyGroup });
      return true;
    }
    const t = this.#text;
    if (!t) return false;
    const frame = tableItem(layout, t.id);
    const cell = frame ? cellAtPoint(frame, p) : undefined;
    if (!frame || !cell) return true;
    if (cell.row === t.cell.row && cell.column === t.cell.column) {
      const item = cellItem(frame, cell);
      const point = item ? pointAt(item, p.x, p.y) : undefined;
      if (point)
        this.host.select({ kind: 'text', nodeId: t.id, cell, anchor: t.anchor, focus: point });
    } else this.host.select({ kind: 'cells', nodeId: t.id, anchor: t.cell, focus: cell });
    return true;
  }

  /** The gesture ended. */
  up(): void {
    this.#grid?.historyGroup.end();
    this.#text = null;
    this.#grid = null;
  }

  /** The text item of the cell under a point of table `id`, for a double-click. */
  cellItemAt(
    layout: SlideLayout,
    id: string,
    p: { x: number; y: number }
  ): { item: ShapeItem; cell: TableCellAddress } | undefined {
    const frame = tableItem(layout, id);
    const cell = frame ? cellAtPoint(frame, p) : undefined;
    const item = frame && cell ? cellItem(frame, cell) : undefined;
    return item && cell ? { item, cell } : undefined;
  }

  /**
   * A word or paragraph press at `p` stays in `cell`: the point is in that cell's text, and
   * not on a grid line, where a press drags the line.
   */
  textPressIn(
    layout: SlideLayout,
    id: string,
    cell: TableCellAddress,
    p: { x: number; y: number }
  ): boolean {
    const under = this.cellItemAt(layout, id, p)?.cell;
    const cursor = this.cursor(layout, id, p);
    const same = under?.row === cell.row && under?.column === cell.column;
    return same && (cursor === undefined || cursor === 'text');
  }

  /** The pointer over table `id`: resize on a grid line of the selected table, else text. */
  cursor(layout: SlideLayout, id: string, p: { x: number; y: number }): string | undefined {
    const frame = tableItem(layout, id);
    if (!frame || this.host.readOnly()) return undefined;
    const line = this.#selectedTable() === id ? gridLineAt(frame, p, this.#tolerance()) : undefined;
    if (line) return line.axis === 'column' ? 'col-resize' : 'row-resize';
    return cellAtPoint(frame, p) ? 'text' : undefined;
  }
}

/** The highlight of a selected cell range, editing chrome over the slide. */
export function cellsOverlay(
  doc: Document,
  layout: SlideLayout,
  sel: EditorSelection
): globalThis.Element[] {
  if (sel.kind !== 'cells') return [];
  const frame = tableItem(layout, sel.nodeId);
  if (!frame) return [];
  const range = expandLayoutRange(frame.table, cellRangeOf(sel));
  return cellRangeRects(doc, frame, range, 'pptx-surface__cell-selection');
}
