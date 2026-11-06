// Pointer gestures on a table: a click in a cell places the caret there, a
// drag inside one cell selects text, a drag across cells selects a cell range, Shift+click
// extends to a range, and a drag on an inner grid line resizes the columns or rows beside it.
// Moved out of `surface.ts` (the 1000-line cap); the surface calls these before its own hits.

import type { LayoutItem, Matrix, SlideLayout } from '../layout/index.ts';
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
import { parentGroupOf } from '../automation/commands/placement.ts';
import {
  cellAtPoint,
  cellItem,
  cellRangeOf,
  cellRangeRects,
  selectedLayoutRange,
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
  /**
   * How far, in table units, the right or bottom edge may grow: to the slide's edge, so the edge
   * stays where the pointer can reach it again. Infinite for an inner line.
   */
  readonly room: number;
  /** Every step of the drag runs with it: the drag is one undo step. */
  readonly historyGroup: HistoryGroup;
  /** The table is not the selection: the first move selects it, a plain click does not. */
  pending: boolean;
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
    // The format painter never drags a grid line. Any table's lines drag, as the table chrome
    // marks the line under the pointer (`surface-table-chrome.ts`).
    const line = painting ? undefined : gridLineAt(frame, p, this.#tolerance());
    if (line) {
      const t = frame.table;
      const lines = line.axis === 'column' ? t.columnX : t.rowY;
      const sizes = lines.slice(1).map((v, i) => Math.round(v - lines[i]));
      // The grid commands act on the selection's table: a drag on another table's line selects it.
      const sel = this.host.selection();
      const on =
        sel.kind === 'shapes'
          ? sel.nodeIds.length === 1 && sel.nodeIds[0] === id
          : sel.kind !== 'none' && sel.nodeId === id;
      const outer = line.index === lines.length - 1;
      const end =
        line.axis === 'column'
          ? apply(frame.matrix, t.width, 0).x
          : apply(frame.matrix, 0, t.height).y;
      const slideEnd = line.axis === 'column' ? layout.slide.width : layout.slide.height;
      // Slide units for each table unit along the axis: a table in a scaled group is not 1.
      const origin = apply(frame.matrix, 0, 0);
      const length = line.axis === 'column' ? t.width : t.height;
      const scale =
        length > 0 ? (end - (line.axis === 'column' ? origin.x : origin.y)) / length : 1;
      this.#grid = {
        pending: !on,
        id,
        ...line,
        start: apply(invert(frame.matrix), p.x, p.y),
        sizes,
        room: outer && scale > 0 ? Math.max(0, slideEnd - end) / scale : Infinity,
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
      if (g.pending) {
        g.pending = false;
        const group = parentGroupOf(layout.items, g.id);
        this.host.select({ kind: 'shapes', nodeIds: [g.id], ...(group ? { group } : {}) });
      }
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
      } else sizes[i] = Math.max(Math.min(MIN_LINE, sizes[i]), sizes[i] + Math.min(delta, g.room));
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

  /** The gesture ended. Returns true when it was a grid line drag. */
  up(): boolean {
    const grid = this.#grid;
    grid?.historyGroup.end();
    this.#text = null;
    this.#grid = null;
    return !!grid;
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

  /**
   * The selected table when `p` is on one of its grid lines: the right and bottom edges resize
   * the last column or row. With `outside`, only just outside its box, where a shape behind the
   * table must not take the press; inside, shapes in front of the table keep theirs.
   */
  edgeAt(
    layout: SlideLayout | undefined,
    p: { x: number; y: number },
    outside = false
  ): string | undefined {
    const sel = this.host.selection();
    const id =
      sel.kind === 'shapes'
        ? sel.nodeIds.length === 1
          ? sel.nodeIds[0]
          : undefined
        : sel.kind === 'none'
          ? undefined
          : sel.nodeId;
    const frame = id ? tableItem(layout, id) : undefined;
    if (!frame || !this.onLine(frame, p)) return undefined;
    if (!outside) return id;
    return inGrid(frame, p) ? undefined : id;
  }

  /**
   * Whether `p` is inside the drawn grid of table item `item`: its cells, where a press places a
   * caret and the border band of the selected table does not reach. False for other items.
   */
  inGrid(item: LayoutItem, p: { x: number; y: number }): boolean {
    return (
      item.kind === 'frame' && !!item.table && inGrid({ matrix: item.matrix, table: item.table }, p)
    );
  }

  /**
   * A press at `p` on table item `item` drags a grid line: true wherever the resize cursor
   * shows. The border band of a selected table yields to it, so its outer lines resize.
   */
  onLine(item: LayoutItem, p: { x: number; y: number }): boolean {
    if (item.kind !== 'frame' || !item.table || this.host.readOnly()) return false;
    return !!gridLineAt(item as Parameters<typeof gridLineAt>[0], p, this.#tolerance());
  }

  /** The pointer over table `id`: resize on a grid line, else text in a cell. */
  cursor(layout: SlideLayout, id: string, p: { x: number; y: number }): string | undefined {
    const frame = tableItem(layout, id);
    if (!frame || this.host.readOnly()) return undefined;
    const line = gridLineAt(frame, p, this.#tolerance());
    if (line) return line.axis === 'column' ? 'col-resize' : 'row-resize';
    return cellAtPoint(frame, p) ? 'text' : undefined;
  }
}

/**
 * A table frame's resize: its right and bottom edges stop at the slide's edge, as its grid
 * edges do, so a handle never leaves the place where the pointer can grab it again. `handle` is
 * the dragged handle's place on the box; `dx` and `dy` are in slide units. With `aspect` (Shift
 * or a locked aspect ratio on a corner), the scale stops where the first edge meets the slide's.
 */
export function clampTableResize(
  item: LayoutItem,
  handle: readonly [number, number],
  dx: number,
  dy: number,
  slide: { readonly width: number; readonly height: number },
  aspect = false
): { dx: number; dy: number } {
  if (item.kind !== 'frame' || !item.table) return { dx, dy };
  const { cx, cy } = item.element.xfrm;
  // The drawn grid's far corner: rows that grow to fit their text can pass the frame's.
  const end = apply(item.matrix, item.table.width, item.table.height);
  const roomX = handle[0] === 1 ? Math.max(0, slide.width - end.x) : Infinity;
  const roomY = handle[1] === 1 ? Math.max(0, slide.height - end.y) : Infinity;
  const corner = handle[0] !== 0.5 && handle[1] !== 0.5;
  if (aspect && corner && cx > 0 && cy > 0) {
    // The largest scale that keeps both dragged edges on the slide.
    const grow = Math.min(roomX / cx, roomY / cy);
    const limit = (d: number, size: number, toward: number): number =>
      toward === 1 ? Math.min(d, grow * size) : toward === 0 ? Math.max(d, -grow * size) : d;
    return { dx: limit(dx, cx, handle[0]), dy: limit(dy, cy, handle[1]) };
  }
  return { dx: Math.min(dx, roomX), dy: Math.min(dy, roomY) };
}

/** Whether `p` is inside a table frame's drawn grid, which rows that grow can make taller. */
function inGrid(
  frame: { readonly matrix: Matrix; readonly table: { width: number; height: number } },
  p: { x: number; y: number }
): boolean {
  const local = apply(invert(frame.matrix), p.x, p.y);
  const { width, height } = frame.table;
  return local.x >= 0 && local.x <= width && local.y >= 0 && local.y <= height;
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
  // The cells that commands change (`selectedLayoutRange`): a whole row or column band shows the
  // cells of its own grid lines; any other range, the rectangle its merged cells widen it to.
  const { range, band } = selectedLayoutRange(frame.table, cellRangeOf(sel));
  return cellRangeRects(doc, frame, range, 'pptx-surface__cell-selection', band);
}
