// Table cells on the laid-out slide: the editing surface sees a cell as a text item, so the
// caret, the text selection, keyboard input, and text formatting work in a cell exactly as in a
// text box. A cell's item is built from the table layout (`layout/table-layout.ts`): its box is
// the cell's box in the frame, and its text is the cell's laid-out text in frame space.

import {
  cellAsSaved,
  child,
  children,
  tableAsDrawn,
  tableCellCounts,
  TABLE_MAX_COLUMNS,
  type Element,
} from '../store/index.ts';
import {
  apply,
  evaluateGeometry,
  invert,
  multiply,
  type CellLayout,
  type LayoutItem,
  type SlideLayout,
  type TableLayout,
} from '../layout/index.ts';
import type { EditorSelection, TableCellAddress } from '../contracts/index.ts';
import { SVG_NS } from '../output/index.ts';
import type { OpBatch } from '../automation/commands/op-batch.ts';
import { findItem } from '../automation/commands/placement.ts';
import {
  emptyBodySpec,
  joinTextBodies,
  newCellSpec,
  tblOf,
  type CellRange,
} from '../automation/commands/table-model.ts';

export type ShapeItem = Extract<LayoutItem, { kind: 'shape' }>;
export type FrameItem = Extract<LayoutItem, { kind: 'frame' }>;
type TableFrame = FrameItem & { readonly table: TableLayout };

const RECT = { kind: 'preset', preset: 'rect', adjust: new Map<string, number>() } as const;

/** The laid-out table frame with node id `id`, or undefined. */
export function tableItem(layout: SlideLayout | undefined, id: string): TableFrame | undefined {
  const item = layout ? findItem(layout.items, id) : undefined;
  return item?.kind === 'frame' && item.table ? (item as TableFrame) : undefined;
}

/** The drawn cell that covers a grid position. */
export function cellLayoutAt(table: TableLayout, at: TableCellAddress): CellLayout | undefined {
  return table.cells.find(
    (c) =>
      at.row >= c.cell.row &&
      at.row < c.cell.row + c.cell.rowSpan &&
      at.column >= c.cell.col &&
      at.column < c.cell.col + c.cell.gridSpan
  );
}

/** The address of a drawn cell: its top-left grid position. */
export const addressOf = (c: CellLayout): TableCellAddress => ({
  row: c.cell.row,
  column: c.cell.col,
});

const items = new WeakMap<CellLayout, ShapeItem>();

/** A table cell as a text item: the cell's box, fill, and text, placed by the frame. */
export function cellItem(frame: TableFrame, at: TableCellAddress): ShapeItem | undefined {
  const cl = cellLayoutAt(frame.table, at);
  if (!cl) return undefined;
  const hit = items.get(cl);
  if (hit) return hit;
  const e = frame.element;
  const item: ShapeItem = {
    kind: 'shape',
    parentMatrix: frame.parentMatrix,
    toSource: frame.toSource,
    matrix: multiply(frame.matrix, [1, 0, 0, 1, cl.x, cl.y]),
    element: {
      kind: 'shape',
      nodeId: e.nodeId,
      partName: e.partName,
      layer: e.layer,
      name: e.name,
      hidden: false,
      xfrm: { x: 0, y: 0, cx: cl.width, cy: cl.height, rot: 0, flipH: false, flipV: false },
      geometry: RECT,
      fill: cl.cell.fill,
      text: cl.cell.text,
    },
    geometry: evaluateGeometry(RECT, cl.width, cl.height),
    text: cl.text,
    // The cell's text is laid out in the frame's space.
    textMatrix: frame.matrix,
  };
  items.set(cl, item);
  return item;
}

/**
 * The text item of a text selection: the selected shape, or the selected table cell. Every
 * caret, selection, and text input path reads text through this.
 */
export function textItemOf(
  layout: SlideLayout | undefined,
  sel: EditorSelection
): ShapeItem | undefined {
  if (sel.kind !== 'text' || !layout) return undefined;
  if (sel.cell) {
    const frame = tableItem(layout, sel.nodeId);
    return frame ? cellItem(frame, sel.cell) : undefined;
  }
  const item = findItem(layout.items, sel.nodeId);
  return item?.kind === 'shape' ? item : undefined;
}

/**
 * A text selection in a table, its cell address moved to the top-left cell of the merged block
 * that covers it: the block's text lives there, so a caret on a covered grid position (a host's
 * address, or a peer's merge around the caret) edits the text the slide shows.
 */
export function snapCellAddress(
  layout: SlideLayout | undefined,
  sel: EditorSelection
): EditorSelection {
  if (sel.kind !== 'text' || !sel.cell) return sel;
  const frame = tableItem(layout, sel.nodeId);
  const cl = frame ? cellLayoutAt(frame.table, sel.cell) : undefined;
  if (!cl || (cl.cell.row === sel.cell.row && cl.cell.col === sel.cell.column)) return sel;
  return { ...sel, cell: addressOf(cl) };
}

/** The text items formatting reads for a selection: shapes, cells, or every cell of a table. */
export function formatItemOf(
  layout: SlideLayout | undefined,
  sel: EditorSelection
): ShapeItem | undefined {
  if (!layout) return undefined;
  if (sel.kind === 'text') return textItemOf(layout, sel);
  const id = sel.kind === 'shapes' ? sel.nodeIds[0] : sel.kind === 'cells' ? sel.nodeId : null;
  if (!id) return undefined;
  const frame = tableItem(layout, id);
  if (frame) {
    const first = sel.kind === 'cells' ? cellRangeOf(sel) : null;
    const cl = first
      ? cellLayoutAt(frame.table, { row: first.r0, column: first.c0 })
      : frame.table.cells[0];
    return cl ? cellItem(frame, addressOf(cl)) : undefined;
  }
  const item = findItem(layout.items, id);
  return item?.kind === 'shape' ? item : undefined;
}

/** The rectangle of a cell selection, first to last. */
export function cellRangeOf(sel: Extract<EditorSelection, { kind: 'cells' }>): CellRange {
  return {
    r0: Math.min(sel.anchor.row, sel.focus.row),
    r1: Math.max(sel.anchor.row, sel.focus.row),
    c0: Math.min(sel.anchor.column, sel.focus.column),
    c1: Math.max(sel.anchor.column, sel.focus.column),
  };
}

/** A range grown over the drawn cells it cuts, so it never splits a merged cell. */
export function expandLayoutRange(table: TableLayout, range: CellRange): CellRange {
  let { r0, r1, c0, c1 } = range;
  for (let changed = true; changed;) {
    changed = false;
    for (const c of table.cells) {
      const cr1 = c.cell.row + c.cell.rowSpan - 1;
      const cc1 = c.cell.col + c.cell.gridSpan - 1;
      if (cr1 < r0 || c.cell.row > r1 || cc1 < c0 || c.cell.col > c1) continue;
      if (c.cell.row < r0 || cr1 > r1 || c.cell.col < c0 || cc1 > c1) {
        r0 = Math.min(r0, c.cell.row);
        r1 = Math.max(r1, cr1);
        c0 = Math.min(c0, c.cell.col);
        c1 = Math.max(c1, cc1);
        changed = true;
      }
    }
  }
  return { r0, r1, c0, c1 };
}

/**
 * The cells a cell range selects on the canvas, as `selectedRange` reads them for commands: a
 * whole row or column band is its own grid lines (`band`); any other range grows over its merged
 * cells.
 */
export function selectedLayoutRange(
  table: TableLayout,
  range: CellRange
): { readonly range: CellRange; readonly band: boolean } {
  const band =
    (range.r0 === 0 && range.r1 === table.rowY.length - 2) ||
    (range.c0 === 0 && range.c1 === table.columnX.length - 2);
  return { range: band ? range : expandLayoutRange(table, range), band };
}

/** The cell under a slide point, or undefined outside the table. */
export function cellAtPoint(
  frame: TableFrame,
  p: { x: number; y: number }
): TableCellAddress | undefined {
  const local = apply(invert(frame.matrix), p.x, p.y);
  const hit = frame.table.cells.find(
    (c) => local.x >= c.x && local.x <= c.x + c.width && local.y >= c.y && local.y <= c.y + c.height
  );
  return hit ? addressOf(hit) : undefined;
}

/** The inner grid line within `tolerance` of a slide point: a column or row boundary index. */
export function gridLineAt(
  frame: TableFrame,
  p: { x: number; y: number },
  tolerance: number
): { readonly axis: 'column' | 'row'; readonly index: number } | undefined {
  const local = apply(invert(frame.matrix), p.x, p.y);
  const t = frame.table;
  if (local.x < -tolerance || local.x > t.width + tolerance) return undefined;
  if (local.y < -tolerance || local.y > t.height + tolerance) return undefined;
  // A line hits only where it is drawn: not inside a merged cell that spans across it.
  const row = lineIndexAt(t.rowY, local.y);
  const column = lineIndexAt(t.columnX, local.x);
  for (let i = 1; i < t.columnX.length; i += 1)
    if (Math.abs(local.x - t.columnX[i]) <= tolerance && !mergedAcross(t, 'column', i, row))
      return { axis: 'column', index: i };
  for (let i = 1; i < t.rowY.length; i += 1)
    if (Math.abs(local.y - t.rowY[i]) <= tolerance && !mergedAcross(t, 'row', i, column))
      return { axis: 'row', index: i };
  return undefined;
}

/** The grid row (or column) that a local coordinate falls in, clamped to the grid. */
function lineIndexAt(lines: readonly number[], v: number): number {
  let i = 0;
  while (i < lines.length - 2 && v >= lines[i + 1]) i += 1;
  return i;
}

/**
 * Whether a merged cell spans across grid line `index` of `axis` at row (or column) `at`: a
 * column line inside a cell that spans two columns there, or a row line inside a cell that spans
 * two rows.
 */
function mergedAcross(t: TableLayout, axis: 'column' | 'row', index: number, at: number): boolean {
  return t.cells.some((c) => {
    const { row, col, rowSpan, gridSpan } = c.cell;
    return axis === 'column'
      ? col < index && col + gridSpan > index && row <= at && row + rowSpan > at
      : row < index && row + rowSpan > index && col <= at && col + gridSpan > at;
  });
}

/** The highlight of a cell range: one rectangle per drawn cell, in slide space. */
export function cellRangeRects(
  doc: Document,
  frame: TableFrame,
  range: CellRange,
  className: string,
  byAnchor = false
): globalThis.Element[] {
  const g = doc.createElementNS(SVG_NS, 'g');
  g.setAttribute('transform', `matrix(${frame.matrix.join(' ')})`);
  // The range's own grid lines: the edge, and the bounds that clip a band's tint.
  const { columnX, rowY, rtl } = frame.table;
  const n = columnX.length - 1;
  const [x0, x1] = rtl
    ? [columnX[n - range.c1 - 1], columnX[n - range.c0]]
    : [columnX[range.c0], columnX[range.c1 + 1]];
  const [y0, y1] = [rowY[range.r0], rowY[range.r1 + 1]];
  const box = [x0, x1, y0, y1].every(Number.isFinite) && x1 > x0 && y1 > y0;
  for (const c of frame.table.cells) {
    const cr1 = c.cell.row + c.cell.rowSpan - 1;
    const cc1 = c.cell.col + c.cell.gridSpan - 1;
    // `byAnchor`: every cell whose top-left grid position is in the range, merged or not.
    const out = byAnchor
      ? c.cell.row < range.r0 ||
        c.cell.row > range.r1 ||
        c.cell.col < range.c0 ||
        c.cell.col > range.c1
      : c.cell.row < range.r0 || cr1 > range.r1 || c.cell.col < range.c0 || cc1 > range.c1;
    if (out) continue;
    // A merged cell that runs past a band is tinted inside the band only, within its edge.
    const left = byAnchor && box ? Math.max(c.x, x0) : c.x;
    const top = byAnchor && box ? Math.max(c.y, y0) : c.y;
    const right = byAnchor && box ? Math.min(c.x + c.width, x1) : c.x + c.width;
    const bottom = byAnchor && box ? Math.min(c.y + c.height, y1) : c.y + c.height;
    if (right <= left || bottom <= top) continue;
    const r = doc.createElementNS(SVG_NS, 'rect');
    r.setAttribute('class', className);
    r.setAttribute('x', String(left));
    r.setAttribute('y', String(top));
    r.setAttribute('width', String(right - left));
    r.setAttribute('height', String(bottom - top));
    g.append(r);
  }
  // One edge around the range's own grid lines, so the cells' borders stay visible inside it
  // and a merged cell that crosses a selected band stays outside the edge.
  if (box) {
    const edge = doc.createElementNS(SVG_NS, 'rect');
    edge.setAttribute('class', `${className}-edge`);
    edge.setAttribute('x', String(x0));
    edge.setAttribute('y', String(y0));
    edge.setAttribute('width', String(x1 - x0));
    edge.setAttribute('height', String(y1 - y0));
    g.append(edge);
  }
  return [g];
}

/**
 * The `a:txBody` of the cell at a grid position in a frame element, as the cascade finds it:
 * the row's `a:tc` at the column index, with its text bodies read as one where the save joins
 * them (`cellAsSaved`). Read only: `selectionBody` gives the body that an edit writes.
 */
export function cellTxBody(frame: Element, at: TableCellAddress): Element | undefined {
  const tbl = tblOf(frame);
  const tc = children(children(tbl, 'a:tr')[at.row], 'a:tc')[at.column];
  return child(tc && tableCellCounts(tbl) ? cellAsSaved(tc) : tc, 'a:txBody');
}

/**
 * Give a row of a marked table the blank cells that the save adds to it, up to `at`, and return
 * the cell at `at`. Concurrent structure edits can leave a row short; the canvas draws the cells
 * the save adds (`tableAsDrawn`), and text typed into one needs its node. Each new cell is the
 * one the canvas draws (`newCellSpec`, `addedTableCell`). Undefined, with nothing written, when
 * the save adds no cell at `at`. Two peers that type at once into added cells of one row each
 * write them: the row then holds more cells than the grid, and the save and the canvas add a
 * grid column for them, so both texts stay.
 */
function addSavedCells(
  batch: OpBatch,
  part: string,
  frame: Element,
  at: TableCellAddress
): Element | undefined {
  const saved = tableAsDrawn(frame, TABLE_MAX_COLUMNS);
  const added = saved?.cells[at.row]?.[at.column];
  if (!added || added.id) return undefined;
  const tr = children(tblOf(frame), 'a:tr')[at.row];
  const tcs = children(tr, 'a:tc');
  const index = tr.children.findIndex((n) => n === tcs[tcs.length - 1]) + 1;
  let id = '';
  for (let c = tcs.length; c <= at.column; c += 1)
    id = batch.insert({
      op: 'insertNode',
      part,
      parent: tr.id,
      index: index + (c - tcs.length),
      spec: newCellSpec(tcs[tcs.length - 1], {}),
    });
  return batch.element(part, id);
}

/** The text body a text selection edits: the shape's `p:txBody`, or the cell's `a:txBody`. */
export function selectionBody(
  batch: OpBatch,
  part: string,
  sel: Extract<EditorSelection, { kind: 'text' }>
): Element | undefined {
  const el = batch.element(part, sel.nodeId);
  if (sel.cell) {
    const tbl = tblOf(el);
    const tr = children(tbl, 'a:tr')[sel.cell.row];
    const tc = children(tr, 'a:tc')[sel.cell.column];
    if (!tc) return child(addSavedCells(batch, part, el, sel.cell), 'a:txBody');
    const [body, ...rest] = children(tc, 'a:txBody');
    // A cell with two text bodies draws as the save writes it, one body with the paragraphs of
    // both (`cellAsSaved`): the edit joins them in the store first, so the caret's paragraph
    // is the one the canvas shows.
    if (body && rest.length > 0 && tableCellCounts(tbl)) {
      joinTextBodies(batch, part, body, rest);
      return batch.element(part, body.id);
    }
    if (body) return body;
    // A drawn cell may have no text body (`a:txBody` is optional, 21.1.3.16): the edit gives
    // it an empty one, formatted like its row's first cell.
    const id = batch.insert({
      op: 'insertNode',
      part,
      parent: tc.id,
      index: 0,
      spec: emptyBodySpec(children(tr, 'a:tc')[0]),
    });
    return batch.element(part, id);
  }
  const body = el.children.find((c) => c.type === 'element' && c.kind === 'txBody');
  return body?.type === 'element' ? body : undefined;
}
