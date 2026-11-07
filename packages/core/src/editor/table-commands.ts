// Table commands: insert a table, insert and delete rows and columns, select a region, merge and
// split cells, and table and cell formatting. Each compiles to one store transaction against
// the table of the selection (`tableTarget`), then re-fits the frame to the laid-out table.

import { intAttr } from '../store/index.ts';
import type {
  EditorSelection,
  TableCellAddress,
  TableCommand,
  TableStyleOption,
  TextPoint,
} from '../contracts/index.ts';
import type { CommandContext, CommandEffect } from './commands.ts';
import { CommandRefusal, OpBatch } from '../automation/commands/op-batch.ts';
import { deleteShapes } from '../automation/commands/shape-commands.ts';
import { cellRangeOf, tableItem } from './table-cells.ts';
import { refitTable } from '../automation/commands/table-fit.ts';
import { insertTableFrame, tablePlaceholderOf } from '../automation/commands/table-insert.ts';
import { markStructureEdit } from '../automation/commands/table-edit-marks.ts';
import { repairTable } from '../automation/commands/table-repair.ts';
import { paragraphLength, paragraphsOf } from '../automation/commands/text-model.ts';
import {
  setCellAnchor,
  setCellFill,
  setCellMargins,
  setTableBorders,
  setTableStyle,
  toggleStyleOption,
} from '../automation/commands/table-format.ts';
import {
  anchorsIn,
  blockOf,
  cellBody,
  clearRange,
  setCellText,
  deleteColumns,
  deleteRows,
  expandRange,
  selectedRange,
  insertColumns,
  insertRows,
  MAX_ROWS,
  MAX_TABLE_SIZE,
  mergeRange,
  readTable,
  setAttr,
  splitColumns,
  splitRows,
  wholeTable,
  type CellRange,
  type TableModel,
} from '../automation/commands/table-model.ts';

type Compiled = { batch: OpBatch; effect: CommandEffect };

const OPTIONS: readonly TableStyleOption[] = [
  'firstRow',
  'lastRow',
  'bandRow',
  'firstCol',
  'lastCol',
  'bandCol',
];
const MAX_LENGTH = 51_206_400;

function slidePart(ctx: CommandContext): string {
  const ref = ctx.session.info.slides[ctx.slide];
  if (!ref) throw new CommandRefusal('disabled.noEditor');
  return ref.partName;
}

/** The table frame a selection is in or on, or undefined. */
export function tableFrameOf(sel: EditorSelection): string | undefined {
  if (sel.kind === 'text') return sel.cell ? sel.nodeId : undefined;
  if (sel.kind === 'cells') return sel.nodeId;
  if (sel.kind === 'shapes' && sel.nodeIds.length === 1) return sel.nodeIds[0];
  return undefined;
}

/** The table of the selection and the cells it covers, grown over merged cells. */
export function tableTarget(
  ctx: CommandContext,
  batch: OpBatch
): { m: TableModel; range: CellRange } {
  const sel = ctx.selection;
  const id = tableFrameOf(sel);
  if (!id)
    throw new CommandRefusal(sel.kind === 'none' ? 'disabled.noSelection' : 'disabled.noTable');
  const part = slidePart(ctx);
  // A probe repairs too (its batch is thrown away), so `can` agrees with `exec`.
  repairTable(batch, part, id);
  const m = readTable(batch, part, id);
  const range =
    sel.kind === 'text' && sel.cell
      ? blockOf(m, sel.cell)
      : sel.kind === 'cells'
        ? selectedRange(m, cellRangeOf(sel))
        : wholeTable(m);
  return { m, range };
}

/** Whether the selection is in or on a table frame. */
export function isTableSelection(ctx: CommandContext): boolean {
  const id = tableFrameOf(ctx.selection);
  return !!id && !!tableItem(ctx.session.layout(ctx.slide), id);
}

/**
 * The text bodies a text command formats in a table: the cell of a caret (with its range),
 * every cell of a cell range, or every cell of a selected table.
 */
export function tableTextTargets(
  batch: OpBatch,
  part: string,
  sel: EditorSelection,
  frameId: string
): { txBodyId: string; shapeId: string; range?: { from: TextPoint; to: TextPoint } }[] {
  repairTable(batch, part, frameId);
  const m = readTable(batch, part, frameId);
  if (sel.kind === 'text' && sel.cell)
    return [
      {
        txBodyId: cellBody(m, sel.cell).id,
        shapeId: frameId,
        range: { from: sel.anchor, to: sel.focus },
      },
    ];
  const range = sel.kind === 'cells' ? selectedRange(m, cellRangeOf(sel)) : wholeTable(m);
  return anchorsIn(m, range).map((a) => ({ txBodyId: cellBody(m, a).id, shapeId: frameId }));
}

/** A caret at the start of a cell. */
const caretIn = (nodeId: string, cell: TableCellAddress): EditorSelection => ({
  kind: 'text',
  nodeId,
  cell,
  anchor: { paragraph: 0, offset: 0 },
  focus: { paragraph: 0, offset: 0 },
});

/** The inserted rows or columns, across the other lines of a cell range `sel`. */
function newLines(
  sel: Extract<EditorSelection, { kind: 'cells' }>,
  axis: 'row' | 'column',
  at: number,
  count: number
): EditorSelection {
  const r = cellRangeOf(sel);
  return axis === 'row'
    ? { ...sel, anchor: { row: at, column: r.c0 }, focus: { row: at + count - 1, column: r.c1 } }
    : { ...sel, anchor: { row: r.r0, column: at }, focus: { row: r.r1, column: at + count - 1 } };
}

/** The selection moved by whole rows or columns, as cells shift after an insert. */
function shifted(sel: EditorSelection, dr: number, dc: number): EditorSelection {
  if (sel.kind === 'text' && sel.cell)
    return { ...sel, cell: { row: sel.cell.row + dr, column: sel.cell.column + dc } };
  if (sel.kind === 'cells') {
    const move = (a: TableCellAddress) => ({ row: a.row + dr, column: a.column + dc });
    return { ...sel, anchor: move(sel.anchor), focus: move(sel.focus) };
  }
  return sel;
}

/**
 * How many rows (or columns) an insert adds: one per selected line of a cell range, and one for
 * a caret, even in a merged cell, or for a selected table. Tab in the last cell adds one row.
 */
function lineCount(sel: EditorSelection, lo: number, hi: number): number {
  return sel.kind === 'cells' ? hi - lo + 1 : 1;
}

/**
 * In a collaboration room, mark the table as one whose structure changed, so a save repairs
 * merge markup that a concurrent structure edit leaves (`table-edit-marks.ts`).
 */
function structureEdited(ctx: CommandContext, batch: OpBatch, m: TableModel): void {
  if (ctx.collaborating === true) markStructureEdit(batch, m.part, m.frame.id);
}

function refit(ctx: CommandContext, batch: OpBatch, m: TableModel): void {
  if (!ctx.probe) refitTable(batch, ctx.slide, m.part, m.frame.id, ctx.session.measurer);
}

/** Positive EMU from the chrome. */
function length(v: number, what: string): string {
  if (!Number.isFinite(v) || Math.round(v) < 1 || v > MAX_LENGTH)
    throw new CommandRefusal('errors.unsupported', what);
  return String(Math.round(v));
}

/**
 * A new table on the current slide (`automation/commands/table-insert.ts`). It fills the selected
 * empty content placeholder, or else the first one on the slide.
 */
function insertTable(ctx: CommandContext, batch: OpBatch, rows: number, cols: number): string {
  const { width: slideW, height: slideH } = ctx.session.info;
  const part = slidePart(ctx);
  const sel = ctx.selection;
  const preferred = sel.kind === 'shapes' && sel.nodeIds.length === 1 ? sel.nodeIds[0] : undefined;
  const into = tablePlaceholderOf(batch, part, ctx.slide, preferred);
  return insertTableFrame(batch, part, rows, cols, slideW, slideH, into);
}

/** The sum of the grid column widths, EMU. */
const gridWidth = (m: TableModel): number =>
  m.gridCols.reduce((sum, g) => sum + (intAttr(g, 'w') ?? 0), 0);

/**
 * Scale the grid columns back to `width` after an insert, so the table keeps its width and the
 * columns keep their proportions, as slide editors insert a column.
 */
function keepGridWidth(batch: OpBatch, m: TableModel, width: number): void {
  const after = readTable(batch, m.part, m.frame.id);
  const total = gridWidth(after);
  if (width <= 0 || total <= width) return;
  let left = width;
  after.gridCols.forEach((g, i) => {
    const last = i === after.gridCols.length - 1;
    const w = last ? left : Math.round(((intAttr(g, 'w') ?? 0) * width) / total);
    left -= w;
    setAttr(batch, m.part, g.id, 'w', String(Math.max(1, w)));
  });
}

/**
 * The table after a paste grows it: rows and columns are added at the bottom and right for the
 * text past its edge, as many as the table limits allow. New columns share the table's width.
 */
function grownForPaste(
  ctx: CommandContext,
  batch: OpBatch,
  m: TableModel,
  range: CellRange,
  texts: readonly unknown[]
): TableModel {
  const widest = texts.reduce<number>(
    (n, row) => (Array.isArray(row) ? Math.max(n, row.length) : n),
    0
  );
  const rows = Math.min(range.r0 + texts.length - m.nR, MAX_ROWS - m.nR);
  const columns = Math.min(range.c0 + widest - m.nC, MAX_TABLE_SIZE - m.nC);
  if (rows <= 0 && columns <= 0) return m;
  // Plain cells: a merged edge row or column is not repeated, so every pasted value has a cell.
  if (rows > 0) insertRows(batch, m, m.nR, rows, m.nR - 1, true);
  let grown = readTable(batch, m.part, m.frame.id);
  if (columns > 0) {
    const width = gridWidth(grown);
    insertColumns(batch, grown, grown.nC, columns, grown.nC - 1, true);
    keepGridWidth(batch, grown, width);
    grown = readTable(batch, m.part, m.frame.id);
  }
  structureEdited(ctx, batch, grown);
  return grown;
}

/** The most characters one cell takes from a paste, as the clipboard parser reads. */
const MAX_PASTED_CELL_CHARS = 32_768;

/**
 * The cells a region command selects. A row or column region names only its own grid lines,
 * as slide editors select a column: a merged cell that crosses into the next column does not
 * widen it, so Delete Columns then deletes that one column and the merged cell shrinks.
 */
function region(m: TableModel, range: CellRange, which: 'table' | 'row' | 'column'): CellRange {
  if (which === 'table') return wholeTable(m);
  if (which === 'row') return { ...range, c0: 0, c1: m.nC - 1 };
  return { ...range, r0: 0, r1: m.nR - 1 };
}

/**
 * The rows (or columns) a delete removes: the cell range's own grid lines on that axis, not the
 * merged cells it touches; a caret's merged cell, or a selected table, as `tableTarget` reads.
 */
function deletedLines(ctx: CommandContext, range: CellRange, axis: 'row' | 'column'): CellRange {
  const sel = ctx.selection;
  if (sel.kind !== 'cells') return range;
  const own = cellRangeOf(sel);
  return axis === 'row'
    ? { ...range, r0: own.r0, r1: own.r1 }
    : { ...range, c0: own.c0, c1: own.c1 };
}

const cells = (nodeId: string, r: CellRange): EditorSelection => ({
  kind: 'cells',
  nodeId,
  anchor: { row: r.r0, column: r.c0 },
  focus: { row: r.r1, column: r.c1 },
});

/** Delete rows or columns; the whole table goes when none would remain. */
function deleteLines(ctx: CommandContext, batch: OpBatch, axis: 'row' | 'column'): Compiled {
  const target = tableTarget(ctx, batch);
  const m = target.m;
  const range = deletedLines(ctx, target.range, axis);
  const id = m.frame.id;
  const all =
    axis === 'row'
      ? range.r0 === 0 && range.r1 === m.nR - 1
      : range.c0 === 0 && range.c1 === m.nC - 1;
  if (all) {
    deleteShapes(batch, m.part, [id]);
    return { batch, effect: { label: 'deleteTable', selection: { kind: 'none' } } };
  }
  if (axis === 'row') deleteRows(batch, m, range.r0, range.r1);
  else deleteColumns(batch, m, range.c0, range.c1);
  structureEdited(ctx, batch, m);
  refit(ctx, batch, m);
  // The caret goes to the cell that takes the deleted cells' place.
  const after = readTable(batch, m.part, id);
  const row = axis === 'row' ? Math.min(range.r0, after.nR - 1) : range.r0;
  const column = axis === 'column' ? Math.min(range.c0, after.nC - 1) : range.c0;
  const a = after.anchors[Math.min(row, after.nR - 1)][Math.min(column, after.nC - 1)];
  return {
    batch,
    effect: { label: axis === 'row' ? 'deleteRow' : 'deleteColumn', selection: caretIn(id, a) },
  };
}

/** Compile a table command against a fresh batch. Throws CommandRefusal. */
export function compileTable(ctx: CommandContext, batch: OpBatch, command: TableCommand): Compiled {
  switch (command.type) {
    case 'insertTable': {
      const id = insertTable(ctx, batch, command.rows, command.cols);
      return {
        batch,
        effect: { label: 'insertTable', selection: caretIn(id, { row: 0, column: 0 }) },
      };
    }
    case 'insertRow': {
      const { m, range } = tableTarget(ctx, batch);
      const count = lineCount(ctx.selection, range.r0, range.r1);
      const above = command.where === 'above';
      const at = above ? range.r0 : range.r1 + 1;
      insertRows(batch, m, at, count, above ? range.r0 : range.r1);
      structureEdited(ctx, batch, m);
      refit(ctx, batch, m);
      // A cell range selects the new rows, ready to fill; a caret stays in its cell.
      const sel = ctx.selection;
      const selection =
        sel.kind === 'cells'
          ? newLines(sel, 'row', at, count)
          : above
            ? shifted(sel, count, 0)
            : sel;
      return { batch, effect: { label: 'insertRow', selection } };
    }
    case 'insertColumn': {
      const { m, range } = tableTarget(ctx, batch);
      const count = lineCount(ctx.selection, range.c0, range.c1);
      const left = command.where === 'left';
      const width = gridWidth(m);
      const at = left ? range.c0 : range.c1 + 1;
      insertColumns(batch, m, at, count, left ? range.c0 : range.c1);
      keepGridWidth(batch, m, width);
      structureEdited(ctx, batch, m);
      refit(ctx, batch, m);
      // A cell range selects the new columns, ready to fill; a caret stays in its cell.
      const sel = ctx.selection;
      const selection =
        sel.kind === 'cells'
          ? newLines(sel, 'column', at, count)
          : left
            ? shifted(sel, 0, count)
            : sel;
      return { batch, effect: { label: 'insertColumn', selection } };
    }
    case 'deleteRow':
      return deleteLines(ctx, batch, 'row');
    case 'deleteColumn':
      return deleteLines(ctx, batch, 'column');
    case 'deleteTable': {
      const { m } = tableTarget(ctx, batch);
      deleteShapes(batch, m.part, [m.frame.id]);
      return { batch, effect: { label: 'deleteTable', selection: { kind: 'none' } } };
    }
    case 'selectTableRegion': {
      // A selection writes nothing: the repair goes to a batch that is thrown away.
      const { m, range } = tableTarget(ctx, new OpBatch(batch.pkg));
      if (!['table', 'row', 'column'].includes(command.region))
        throw new CommandRefusal('errors.unsupported', 'region');
      return {
        batch,
        effect: { label: 'select', selection: cells(m.frame.id, region(m, range, command.region)) },
      };
    }
    case 'mergeCells': {
      const target = tableTarget(ctx, batch);
      const m = target.m;
      // A merge is a rectangle: a band grows over the merged cells that cross it.
      const range = expandRange(m, target.range);
      mergeRange(batch, m, range);
      structureEdited(ctx, batch, m);
      refit(ctx, batch, m);
      // The caret goes after the merged text, so the next key adds to it.
      const at = { row: range.r0, column: range.c0 };
      const paragraphs = paragraphsOf(cellBody(readTable(batch, m.part, m.frame.id), at));
      const end = {
        paragraph: Math.max(0, paragraphs.length - 1),
        offset: paragraphs.length ? paragraphLength(paragraphs[paragraphs.length - 1]) : 0,
      };
      return {
        batch,
        effect: {
          label: 'mergeCells',
          selection: { kind: 'text', nodeId: m.frame.id, cell: at, anchor: end, focus: end },
        },
      };
    }
    case 'splitCell': {
      const { rows, cols } = command;
      if (![rows, cols].every((n) => Number.isInteger(n) && n >= 1 && n <= MAX_TABLE_SIZE))
        throw new CommandRefusal('errors.unsupported', 'size');
      if (rows === 1 && cols === 1) throw new CommandRefusal('errors.unsupported', 'size');
      const { m, range } = tableTarget(ctx, batch);
      if (anchorsIn(m, range).length !== 1) throw new CommandRefusal('disabled.splitOneCell');
      const at = { row: range.r0, column: range.c0 };
      const area = splitColumns(batch, m, at, cols);
      splitRows(batch, readTable(batch, m.part, m.frame.id), area, rows);
      structureEdited(ctx, batch, m);
      refit(ctx, batch, m);
      return { batch, effect: { label: 'splitCell', selection: caretIn(m.frame.id, at) } };
    }
    case 'setCellFill': {
      const { m, range } = tableTarget(ctx, batch);
      setCellFill(batch, m, range, command.color);
      return { batch, effect: { label: 'cellFill' } };
    }
    case 'setTableBorders': {
      const { m, range } = tableTarget(ctx, batch);
      setTableBorders(batch, m, range, command.scope, 'spec' in command ? command.spec : undefined);
      return { batch, effect: { label: 'tableBorders' } };
    }
    case 'setTableStyle': {
      const { m } = tableTarget(ctx, batch);
      setTableStyle(batch, m, command.styleId);
      refit(ctx, batch, m);
      return { batch, effect: { label: 'tableStyle' } };
    }
    case 'toggleTableStyleOption': {
      if (!OPTIONS.includes(command.option))
        throw new CommandRefusal('errors.unsupported', 'option');
      const { m } = tableTarget(ctx, batch);
      toggleStyleOption(batch, m, command.option);
      refit(ctx, batch, m);
      return { batch, effect: { label: 'tableStyleOption' } };
    }
    case 'setCellMargins': {
      const { m, range } = tableTarget(ctx, batch);
      setCellMargins(batch, m, range, command.margins);
      refit(ctx, batch, m);
      return { batch, effect: { label: 'cellMargins' } };
    }
    case 'setRowHeight': {
      const { m, range } = tableTarget(ctx, batch);
      const h = length(command.height, 'height');
      for (let r = range.r0; r <= range.r1; r += 1) setAttr(batch, m.part, m.rows[r].id, 'h', h);
      refit(ctx, batch, m);
      return { batch, effect: { label: 'rowHeight' } };
    }
    case 'setColumnWidth': {
      const { m, range } = tableTarget(ctx, batch);
      const w = length(command.width, 'width');
      for (let c = range.c0; c <= range.c1; c += 1)
        setAttr(batch, m.part, m.gridCols[c].id, 'w', w);
      refit(ctx, batch, m);
      return { batch, effect: { label: 'columnWidth' } };
    }
    case 'pasteCells': {
      const target = tableTarget(ctx, batch);
      const range = target.range;
      const texts = command.texts;
      if (!Array.isArray(texts) || texts.length === 0)
        throw new CommandRefusal('errors.unsupported', 'texts');
      const m = grownForPaste(ctx, batch, target.m, range, texts);
      let r1 = range.r0;
      let c1 = range.c0;
      // Text past the largest table is left out: only the rows and columns that fit are read.
      texts.slice(0, m.nR - range.r0).forEach((row, i) => {
        if (!Array.isArray(row)) throw new CommandRefusal('errors.unsupported', 'texts');
        row.slice(0, m.nC - range.c0).forEach((text, j) => {
          if (typeof text !== 'string') throw new CommandRefusal('errors.unsupported', 'texts');
          const r = range.r0 + i;
          const c = range.c0 + j;
          const a = m.anchors[r][c];
          if (a.row !== r || a.column !== c) return;
          setCellText(batch, m, a, text.slice(0, MAX_PASTED_CELL_CHARS));
          r1 = Math.max(r1, r);
          c1 = Math.max(c1, c);
        });
      });
      refit(ctx, batch, m);
      const filled = { r0: range.r0, c0: range.c0, r1, c1 };
      return { batch, effect: { label: 'pasteCells', selection: cells(m.frame.id, filled) } };
    }
    case 'distributeRows':
    case 'distributeColumns': {
      const { m, range } = tableTarget(ctx, batch);
      const rowsAxis = command.type === 'distributeRows';
      const laid = tableItem(ctx.session.layout(ctx.slide), m.frame.id)?.table;
      if (!laid) throw new CommandRefusal('disabled.noTable');
      // One cell distributes the whole table.
      const [lo, hi] = rowsAxis
        ? range.r0 === range.r1
          ? [0, m.nR - 1]
          : [range.r0, range.r1]
        : range.c0 === range.c1
          ? [0, m.nC - 1]
          : [range.c0, range.c1];
      // Laid-out sizes (rows grow to fit text) when the layout still matches the repaired
      // table, else the file's sizes; column widths are the grid's.
      const sizes = rowsAxis
        ? laid.rowY.length === m.nR + 1
          ? laid.rowY.slice(1).map((y, i) => y - laid.rowY[i])
          : m.rows.map((tr) => intAttr(tr, 'h') ?? 0)
        : m.gridCols.map((g) => intAttr(g, 'w') ?? 0);
      let total = 0;
      for (let i = lo; i <= hi; i += 1) total += sizes[i];
      const each = Math.round(total / (hi - lo + 1));
      if (!(each > 0)) throw new CommandRefusal('errors.unsupported', 'distribute');
      for (let i = lo; i <= hi; i += 1) {
        const node = rowsAxis ? m.rows[i] : m.gridCols[i];
        setAttr(batch, m.part, node.id, rowsAxis ? 'h' : 'w', String(each));
      }
      refit(ctx, batch, m);
      return { batch, effect: { label: command.type } };
    }
    case 'setTableGrid': {
      const { m } = tableTarget(ctx, batch);
      if (command.columns && command.columns.length !== m.nC)
        throw new CommandRefusal('errors.unsupported', 'columns');
      if (command.rows && command.rows.length !== m.nR)
        throw new CommandRefusal('errors.unsupported', 'rows');
      command.columns?.forEach((w, c) =>
        setAttr(batch, m.part, m.gridCols[c].id, 'w', length(w, 'width'))
      );
      command.rows?.forEach((h, r) =>
        setAttr(batch, m.part, m.rows[r].id, 'h', length(h, 'height'))
      );
      refit(ctx, batch, m);
      return { batch, effect: { label: 'tableGrid' } };
    }
    default:
      // The last case of `compile`: a host without type checks can send any `type`. Refuse it
      // with a reason; never fall through to a missing result.
      throw new CommandRefusal('errors.unsupported', String((command as { type?: unknown }).type));
  }
}

/** Clear the text of a cell range (Delete with cells selected). */
export function clearCells(ctx: CommandContext, batch: OpBatch): Compiled {
  const { m, range } = tableTarget(ctx, batch);
  clearRange(batch, m, range);
  refit(ctx, batch, m);
  return { batch, effect: { label: 'clearCells' } };
}

/** Vertical alignment for a table selection; false when the selection is not in a table. */
export function setTableAnchor(
  ctx: CommandContext,
  batch: OpBatch,
  anchor: 't' | 'ctr' | 'b'
): boolean {
  if (!isTableSelection(ctx)) return false;
  const { m, range } = tableTarget(ctx, batch);
  setCellAnchor(batch, m, range, anchor);
  return true;
}
