// Tables (`a:tbl` in a graphic frame, ECMA-376 Part 1, 21.1.3.13): values, cells, rows,
// columns, merges, and sizes. The grid edits are the editor's own (`commands/table-model.ts`);
// a frame on a slide re-fits its rows to their text when the host has a measurer.

import { attr, boolAttr, child, elements, intAttr, NS, type Element } from '../store/index.ts';
import { tableStyleFor } from '../cascade/table-styles.ts';
import { refitTable } from './commands/table-fit.ts';
import { setCellAnchor, setCellFillXml } from './commands/table-format.ts';
import {
  blockOf,
  cellBodyAsSaved,
  clearRange,
  deleteColumns,
  expandRange,
  MAX_ROWS,
  MAX_TABLE_SIZE,
  deleteRows,
  insertColumns,
  insertRows,
  joinCellBodies,
  mergeRange,
  readTable,
  splitColumns,
  splitRows,
  tblOf,
  writeBlock,
  type CellRange,
  type TableModel,
} from './commands/table-model.ts';
import { deleteShapes } from './commands/shape-commands.ts';
import { markStructureEdit } from './commands/table-edit-marks.ts';
import type { AutomationCellAddress, AutomationFillWrite } from './operations.ts';
import type {
  AutomationFillRead,
  AutomationRef,
  AutomationTableRead,
  AutomationValue,
} from './protocol.ts';
import { fillRead, resolvedShape } from './plan-reads.ts';
import { bodyTextOf, colorSpec, replaceBodyText, type TextPlanContext } from './plan-text.ts';
import { slideNumberOf } from './commands/text-model.ts';
import {
  AutomationFailure,
  badArgument,
  edit,
  indexIn,
  numberIn,
  ownerSlideIndex,
  shapeAt,
  toEmu,
  toPoints,
  type ShapeLocation,
} from './targets.ts';

/**
 * The table of a shape, or an `unsupported-content` failure for any other shape. An edit
 * (`read` false) first joins the text bodies of each cell with more than one
 * (`joinCellBodies`), so it keeps the text of every body; a read writes nothing and reads such a
 * cell as the save writes it (`cellBodyAsSaved`).
 */
export function tableAt(
  ctx: TextPlanContext,
  ref: AutomationRef,
  read = false
): { at: ShapeLocation; m: TableModel } {
  const at = shapeAt(ctx, ref);
  if (!tblOf(at.el))
    throw new AutomationFailure('unsupported-content', 'that shape is not a table', 'not-table');
  if (!read) edit(() => joinCellBodies(ctx.batch, at.part, at.el.id));
  return { at, m: edit(() => readTable(ctx.batch, at.part, at.el.id)) };
}

/**
 * In a collaboration room, mark the table as one whose structure changed, so a save repairs the
 * merge markup that a concurrent structure edit leaves, as the editor's table commands do
 * (`commands/table-edit-marks.ts`).
 */
function structureEdited(ctx: TextPlanContext, at: ShapeLocation): void {
  if (ctx.collaborating === true) edit(() => markStructureEdit(ctx.batch, at.part, at.el.id));
}

/** Re-fit a slide table's rows to their text. */
export function refit(ctx: TextPlanContext, at: ShapeLocation): void {
  const index = ownerSlideIndex(ctx.batch.pkg, at.owner);
  if (index === undefined || !ctx.measurer) return;
  const measurer = ctx.measurer;
  edit(() => refitTable(ctx.batch, index, at.part, at.el.id, measurer));
}

/** The text of the cell at a grid position: the merged block's text, or '' when covered. */
function cellText(ctx: TextPlanContext, m: TableModel, row: number, column: number): string {
  const anchor = m.anchors[row]![column]!;
  if (anchor.row !== row || anchor.column !== column) return '';
  const txBody = cellBodyAsSaved(m.tbl, m.tcs[row]![column]!);
  return txBody ? bodyTextOf(txBody, slideNumberOf(ctx.batch, m.part)) : '';
}

export function readTableValue(ctx: TextPlanContext, ref: AutomationRef): AutomationValue {
  const { m } = tableAt(ctx, ref, true);
  const tblPr = child(m.tbl, 'a:tblPr');
  const merged: AutomationTableRead['mergedAreas'][number][] = [];
  for (let r = 0; r < m.nR; r += 1)
    for (let c = 0; c < m.nC; c += 1) {
      const a = m.anchors[r]![c]!;
      if (a.row !== r || a.column !== c) continue;
      const block = blockOf(m, a);
      if (block.r1 > block.r0 || block.c1 > block.c0)
        merged.push({
          rowIndex: r,
          columnIndex: c,
          rowCount: block.r1 - block.r0 + 1,
          columnCount: block.c1 - block.c0 + 1,
        });
    }
  const flag = (name: string): boolean => boolAttr(tblPr, name) ?? false;
  const table: AutomationTableRead = {
    rowCount: m.nR,
    columnCount: m.nC,
    values: m.rows.map((_, r) => m.gridCols.map((__, c) => cellText(ctx, m, r, c))),
    columnWidths: m.gridCols.map((g) => toPoints(intAttr(g, 'w') ?? 0)),
    rowHeights: m.rows.map((row) => toPoints(intAttr(row, 'h') ?? 0)),
    mergedAreas: merged,
    style: {
      styleId: child(tblPr, 'a:tableStyleId') ? textIn(child(tblPr, 'a:tableStyleId')) : null,
      styleName: attr(tableStyleFor(ctx.batch.pkg, tblPr), 'styleName') || null,
      firstRow: flag('firstRow'),
      lastRow: flag('lastRow'),
      firstColumn: flag('firstCol'),
      lastColumn: flag('lastCol'),
      bandedRows: flag('bandRow'),
      bandedColumns: flag('bandCol'),
    },
  };
  return { kind: 'table', table };
}

function textIn(el: ReturnType<typeof child>): string {
  return (el?.children ?? [])
    .map((c) => (c.type === 'text' ? c.value : ''))
    .join('')
    .trim();
}

export function readTableCell(
  ctx: TextPlanContext,
  ref: AutomationRef,
  row: number,
  column: number
): AutomationValue {
  const { m } = tableAt(ctx, ref, true);
  if (
    !Number.isInteger(row) ||
    !Number.isInteger(column) ||
    row < 0 ||
    column < 0 ||
    row >= m.nR ||
    column >= m.nC
  )
    return { kind: 'null' };
  const anchor = m.anchors[row]![column]!;
  const block = blockOf(m, anchor);
  const tcPr = child(m.tcs[anchor.row]![anchor.column], 'a:tcPr');
  return {
    kind: 'tableCell',
    cell: {
      text: cellText(ctx, m, anchor.row, anchor.column),
      rowIndex: anchor.row,
      columnIndex: anchor.column,
      rowCount: block.r1 - block.r0 + 1,
      columnCount: block.c1 - block.c0 + 1,
      // CT_TableCellProperties defaults (ECMA-376 Part 1, 21.1.3.17): `t`, not centered.
      anchor: attr(tcPr, 'anchor') ?? 't',
      anchorCenter: boolAttr(tcPr, 'anchorCtr') ?? false,
    },
  };
}

/** Write one cell's text: the anchor cell of the block that covers the position. */
function writeCell(
  ctx: TextPlanContext,
  at: ShapeLocation,
  m: TableModel,
  row: number,
  column: number,
  text: string
): void {
  const anchor = m.anchors[row]![column]!;
  const tc = m.tcs[anchor.row]![anchor.column]!;
  const txBody = child(tc, 'a:txBody');
  if (!txBody) throw new AutomationFailure('unsupported-content', 'a cell with no text body');
  replaceBodyText(ctx, at, txBody.id, text);
}

export function setTableCell(
  ctx: TextPlanContext,
  ref: AutomationRef,
  row: number,
  column: number,
  text: string
): AutomationValue {
  const { at, m } = tableAt(ctx, ref);
  writeCell(ctx, at, m, indexIn(row, 'row', m.nR), indexIn(column, 'column', m.nC), text);
  refit(ctx, at);
  return { kind: 'applied' };
}

/** Write row-major values into a table, from its top-left cell. */
export function writeCells(
  ctx: TextPlanContext,
  ref: AutomationRef,
  values: readonly (readonly string[])[]
): void {
  if (!Array.isArray(values)) badArgument('values');
  const { at, m } = tableAt(ctx, ref);
  if (values.length > m.nR) badArgument('values');
  values.forEach((row, r) => {
    if (!Array.isArray(row) || row.length > m.nC) badArgument('values');
    row.forEach((text, c) => {
      if (typeof text !== 'string') badArgument('values');
      // Read again: a write replaces the cell's paragraphs.
      const fresh = edit(() => readTable(ctx.batch, at.part, at.el.id));
      const anchor = fresh.anchors[r]![c]!;
      if (anchor.row === r && anchor.column === c) writeCell(ctx, at, fresh, r, c, text);
    });
  });
  refit(ctx, at);
}

export function insertTableLines(
  ctx: TextPlanContext,
  ref: AutomationRef,
  axis: 'row' | 'column',
  index: number | undefined,
  count: number
): AutomationValue {
  const { at, m } = tableAt(ctx, ref);
  const n = axis === 'row' ? m.nR : m.nC;
  const at0 = index === undefined ? n : indexIn(index, 'index', n + 1);
  const many = numberIn(count, 'count', 1, 1000);
  if (!Number.isInteger(many)) badArgument('count');
  // New lines copy the format of the line before them, or of the first line at the start.
  const source = Math.max(0, at0 - 1);
  edit(() =>
    axis === 'row'
      ? insertRows(ctx.batch, m, at0, many, source)
      : insertColumns(ctx.batch, m, at0, many, source)
  );
  structureEdited(ctx, at);
  refit(ctx, at);
  return { kind: 'applied' };
}

export function deleteTableLines(
  ctx: TextPlanContext,
  ref: AutomationRef,
  axis: 'row' | 'column',
  index: number,
  count: number
): AutomationValue {
  const { at, m } = tableAt(ctx, ref);
  const n = axis === 'row' ? m.nR : m.nC;
  const first = indexIn(index, 'index', n);
  const many = numberIn(count, 'count', 1, n - first);
  if (!Number.isInteger(many)) badArgument('count');
  edit(() => {
    // Deleting every line deletes the table, as the editor does.
    if (first === 0 && many === n) return deleteShapes(ctx.batch, at.part, [at.el.id]);
    if (axis === 'row') deleteRows(ctx.batch, m, first, first + many - 1);
    else deleteColumns(ctx.batch, m, first, first + many - 1);
  });
  if (first === 0 && many === n) return { kind: 'applied' };
  structureEdited(ctx, at);
  refit(ctx, at);
  return { kind: 'applied' };
}

export function mergeTableCells(
  ctx: TextPlanContext,
  ref: AutomationRef,
  row: number,
  column: number,
  rowCount: number,
  columnCount: number
): AutomationValue {
  const { at, m } = tableAt(ctx, ref);
  const r0 = indexIn(row, 'row', m.nR);
  const c0 = indexIn(column, 'column', m.nC);
  const rows = numberIn(rowCount, 'rowCount', 1, m.nR - r0);
  const cols = numberIn(columnCount, 'columnCount', 1, m.nC - c0);
  if (!Number.isInteger(rows) || !Number.isInteger(cols)) badArgument('rowCount');
  const range = { r0, r1: r0 + rows - 1, c0, c1: c0 + cols - 1 };
  edit(() => mergeRange(ctx.batch, m, range));
  structureEdited(ctx, at);
  refit(ctx, at);
  return { kind: 'applied' };
}

export function setTableSize(
  ctx: TextPlanContext,
  ref: AutomationRef,
  target: 'row' | 'column',
  index: number,
  size: number
): AutomationValue {
  const { at, m } = tableAt(ctx, ref);
  const line = target === 'row' ? m.rows : m.gridCols;
  const el = line[indexIn(index, 'index', line.length)]!;
  const emu = toEmu(numberIn(size, 'lineSize', 0, 4032));
  ctx.batch.push({
    op: 'setAttribute',
    part: at.part,
    node: el.id,
    name: { ns: '', local: target === 'row' ? 'h' : 'w' },
    value: String(emu),
  });
  refit(ctx, at);
  return { kind: 'applied' };
}

/** The block of the cell at a checked grid position. */
function blockAt(m: TableModel, row: number, column: number): CellRange {
  const r = indexIn(row, 'row', m.nR);
  const c = indexIn(column, 'column', m.nC);
  return blockOf(m, m.anchors[r]![c]!);
}

/** A whole count in [1, max], or `bad-argument` naming `name`. */
function countIn(value: number, name: string, max: number): number {
  const n = numberIn(value, name, 1, max);
  if (!Number.isInteger(n)) badArgument(name);
  return n;
}

export function setTableCellAnchor(
  ctx: TextPlanContext,
  ref: AutomationRef,
  row: number,
  column: number,
  anchor: string,
  anchorCenter: boolean
): AutomationValue {
  const { at, m } = tableAt(ctx, ref);
  if (anchor !== 't' && anchor !== 'ctr' && anchor !== 'b') badArgument('anchor');
  if (typeof anchorCenter !== 'boolean') badArgument('anchorCenter');
  edit(() => setCellAnchor(ctx.batch, m, blockAt(m, row, column), anchor, anchorCenter));
  refit(ctx, at);
  return { kind: 'applied' };
}

/**
 * Split the cell's block into `rowCount` by `columnCount` cells, as the editor's Split Cells
 * does: along grid lines when the block divides evenly, else into new grid lines.
 */
export function splitTableCell(
  ctx: TextPlanContext,
  ref: AutomationRef,
  row: number,
  column: number,
  rowCount: number,
  columnCount: number
): AutomationValue {
  const { at, m } = tableAt(ctx, ref);
  const block = blockAt(m, row, column);
  const rows = countIn(rowCount, 'rowCount', MAX_ROWS);
  const cols = countIn(columnCount, 'columnCount', MAX_TABLE_SIZE);
  if (rows === 1 && cols === 1) return { kind: 'applied' };
  edit(() => {
    const area = splitColumns(ctx.batch, m, { row: block.r0, column: block.c0 }, cols);
    splitRows(ctx.batch, readTable(ctx.batch, at.part, at.el.id), area, rows);
  });
  structureEdited(ctx, at);
  refit(ctx, at);
  return { kind: 'applied' };
}

/**
 * Make the cell span `rowCount` rows and `columnCount` columns from its top-left grid position.
 * The old block comes apart first; the new one must not cut another merged block.
 */
export function resizeTableCell(
  ctx: TextPlanContext,
  ref: AutomationRef,
  row: number,
  column: number,
  rowCount: number,
  columnCount: number
): AutomationValue {
  const { at, m } = tableAt(ctx, ref);
  const block = blockAt(m, row, column);
  const rows = countIn(rowCount, 'rowCount', m.nR - block.r0);
  const cols = countIn(columnCount, 'columnCount', m.nC - block.c0);
  const target: CellRange = {
    r0: block.r0,
    c0: block.c0,
    r1: block.r0 + rows - 1,
    c1: block.c0 + cols - 1,
  };
  edit(() => {
    for (let r = block.r0; r <= block.r1; r += 1)
      for (let c = block.c0; c <= block.c1; c += 1)
        writeBlock(ctx.batch, m, { r0: r, r1: r, c0: c, c1: c });
    if (rows === 1 && cols === 1) return;
    const fresh = readTable(ctx.batch, at.part, at.el.id);
    const whole = expandRange(fresh, target);
    if (
      whole.r0 !== target.r0 ||
      whole.c0 !== target.c0 ||
      whole.r1 !== target.r1 ||
      whole.c1 !== target.c1
    )
      throw new AutomationFailure(
        'unsupported-content',
        'the new span cuts another merged cell',
        'merge-overlap'
      );
    mergeRange(ctx.batch, fresh, target);
  });
  structureEdited(ctx, at);
  refit(ctx, at);
  return { kind: 'applied' };
}

/** `a:rPr` children that are links, not format (CT_TextCharacterProperties, 21.1.2.3.9). */
const RUN_LINKS = new Set(['hlinkClick', 'hlinkMouseOver']);

/**
 * Remove the format of a run's `a:rPr`. A run with a link keeps its `a:rPr` and the link
 * elements in it, so a format clear never removes a hyperlink.
 */
function clearRunFormat(ctx: TextPlanContext, part: string, rPr: Element): void {
  const links = elements(rPr).filter((e) => e.ns === NS.a && RUN_LINKS.has(e.local));
  if (links.length === 0) {
    ctx.batch.push({ op: 'removeNode', part, node: rPr.id });
    return;
  }
  for (const a of rPr.attrs)
    ctx.batch.push({
      op: 'setAttribute',
      part,
      node: rPr.id,
      name: { ns: a.ns, local: a.local },
      value: null,
    });
  for (const e of elements(rPr))
    if (!links.includes(e)) ctx.batch.push({ op: 'removeNode', part, node: e.id });
}

/** Remove the direct format of a cell: its `a:tcPr`, paragraph properties, and run properties. */
function clearCellFormat(ctx: TextPlanContext, part: string, tc: Element): void {
  const tcPr = child(tc, 'a:tcPr');
  if (tcPr) ctx.batch.push({ op: 'removeNode', part, node: tcPr.id });
  for (const p of elements(child(tc, 'a:txBody')).filter((e) => e.local === 'p'))
    for (const e of elements(p)) {
      if (e.local === 'pPr' || e.local === 'endParaRPr')
        ctx.batch.push({ op: 'removeNode', part, node: e.id });
      else {
        const rPr = child(e, 'a:rPr');
        if (rPr) clearRunFormat(ctx, part, rPr);
      }
    }
}

/** Remove the text, the direct format, or both, of every cell of a table. */
export function clearTable(
  ctx: TextPlanContext,
  ref: AutomationRef,
  text: boolean,
  format: boolean
): AutomationValue {
  if (typeof text !== 'boolean' || typeof format !== 'boolean') badArgument('options');
  const { at, m } = tableAt(ctx, ref);
  const all: CellRange = { r0: 0, c0: 0, r1: m.nR - 1, c1: m.nC - 1 };
  edit(() => {
    if (text) clearRange(ctx.batch, m, all);
    if (format) {
      const fresh = readTable(ctx.batch, at.part, at.el.id);
      for (const tcs of fresh.tcs) for (const tc of tcs) clearCellFormat(ctx, at.part, tc);
    }
  });
  refit(ctx, at);
  return { kind: 'applied' };
}

/** The block that covers a checked cell address. */
export function cellBlock(m: TableModel, cell: AutomationCellAddress): CellRange {
  if (typeof cell !== 'object' || cell === null) badArgument('cell');
  return blockAt(m, cell.row, cell.column);
}

/** The fill of a cell: the cascade's (direct, then table style) on a slide, else the file's. */
export function readTableCellFill(
  ctx: TextPlanContext,
  ref: AutomationRef,
  cell: AutomationCellAddress
): AutomationValue {
  const { at, m } = tableAt(ctx, ref, true);
  const block = cellBlock(m, cell);
  const tc = m.tcs[block.r0]![block.c0]!;
  const resolved = resolvedShape(ctx, at);
  const hit =
    resolved?.kind === 'graphicFrame'
      ? resolved.table?.cells.find((c) => c.nodeId === tc.id)
      : undefined;
  if (hit) return { kind: 'fill', fill: fillRead(hit.fill) };
  const tcPr = child(tc, 'a:tcPr');
  const solid = child(tcPr, 'a:solidFill');
  const color = attr(child(solid, 'a:srgbClr'), 'val');
  const fill: AutomationFillRead = {
    type: solid ? 'solid' : 'noFill',
    foregroundColor: color ? `#${color.toUpperCase()}` : null,
    transparency: solid ? 0 : null,
  };
  return { kind: 'fill', fill };
}

/** Set a cell's fill (`a:tcPr`, ECMA-376 Part 1, 21.1.3.17): solid, none, or a transparency. */
export function setTableCellFill(
  ctx: TextPlanContext,
  ref: AutomationRef,
  cell: AutomationCellAddress,
  fill: AutomationFillWrite
): AutomationValue {
  if (typeof fill !== 'object' || fill === null) badArgument('fill');
  const { m } = tableAt(ctx, ref);
  const block = cellBlock(m, cell);
  let xml: string;
  if (fill.kind === 'clear') xml = '<a:noFill/>';
  else if (fill.kind === 'solid')
    xml = `<a:solidFill>${colorSpec(fill.color, 'color', fill.transparency)}</a:solidFill>`;
  else if (fill.kind === 'transparency') {
    const current = readTableCellFill(ctx, ref, cell);
    const color = current.kind === 'fill' ? current.fill.foregroundColor : null;
    if (!color)
      throw new AutomationFailure('unsupported-content', 'the fill is not solid', 'transparency');
    xml = `<a:solidFill>${colorSpec(color, 'color', fill.transparency)}</a:solidFill>`;
  } else badArgument('fill');
  edit(() => setCellFillXml(ctx.batch, m, block, xml));
  return { kind: 'applied' };
}
