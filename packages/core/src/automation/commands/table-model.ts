// Table structure on the canonical tree (ECMA-376 Part 1, 21.1.3 "Table"): the grid of `a:tc`
// cells with their merges, and the structure edits (rows, columns, merge, split). Every edit
// compiles to ops on one batch and keeps the table valid: one `a:gridCol` per grid column, one
// `a:tc` per grid column in every `a:tr`, and `gridSpan`/`rowSpan`/`hMerge`/`vMerge` that agree.
//
// Merges follow 21.1.3.16 tc: every cell of a merged block's first row
// carries the block's `rowSpan`, and every cell of its first column the block's `gridSpan`; the
// other cells of the first row carry `hMerge="1"`, the other cells of the first column
// `vMerge="1"`, and the rest both markers. Reference files save a 2 × 2 merge so.

import {
  addedTableCell,
  attr,
  cellAsSaved,
  child,
  children,
  elements,
  intAttr,
  joinedTextBodies,
  NS,
  path,
  tableCellCounts,
  tableGridOf,
  TABLE_MAX_COLUMNS,
  TABLE_MAX_ROWS,
  TABLE_REPAIR_MAX_CELLS,
} from '../../store/index.ts';
import { specOf, type Element, type NodeSpec } from '../../store/index.ts';
import type { TableCellAddress } from '../../contracts/neutral.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { paragraphText, paragraphsOf } from './text-model.ts';

/**
 * The most columns the editor makes (the Insert Table limit), and the most cells in a row of a
 * table that the repair rewrites.
 */
export const MAX_TABLE_SIZE = TABLE_REPAIR_MAX_CELLS;
/** The most rows the editor grows a table to. */
export const MAX_ROWS = TABLE_MAX_ROWS;

/** A rectangle of grid positions, inclusive. */
export interface CellRange {
  readonly r0: number;
  readonly r1: number;
  readonly c0: number;
  readonly c1: number;
}

/** One table read from the tree. */
export interface TableModel {
  readonly part: string;
  readonly frame: Element;
  readonly tbl: Element;
  readonly gridCols: readonly Element[];
  readonly rows: readonly Element[];
  /** `tcs[r][c]`: the `a:tc` at grid position (r, c). */
  readonly tcs: readonly (readonly Element[])[];
  /** The top-left cell of the merged block that covers each grid position. */
  readonly anchors: readonly (readonly TableCellAddress[])[];
  readonly nR: number;
  readonly nC: number;
}

const A = NS.a;

/** The `a:tbl` of a graphic frame, or undefined. */
export function tblOf(frame: Element | undefined): Element | undefined {
  return path(frame, 'a:graphic', 'a:graphicData', 'a:tbl');
}

/**
 * Read the table of a graphic frame. Refuses an element that is not a table, a table whose
 * rows do not hold one cell per grid column (`errors.unsupported`, `irregular table`): editing
 * one would guess at what the file means, and a table larger than the cascade draws
 * (`disabled.tableTooLarge`).
 */
export function readTable(batch: OpBatch, part: string, frameId: string): TableModel {
  const frame = batch.element(part, frameId);
  const tbl = tblOf(frame);
  if (frame.kind !== 'graphicFrame' || !tbl) throw new CommandRefusal('disabled.noTable');
  const gridCols = children(child(tbl, 'a:tblGrid'), 'a:gridCol');
  const rows = children(tbl, 'a:tr');
  const nC = gridCols.length;
  const nR = rows.length;
  const tcs = rows.map((tr) => children(tr, 'a:tc'));
  if (nC === 0 || nR === 0 || tcs.some((row) => row.length !== nC))
    throw new CommandRefusal('errors.unsupported', 'irregular table');
  // The cascade draws at most this many rows and columns; past them its merge grid and this one
  // could disagree, so a command would edit cells other than the ones the slide shows.
  if (nR > TABLE_MAX_ROWS || nC > TABLE_MAX_COLUMNS)
    throw new CommandRefusal('disabled.tableTooLarge');
  // The cascade's merge grid: a cell with an orphan `hMerge` or `vMerge` (no span reaches it) is
  // a drawn cell with its own spans, and a span that runs into an earlier block ends there.
  const grid = tableGridOf(frame, tcs, nC);
  const anchors: TableCellAddress[][] = rows.map((_, r) =>
    Array.from({ length: nC }, (_, c) => {
      const b = grid.blocks[grid.owner[r * nC + c]];
      return { row: b.row, column: b.col };
    })
  );
  return { part, frame, tbl, gridCols, rows, tcs, anchors, nR, nC };
}

const same = (a: TableCellAddress, b: TableCellAddress): boolean =>
  a.row === b.row && a.column === b.column;

/** The grid positions the merged block over a position covers. */
export function blockOf(m: TableModel, at: TableCellAddress): CellRange {
  const a = m.anchors[at.row]?.[at.column];
  if (!a) throw new CommandRefusal('errors.unsupported', 'cell');
  let r1 = a.row;
  let c1 = a.column;
  while (r1 + 1 < m.nR && same(m.anchors[r1 + 1][a.column], a)) r1 += 1;
  while (c1 + 1 < m.nC && same(m.anchors[a.row][c1 + 1], a)) c1 += 1;
  return { r0: a.row, r1, c0: a.column, c1 };
}

/**
 * The cells a cell range selects: a whole row or column band (every row, or every column) is
 * its own grid lines, as Select Row and Select Column name them; any other range grows over the
 * merged cells that cross its edge (`expandRange`). The canvas highlight reads the same rule
 * (`selectedLayoutRange`), so what shows selected is what a command changes.
 */
export function selectedRange(m: TableModel, range: CellRange): CellRange {
  return isBand(range, m.nR, m.nC) ? range : expandRange(m, range);
}

/** Whether a range covers every row or every column of a grid of `rows` × `columns`. */
export function isBand(range: CellRange, rows: number, columns: number): boolean {
  return (range.r0 === 0 && range.r1 === rows - 1) || (range.c0 === 0 && range.c1 === columns - 1);
}

/** A range grown until no merged block crosses its edge. */
export function expandRange(m: TableModel, range: CellRange): CellRange {
  let r0 = Math.max(0, Math.min(range.r0, m.nR - 1));
  let r1 = Math.max(r0, Math.min(range.r1, m.nR - 1));
  let c0 = Math.max(0, Math.min(range.c0, m.nC - 1));
  let c1 = Math.max(c0, Math.min(range.c1, m.nC - 1));
  for (let changed = true; changed;) {
    changed = false;
    for (let r = r0; r <= r1; r += 1)
      for (let c = c0; c <= c1; c += 1) {
        const b = blockOf(m, { row: r, column: c });
        if (b.r0 < r0 || b.r1 > r1 || b.c0 < c0 || b.c1 > c1) {
          r0 = Math.min(r0, b.r0);
          r1 = Math.max(r1, b.r1);
          c0 = Math.min(c0, b.c0);
          c1 = Math.max(c1, b.c1);
          changed = true;
        }
      }
  }
  return { r0, r1, c0, c1 };
}

/** The top-left cells of the merged blocks inside a range, row by row. */
export function anchorsIn(m: TableModel, range: CellRange): TableCellAddress[] {
  const out: TableCellAddress[] = [];
  for (let r = range.r0; r <= range.r1; r += 1)
    for (let c = range.c0; c <= range.c1; c += 1) {
      const a = m.anchors[r][c];
      if (a.row === r && a.column === c) out.push(a);
    }
  return out;
}

/** The whole grid as a range. */
export const wholeTable = (m: TableModel): CellRange => ({
  r0: 0,
  r1: m.nR - 1,
  c0: 0,
  c1: m.nC - 1,
});

/** The `a:txBody` of the cell at an address (its block's top-left cell). */
export function cellBody(m: TableModel, at: TableCellAddress): Element {
  const a = m.anchors[at.row]?.[at.column];
  const body = a ? child(m.tcs[a.row][a.column], 'a:txBody') : undefined;
  if (!body) throw new CommandRefusal('disabled.noText');
  return body;
}

export const setAttr = (
  batch: OpBatch,
  part: string,
  node: string,
  name: string,
  value: string | null
): void => batch.push({ op: 'setAttribute', part, node, name: { ns: '', local: name }, value });

export const aSpec = (
  local: string,
  attrs: Record<string, string> = {},
  kids: NodeSpec['children'] = []
): NodeSpec => ({
  ns: A,
  local,
  prefix: 'a',
  attrs: Object.entries(attrs).map(([k, v]) => ({ ns: '', local: k, value: v })),
  children: kids,
});

/**
 * The merge attributes of grid position (r, c) inside block `b`: the height on the first row and
 * the width on the first column, `hMerge` off the first column and `vMerge` off the first row.
 */
export function blockAttrs(b: CellRange, r: number, c: number): Record<string, string> {
  const g = b.c1 - b.c0 + 1;
  const k = b.r1 - b.r0 + 1;
  const out: Record<string, string> = {};
  if (c === b.c0 && g > 1) out.gridSpan = String(g);
  if (r === b.r0 && k > 1) out.rowSpan = String(k);
  if (c !== b.c0) out.hMerge = '1';
  if (r !== b.r0) out.vMerge = '1';
  return out;
}

const MERGE_ATTRS = ['gridSpan', 'rowSpan', 'hMerge', 'vMerge'] as const;

/**
 * Set the merge attributes of the cell at (r, c) to exactly `want`. A covered cell may have no
 * text body (`a:txBody` is optional, 21.1.3.16); one that becomes a drawn cell gets an empty
 * body, formatted like its row's first cell, so it takes text. A drawn cell without a body stays
 * as it is: typing gives it one (`selectionBody`), and in a collaboration room an undo of this
 * write would take a body that a peer typed into at the same time.
 */
function writeCell(
  batch: OpBatch,
  m: TableModel,
  r: number,
  c: number,
  want: Record<string, string>
): void {
  const tc = m.tcs[r][c];
  const on = (name: string) => ['1', 'true'].includes(attr(tc, name) ?? '');
  const covered = on('hMerge') || on('vMerge');
  for (const name of MERGE_ATTRS)
    if (attr(tc, name) !== want[name]) setAttr(batch, m.part, tc.id, name, want[name] ?? null);
  if (covered && !want.hMerge && !want.vMerge && !child(batch.element(m.part, tc.id), 'a:txBody'))
    batch.insert({
      op: 'insertNode',
      part: m.part,
      parent: tc.id,
      index: 0,
      spec: emptyBodySpec(m.tcs[r][0]),
    });
}

/** Rewrite the merge attributes of every cell of a block. */
export function writeBlock(batch: OpBatch, m: TableModel, b: CellRange): void {
  for (let r = b.r0; r <= b.r1; r += 1)
    for (let c = b.c0; c <= b.c1; c += 1) writeCell(batch, m, r, c, blockAttrs(b, r, c));
}

/**
 * The text body of a new, empty cell, formatted like `source`: the source's body and list
 * properties, and its first paragraph's properties and character format as the end mark
 * (`addedTableCell`).
 */
export function emptyBodySpec(source: Element | undefined): NodeSpec {
  return specOf(child(addedTableCell(source), 'a:txBody')!);
}

/**
 * A new cell: an empty body formatted like `source`, the source's cell properties, and `attrs`.
 * The cell is `addedTableCell`, which the cascade also draws where the save adds a cell.
 */
export function newCellSpec(source: Element | undefined, attrs: Record<string, string>): NodeSpec {
  const spec = specOf(addedTableCell(source));
  return {
    ...spec,
    attrs: Object.entries(attrs).map(([k, v]) => ({ ns: '', local: k, value: v })),
  };
}

/**
 * Join the text bodies of a cell into `keep`, its first body: a cell holds one `a:txBody`
 * (21.1.3.16 tc), and no text goes. The paragraphs of each body that `joinedTextBodies` names
 * end up in `keep`, in body order; the other paragraphs and `extras` go. The save and an edit of
 * the cell join the bodies so, and the canvas draws the joined body before (`cellAsSaved`).
 */
export function joinTextBodies(
  batch: OpBatch,
  part: string,
  keep: Element,
  extras: readonly Element[]
): void {
  const join = new Set(joinedTextBodies([keep, ...extras]));
  for (const extra of extras) {
    if (join.has(extra))
      for (const p of paragraphsOf(extra))
        batch.push({
          op: 'moveNode',
          part,
          node: p.id,
          parent: keep.id,
          index: batch.element(part, keep.id).children.length,
        });
    batch.push({ op: 'removeNode', part, node: extra.id });
  }
  // A blank first body gives way to the bodies that hold content.
  if (!join.has(keep))
    for (const p of paragraphsOf(keep)) batch.push({ op: 'removeNode', part, node: p.id });
}

/**
 * Join the text bodies of each cell with more than one in the table of frame `frameId`
 * (`joinTextBodies`), in every table that the save joins them in (`cellAsSaved`). A file, or two
 * peers who type at once into a cell with no text body, can leave such a cell. An edit that
 * reads the first body only would leave the text of the others behind, so an edit joins first.
 */
export function joinCellBodies(batch: OpBatch, part: string, frameId: string): void {
  const tbl = tblOf(batch.element(part, frameId));
  if (!tableCellCounts(tbl)) return;
  for (const tr of children(tbl, 'a:tr'))
    for (const tc of children(tr, 'a:tc')) {
      const [keep, ...rest] = children(tc, 'a:txBody');
      if (keep && rest.length > 0) joinTextBodies(batch, part, keep, rest);
    }
}

/**
 * The text body of cell `tc` of table `tbl` as a read sees it: its bodies joined where the save
 * joins them (`cellAsSaved`). Read only; an edit joins them first (`joinCellBodies`).
 */
export function cellBodyAsSaved(tbl: Element | undefined, tc: Element): Element | undefined {
  return child(tableCellCounts(tbl) ? cellAsSaved(tc) : tc, 'a:txBody');
}

/** Index in `parent` just after its last child named one of `locals`. */
function afterLast(parent: Element, locals: readonly string[]): number {
  let at = 0;
  parent.children.forEach((n, i) => {
    if (n.type === 'element' && locals.includes(n.local)) at = i + 1;
  });
  return at;
}

/**
 * Insert `count` rows at row index `at` (0 to the row count), formatted like row `source`. A
 * merged block that crosses the insertion line grows over the new rows; elsewhere the new rows
 * repeat the source row's horizontal merges, or hold one cell per column when `plain` (a paste
 * fills every new cell).
 */
export function insertRows(
  batch: OpBatch,
  m: TableModel,
  at: number,
  count: number,
  source: number,
  plain = false
): void {
  if (m.nR + count > MAX_ROWS) throw new CommandRefusal('disabled.tableTooLarge');
  const grown = new Set<string>();
  const rows: NodeSpec[][] = Array.from({ length: count }, () => []);
  for (let c = 0; c < m.nC; c += 1) {
    const crossing = at > 0 && at < m.nR && same(m.anchors[at - 1][c], m.anchors[at][c]);
    if (crossing) {
      const b = blockOf(m, m.anchors[at][c]);
      const key = `${b.r0}:${b.c0}`;
      if (!grown.has(key)) {
        grown.add(key);
        // Every column of the lengthened block carries the height in its first row.
        const k = b.r1 - b.r0 + 1 + count;
        for (let cc = b.c0; cc <= b.c1; cc += 1)
          setAttr(batch, m.part, m.tcs[b.r0][cc].id, 'rowSpan', String(k));
      }
      const grownBlock = { ...b, r1: b.r1 + count };
      for (let i = 0; i < count; i += 1)
        rows[i].push(newCellSpec(m.tcs[at][c], blockAttrs(grownBlock, at + i, c)));
    } else {
      const b = blockOf(m, m.anchors[source][c]);
      const line = plain ? { r0: 0, r1: 0, c0: c, c1: c } : { r0: 0, r1: 0, c0: b.c0, c1: b.c1 };
      // A covered cell holds no fill of its own: the new cell takes its block's anchor's look.
      for (let i = 0; i < count; i += 1)
        rows[i].push(newCellSpec(m.tcs[b.r0][b.c0], blockAttrs(line, 0, c)));
    }
  }
  const before = m.rows[at];
  const index = before
    ? m.tbl.children.findIndex((n) => n.id === before.id)
    : afterLast(m.tbl, ['tblGrid', 'tr']);
  const h = String(intAttr(m.rows[source], 'h') ?? 370_840);
  rows.forEach((cells, i) =>
    batch.insert({
      op: 'insertNode',
      part: m.part,
      parent: m.tbl.id,
      index: index + i,
      spec: aSpec('tr', { h }, cells),
    })
  );
}

/**
 * Insert `count` columns at grid index `at`, as wide as column `source` and formatted like it.
 * A merged block that crosses the insertion line grows over the new columns; elsewhere the new
 * columns repeat the source column's vertical merges, or hold one cell per row when `plain` (a
 * paste fills every new cell).
 */
export function insertColumns(
  batch: OpBatch,
  m: TableModel,
  at: number,
  count: number,
  source: number,
  plain = false
): void {
  if (m.nC + count > MAX_TABLE_SIZE) throw new CommandRefusal('disabled.tableTooLarge');
  const w = String(intAttr(m.gridCols[source], 'w') ?? 0);
  const grid = child(m.tbl, 'a:tblGrid')!;
  const gridIndex = m.gridCols[at]
    ? grid.children.findIndex((n) => n.id === m.gridCols[at].id)
    : afterLast(grid, ['gridCol']);
  for (let i = 0; i < count; i += 1)
    batch.insert({
      op: 'insertNode',
      part: m.part,
      parent: grid.id,
      index: gridIndex + i,
      spec: aSpec('gridCol', { w }),
    });
  const grown = new Set<string>();
  m.rows.forEach((tr, r) => {
    const crossing = at > 0 && at < m.nC && same(m.anchors[r][at - 1], m.anchors[r][at]);
    const specs: NodeSpec[] = [];
    if (crossing) {
      const b = blockOf(m, m.anchors[r][at]);
      const key = `${b.r0}:${b.c0}`;
      if (!grown.has(key)) {
        grown.add(key);
        // Every row of the widened block carries the width in its first column.
        for (let rr = b.r0; rr <= b.r1; rr += 1)
          setAttr(batch, m.part, m.tcs[rr][b.c0].id, 'gridSpan', String(b.c1 - b.c0 + 1 + count));
      }
      const grownBlock = { ...b, c1: b.c1 + count };
      for (let i = 0; i < count; i += 1)
        specs.push(newCellSpec(m.tcs[r][at], blockAttrs(grownBlock, r, at + i)));
    } else {
      const b = blockOf(m, m.anchors[r][source]);
      const line = plain ? { r0: r, r1: r, c0: 0, c1: 0 } : { r0: b.r0, r1: b.r1, c0: 0, c1: 0 };
      // A covered cell holds no fill of its own: the new cell takes its block's anchor's look.
      for (let i = 0; i < count; i += 1)
        specs.push(newCellSpec(m.tcs[b.r0][b.c0], blockAttrs(line, r, 0)));
    }
    const next = m.tcs[r][at];
    const index = next ? tr.children.findIndex((n) => n.id === next.id) : afterLast(tr, ['tc']);
    specs.forEach((spec, i) =>
      batch.insert({ op: 'insertNode', part: m.part, parent: tr.id, index: index + i, spec })
    );
  });
}

/**
 * Make the covered cell at `to` the top-left cell of what remains of a block whose top-left
 * cell `from` is deleted: it takes the block's text and cell properties. The text body moves
 * with its node ids, so a peer who types in the merged cell at the same time keeps the typing.
 */
function moveAnchor(batch: OpBatch, m: TableModel, from: Element, to: Element): void {
  const body = child(from, 'a:txBody');
  if (body) {
    const old = child(batch.element(m.part, to.id), 'a:txBody');
    if (old) batch.push({ op: 'removeNode', part: m.part, node: old.id });
    // CT_TableCell order: `a:txBody`, `a:tcPr`, `a:extLst`.
    batch.push({ op: 'moveNode', part: m.part, node: body.id, parent: to.id, index: 0 });
  }
  const src = child(from, 'a:tcPr');
  if (!src) return;
  const now = batch.element(m.part, to.id);
  const old = child(now, 'a:tcPr');
  const index = old ? now.children.findIndex((n) => n.id === old.id) : afterLast(now, ['txBody']);
  if (old) batch.push({ op: 'removeNode', part: m.part, node: old.id });
  batch.insert({ op: 'insertNode', part: m.part, parent: to.id, index, spec: specOf(src) });
}

/** Delete rows `r0`–`r1`. Merged blocks that reach past them shrink and keep their text. */
export function deleteRows(batch: OpBatch, m: TableModel, r0: number, r1: number): void {
  for (const a of anchorsIn(m, wholeTable(m))) {
    const b = blockOf(m, a);
    if (b.r1 < r0 || b.r0 > r1 || (b.r0 >= r0 && b.r1 <= r1)) continue;
    const rows: number[] = [];
    for (let r = b.r0; r <= b.r1; r += 1) if (r < r0 || r > r1) rows.push(r);
    if (b.r0 >= r0) moveAnchor(batch, m, m.tcs[b.r0][b.c0], m.tcs[rows[0]][b.c0]);
    // The rows of the block that stay, renumbered from 0.
    const kept = { r0: 0, r1: rows.length - 1, c0: b.c0, c1: b.c1 };
    rows.forEach((r, i) => {
      for (let c = b.c0; c <= b.c1; c += 1) writeCell(batch, m, r, c, blockAttrs(kept, i, c));
    });
  }
  for (let r = r0; r <= r1; r += 1)
    batch.push({ op: 'removeNode', part: m.part, node: m.rows[r].id });
}

/** Delete columns `c0`–`c1`. Merged blocks that reach past them shrink and keep their text. */
export function deleteColumns(batch: OpBatch, m: TableModel, c0: number, c1: number): void {
  for (const a of anchorsIn(m, wholeTable(m))) {
    const b = blockOf(m, a);
    if (b.c1 < c0 || b.c0 > c1 || (b.c0 >= c0 && b.c1 <= c1)) continue;
    const cols: number[] = [];
    for (let c = b.c0; c <= b.c1; c += 1) if (c < c0 || c > c1) cols.push(c);
    if (b.c0 >= c0) moveAnchor(batch, m, m.tcs[b.r0][b.c0], m.tcs[b.r0][cols[0]]);
    // The columns of the block that stay, renumbered from 0.
    const kept = { r0: b.r0, r1: b.r1, c0: 0, c1: cols.length - 1 };
    for (let r = b.r0; r <= b.r1; r += 1)
      cols.forEach((c, i) => writeCell(batch, m, r, c, blockAttrs(kept, r, i)));
  }
  for (let c = c0; c <= c1; c += 1) {
    batch.push({ op: 'removeNode', part: m.part, node: m.gridCols[c].id });
    for (let r = 0; r < m.nR; r += 1)
      batch.push({ op: 'removeNode', part: m.part, node: m.tcs[r][c].id });
  }
}

/** Whether a cell body holds no text. */
function emptyBody(body: Element | undefined): boolean {
  return !body || paragraphsOf(body).every((p) => paragraphText(p).length === 0);
}

/**
 * Merge a cell range into one cell. The text of every merged cell that holds text moves into
 * the top-left cell, in reading order, and each cell it left keeps one empty paragraph (a text
 * body holds at least one, 21.1.2.1.1). The moved paragraphs keep their ids, so a peer who types
 * in one at the same time types in the merged cell. A merged cell without text gets a new copy
 * of its empty paragraph, so text that a peer types into the old one at the same time is a
 * peer's text in a paragraph the merge removed. In a collaboration room the room's read shows
 * such text, and text a concurrent structure change leaves in a covered cell, in the merged cell
 * (`@pptx-editor.dev/pro`, `merged-cells.ts`).
 */
export function mergeRange(batch: OpBatch, m: TableModel, range: CellRange): void {
  const blocks = anchorsIn(m, range);
  if (blocks.length < 2) throw new CommandRefusal('disabled.mergeOneCell');
  const target = child(m.tcs[range.r0][range.c0], 'a:txBody');
  if (!target) throw new CommandRefusal('disabled.noText');
  let targetEmpty = emptyBody(target);
  for (const a of blocks.slice(1)) {
    const body = child(m.tcs[a.row][a.column], 'a:txBody');
    if (!body) continue;
    const ps = paragraphsOf(body);
    const end = batch.element(m.part, body.id).children.length;
    if (emptyBody(body)) {
      for (const p of ps) batch.push({ op: 'removeNode', part: m.part, node: p.id });
      batch.insert({
        op: 'insertNode',
        part: m.part,
        parent: body.id,
        index: end - ps.length,
        spec: ps[0] ? specOf(ps[0]) : emptyParagraphSpec(),
      });
      continue;
    }
    if (targetEmpty) {
      for (const p of paragraphsOf(batch.element(m.part, target.id)))
        batch.push({ op: 'removeNode', part: m.part, node: p.id });
      targetEmpty = false;
    }
    for (const p of ps)
      batch.push({
        op: 'moveNode',
        part: m.part,
        node: p.id,
        parent: target.id,
        index: batch.element(m.part, target.id).children.length,
      });
    batch.insert({
      op: 'insertNode',
      part: m.part,
      parent: body.id,
      index: batch.element(m.part, body.id).children.length,
      spec: emptyParagraphSpec(),
    });
  }
  writeBlock(batch, m, range);
}

/** An empty paragraph. */
const emptyParagraphSpec = (): NodeSpec =>
  aSpec('p', {}, [aSpec('endParaRPr', { lang: 'en-US', dirty: '0' })]);

/**
 * Split the cell at `at` into `cols` columns. A merged cell splits along its grid lines when
 * its width divides evenly; a cell one grid column wide gets new, equal grid columns, and the
 * cells above and below it widen over them. Returns the columns the cell covers afterwards.
 */
export function splitColumns(
  batch: OpBatch,
  m: TableModel,
  at: TableCellAddress,
  cols: number
): CellRange {
  const b = blockOf(m, at);
  const g = b.c1 - b.c0 + 1;
  if (cols <= 1) return b;
  if (g % cols === 0) {
    const w = g / cols;
    for (let i = 0; i < cols; i += 1)
      writeBlock(batch, m, { ...b, c0: b.c0 + i * w, c1: b.c0 + (i + 1) * w - 1 });
    return b;
  }
  if (g !== 1) throw new CommandRefusal('disabled.splitUneven');
  if (m.nC + cols - 1 > MAX_TABLE_SIZE) throw new CommandRefusal('disabled.tableTooLarge');
  // The new grid columns share the old column's width.
  const total = intAttr(m.gridCols[b.c0], 'w') ?? 0;
  const each = Math.floor(total / cols);
  setAttr(batch, m.part, m.gridCols[b.c0].id, 'w', String(total - each * (cols - 1)));
  const grid = child(m.tbl, 'a:tblGrid')!;
  const gi = grid.children.findIndex((n) => n.id === m.gridCols[b.c0].id) + 1;
  for (let i = 0; i < cols - 1; i += 1)
    batch.insert({
      op: 'insertNode',
      part: m.part,
      parent: grid.id,
      index: gi + i,
      spec: aSpec('gridCol', { w: String(each) }),
    });
  const grown = new Set<string>();
  m.rows.forEach((tr, r) => {
    const tc = m.tcs[r][b.c0];
    const specs: NodeSpec[] = [];
    if (r >= b.r0 && r <= b.r1) {
      // New cells of the split block: each one column wide, as tall as the block.
      const line = { r0: b.r0, r1: b.r1, c0: 0, c1: 0 };
      for (let i = 0; i < cols - 1; i += 1) specs.push(newCellSpec(tc, blockAttrs(line, r, 0)));
    } else {
      const ob = blockOf(m, m.anchors[r][b.c0]);
      const key = `${ob.r0}:${ob.c0}`;
      if (!grown.has(key)) {
        grown.add(key);
        for (let rr = ob.r0; rr <= ob.r1; rr += 1)
          setAttr(batch, m.part, m.tcs[rr][ob.c0].id, 'gridSpan', String(ob.c1 - ob.c0 + cols));
      }
      const wide = { ...ob, c1: ob.c1 + cols - 1 };
      for (let i = 0; i < cols - 1; i += 1)
        specs.push(newCellSpec(tc, blockAttrs(wide, r, b.c0 + 1 + i)));
    }
    const index = tr.children.findIndex((n) => n.id === tc.id) + 1;
    specs.forEach((spec, i) =>
      batch.insert({ op: 'insertNode', part: m.part, parent: tr.id, index: index + i, spec })
    );
  });
  return { ...b, c1: b.c0 + cols - 1 };
}

/**
 * Split the cells of `area` (one block's rows, after `splitColumns`) into `rows` rows, the same
 * way: along grid lines when the height divides evenly, else into new rows that share the old
 * row's height while the cells beside it lengthen over them.
 */
export function splitRows(batch: OpBatch, m: TableModel, area: CellRange, rows: number): void {
  if (rows <= 1) return;
  const k = area.r1 - area.r0 + 1;
  const parts = anchorsIn(m, { ...area, r1: area.r0 }).map((a) => blockOf(m, a));
  if (k % rows === 0) {
    const h = k / rows;
    for (const cb of parts)
      for (let i = 0; i < rows; i += 1)
        writeBlock(batch, m, { ...cb, r0: area.r0 + i * h, r1: area.r0 + (i + 1) * h - 1 });
    return;
  }
  if (k !== 1) throw new CommandRefusal('disabled.splitUneven');
  if (m.nR + rows - 1 > MAX_ROWS) throw new CommandRefusal('disabled.tableTooLarge');
  const tr = m.rows[area.r0];
  const total = intAttr(tr, 'h') ?? 0;
  const each = Math.floor(total / rows);
  setAttr(batch, m.part, tr.id, 'h', String(total - each * (rows - 1)));
  const grown = new Set<string>();
  const newRows: NodeSpec[][] = Array.from({ length: rows - 1 }, () => []);
  for (let c = 0; c < m.nC; c += 1) {
    const tc = m.tcs[area.r0][c];
    if (c >= area.c0 && c <= area.c1) {
      const cb = blockOf(m, m.anchors[area.r0][c]);
      const line = { r0: 0, r1: 0, c0: cb.c0, c1: cb.c1 };
      for (const cells of newRows) cells.push(newCellSpec(tc, blockAttrs(line, 0, c)));
      continue;
    }
    const ob = blockOf(m, m.anchors[area.r0][c]);
    const key = `${ob.r0}:${ob.c0}`;
    if (!grown.has(key)) {
      grown.add(key);
      for (let cc = ob.c0; cc <= ob.c1; cc += 1)
        setAttr(batch, m.part, m.tcs[ob.r0][cc].id, 'rowSpan', String(ob.r1 - ob.r0 + rows));
    }
    const tall = { ...ob, r1: ob.r1 + rows - 1 };
    newRows.forEach((cells, i) =>
      cells.push(newCellSpec(tc, blockAttrs(tall, area.r0 + 1 + i, c)))
    );
  }
  const index = m.tbl.children.findIndex((n) => n.id === tr.id) + 1;
  newRows.forEach((cells, i) =>
    batch.insert({
      op: 'insertNode',
      part: m.part,
      parent: m.tbl.id,
      index: index + i,
      spec: aSpec('tr', { h: String(each) }, cells),
    })
  );
}

/** Remove the text of one cell, keeping one paragraph with its format. */
function clearCell(batch: OpBatch, part: string, tc: Element): void {
  const body = child(tc, 'a:txBody');
  if (!body || emptyBody(body)) return;
  const ps = paragraphsOf(body);
  for (const p of ps.slice(1)) batch.push({ op: 'removeNode', part, node: p.id });
  for (const e of elements(ps[0]))
    if (e.kind === 'run' || e.kind === 'break' || e.kind === 'field')
      batch.push({ op: 'removeNode', part, node: e.id });
}

/** Remove the text of the cells in a range, keeping one paragraph with its format. */
export function clearRange(batch: OpBatch, m: TableModel, range: CellRange): void {
  for (const a of anchorsIn(m, range)) clearCell(batch, m.part, m.tcs[a.row][a.column]);
}

const textRun = (line: string, rPr: Element | undefined): NodeSpec => ({
  ns: NS.a,
  local: 'r',
  prefix: 'a',
  children: [
    ...(rPr ? [specOf(rPr)] : []),
    { ns: NS.a, local: 't', prefix: 'a', children: [{ text: line }] },
  ],
});

/**
 * Replace the text of the cell at grid position `at` (a drawn cell's top-left position) with
 * plain `text`: one paragraph for each line, each with the first paragraph's properties, and
 * every run in the character format of the cell's first run.
 */
export function setCellText(
  batch: OpBatch,
  m: TableModel,
  at: TableCellAddress,
  text: string
): void {
  const tc = m.tcs[at.row][at.column];
  if (!child(tc, 'a:txBody'))
    batch.insert({
      op: 'insertNode',
      part: m.part,
      parent: tc.id,
      index: 0,
      spec: emptyBodySpec(tc),
    });
  const body = child(batch.element(m.part, tc.id), 'a:txBody')!;
  const [first] = paragraphsOf(body);
  const run = first ? elements(first).find((e) => e.kind === 'run') : undefined;
  const rPr = run ? child(run, 'a:rPr') : undefined;
  const pPr = first ? child(first, 'a:pPr') : undefined;
  const endPr = first ? child(first, 'a:endParaRPr') : undefined;
  clearCell(batch, m.part, batch.element(m.part, tc.id));
  // An empty cell can still hold empty paragraphs (`clearCell` leaves them): one stays.
  for (const p of paragraphsOf(child(batch.element(m.part, tc.id), 'a:txBody')!).slice(1))
    batch.push({ op: 'removeNode', part: m.part, node: p.id });
  const lines = text.split(/\r\n|\r|\n/);
  if (!first) return;
  if (lines[0])
    batch.insert({
      op: 'insertNode',
      part: m.part,
      parent: first.id,
      index: pPr ? 1 : 0,
      spec: textRun(lines[0], rPr),
    });
  const at0 = body.children.indexOf(first);
  lines.slice(1).forEach((line, i) => {
    const spec: NodeSpec = {
      ns: NS.a,
      local: 'p',
      prefix: 'a',
      children: [
        ...(pPr ? [specOf(pPr)] : []),
        ...(line ? [textRun(line, rPr)] : []),
        ...(endPr ? [specOf(endPr)] : []),
      ],
    };
    batch.insert({ op: 'insertNode', part: m.part, parent: body.id, index: at0 + 1 + i, spec });
  });
}
