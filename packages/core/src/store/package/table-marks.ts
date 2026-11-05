// The structure mark of a table frame, and the merge grid the editor reads for a frame. A
// structure command (insert or delete rows and columns, merge, split) in a collaboration room
// adds a mark to the table it changes (`automation/commands/table-edit-marks.ts` writes and removes marks).
// Peers who change one table's structure at the same time can leave merge markup that fails the
// grid: a merge and a concurrent delete of one of its covered rows leave the merged cell's span
// over the row below, and a row insert and a concurrent column insert leave a row with one cell
// too few. The save writes every table with its cell counts repaired (`tableCellCounts`, new
// cells as `addedTableCell`), one text body in each cell (`cellAsSaved`), and a marked table
// with the merge markers of `editedTableGrid`. The cascade and find read a table by
// `tableAsDrawn`: a marked table with the same cell counts and the same grid (`tableGridOf`),
// so the canvas draws the grid the save writes, with no store write, and every table with the
// cell text the save writes. A table without a mark keeps the grid and the cell counts of its
// file's markup. Concurrent deletes of the last
// columns or rows leave a marked table with no grid column or no row: the save leaves that frame
// out, the cascade does not draw it, and the selection pane does not list it (`isEmptiedTable`).
//
// A mark is an `a:ext` element (ECMA-376 Part 1, 20.1.2.2.15 ext) with the editor's `uri`, a
// direct child of the frame's `p:cNvPr`. It lives only in the room: open, copy, paste, and save
// remove every mark. A mark in a room is attacker-controlled like the rest of the room; a forged
// one makes a table read and save by `editedTableGrid`, which only ever writes valid markup.

import { NS } from './namespaces.ts';
import { deriveKind } from './node-kinds.ts';
import { attr, child, children, elements, intAttr } from './query.ts';
import {
  editedTableGrid,
  resolveTableGrid,
  TABLE_MAX_ROWS,
  tableCellHasText,
  type TableGrid,
} from './table-grid.ts';
import type { Element, Node, NodeKind } from './tree.ts';

/** The `uri` of the extension that marks a table whose structure a room changed. @internal */
export const TABLE_EDIT_URI = 'urn:pptx-editor.dev:table-structure-edit';

/** The `p:cNvPr` of a graphic frame. */
const cNvPrOf = (frame: Element | undefined): Element | undefined =>
  child(child(frame, 'p:nvGraphicFramePr'), 'p:cNvPr');

/** Whether an element is a structure mark. @internal */
export const isStructureMark = (el: Element): boolean =>
  el.ns === NS.a && el.local === 'ext' && attr(el, 'uri') === TABLE_EDIT_URI;

/** The structure marks of a graphic frame, in document order. @internal */
export function structureMarksOf(frame: Element | undefined): Element[] {
  return elements(cNvPrOf(frame)).filter(isStructureMark);
}

/** Whether a structure command in a collaboration room changed the table of `frame`. @internal */
export function isStructureEdited(frame: Element | undefined): boolean {
  return elements(cNvPrOf(frame)).some(isStructureMark);
}

/**
 * Whether `frame` holds a table with a structure mark and no grid column or no row. Peers who
 * delete the last columns (or rows) of a table at the same time leave such a table, and a row
 * that one of them inserted can keep a cell of a column that the other deleted. A local delete
 * of the last column or row deletes the whole table, so the save leaves this frame out, and the
 * cascade does not draw it. A table without a mark keeps what its file wrote. @internal
 */
export function isEmptiedTable(frame: Element | undefined): boolean {
  const tbl = child(child(child(frame, 'a:graphic'), 'a:graphicData'), 'a:tbl');
  if (!tbl || !isStructureEdited(frame)) return false;
  const has = (el: Element | undefined, local: string) =>
    elements(el).some((c) => c.ns === NS.a && c.local === local);
  return !has(child(tbl, 'a:tblGrid'), 'gridCol') || !has(tbl, 'tr');
}

/**
 * The most cells in a row, and rows, of a table that the table repair rewrites. A larger table
 * comes from a file, which would decide the work a repair does, so the repair and the save
 * leave its markup as it is. @internal
 */
export const TABLE_REPAIR_MAX_CELLS = 75;

/** Whether the save repairs the merge markers of the table of a marked `frame`. */
function repairable(frame: Element | undefined): boolean {
  const tbl = child(child(child(frame, 'a:graphic'), 'a:graphicData'), 'a:tbl');
  const rows = elements(tbl).filter((r) => r.ns === NS.a && r.local === 'tr');
  if (rows.length > TABLE_MAX_ROWS) return false;
  return rows.every(
    (tr) =>
      elements(tr).filter((c) => c.ns === NS.a && c.local === 'tc').length <= TABLE_REPAIR_MAX_CELLS
  );
}

/**
 * The merge grid of the table of `frame`: `editedTableGrid` when the frame has a structure
 * mark and the save repairs the table, else `resolveTableGrid`. `cells` and `columns` follow
 * `resolveTableGrid`. The cascade, the table commands, find, and the save read a table by this
 * one function. @internal
 */
export function tableGridOf(
  frame: Element | undefined,
  cells: readonly (readonly (Element | undefined)[])[],
  columns: number
): TableGrid {
  return isStructureEdited(frame) && repairable(frame)
    ? editedTableGrid(cells, columns, tableCellHasText)
    : resolveTableGrid(cells, columns);
}

/** The width of a grid column that the cell-count repair adds when the last one has none. */
const ADDED_COLUMN_WIDTH = 914_400;

/** The cells of a table after the save's cell-count repair (`tableCellCounts`). @internal */
export interface TableCellCounts {
  /**
   * Each row's `a:tc` elements that the save keeps, in order. A row with fewer than `columns`
   * gets blank cells (`addedTableCell`) at its end.
   */
  readonly rows: readonly (readonly Element[])[];
  /** Grid columns after the repair: the `a:gridCol` count, or the widest row when wider. */
  readonly columns: number;
  /** The width (EMU) of each grid column the repair adds: the last column's width. */
  readonly addedColumnWidth: number;
}

/**
 * A cell the repair drops past the grid: no merge attribute, no text, and no field, in any of
 * its text bodies (`cellAsSaved`).
 */
function droppable(tc: Element): boolean {
  if (['gridSpan', 'rowSpan', 'hMerge', 'vMerge'].some((n) => attr(tc, n) !== undefined))
    return false;
  const cell = cellAsSaved(tc);
  const field = elements(child(cell, 'a:txBody')).some((p) =>
    elements(p).some((el) => el.ns === NS.a && el.local === 'fld')
  );
  return !field && !tableCellHasText(cell);
}

/**
 * The save's cell-count repair of `tbl` (ECMA-376 Part 1, 21.1.3.16 tc: one `a:tc` per
 * `a:gridCol` in every row). Blank cells past the grid go first, so a repair that two peers make
 * at once converges; a row that is still longer than the grid adds grid columns at the table's
 * end; a short row gets blank cells at its end. The save writes this, and the cascade draws a
 * marked table by it (`tableAsDrawn`). Undefined when the repair leaves the table as it is: no
 * grid column or no row, or larger than the editor makes (a file would decide the work a repair
 * does). @internal
 */
export function tableCellCounts(tbl: Element | undefined): TableCellCounts | undefined {
  const gridCols = children(child(tbl, 'a:tblGrid'), 'a:gridCol');
  const trs = children(tbl, 'a:tr');
  if (trs.length === 0 || gridCols.length === 0 || trs.length > TABLE_MAX_ROWS) return undefined;
  const raw = trs.map((tr) => children(tr, 'a:tc'));
  if (raw.some((tcs) => tcs.length > TABLE_REPAIR_MAX_CELLS)) return undefined;
  const rows = raw.map((tcs) => {
    let n = tcs.length;
    while (n > gridCols.length && droppable(tcs[n - 1])) n -= 1;
    return n === tcs.length ? tcs : tcs.slice(0, n);
  });
  const widest = rows.reduce((n, tcs) => Math.max(n, tcs.length), 0);
  return {
    rows,
    columns: Math.max(widest, gridCols.length),
    addedColumnWidth: intAttr(gridCols[gridCols.length - 1], 'w') ?? ADDED_COLUMN_WIDTH,
  };
}

/** `a:rPr` children that are not format: links and extensions (CT_TextCharacterProperties). */
const RUN_STATE_CHILDREN = new Set(['hlinkClick', 'hlinkMouseOver', 'extLst']);

/** A detached copy of `el`, every node with the empty id, named `local`, typed under `context`. */
function detached(
  el: Element,
  context: NodeKind | null,
  local = el.local,
  keep: (c: Element) => boolean = () => true
): Element {
  const { kind, childContext } = deriveKind(el.ns, local, context);
  const kids: Node[] = [];
  for (const c of el.children)
    if (c.type === 'text') kids.push({ type: 'text', id: '', value: c.value });
    else if (keep(c)) kids.push(detached(c, childContext));
  return { ...el, id: '', kind, local, children: kids };
}

/** A new DrawingML element with the empty id, typed under `context`. */
function made(
  local: string,
  context: NodeKind | null,
  kids: (context: NodeKind) => Node[] = () => [],
  attrs: Element['attrs'] = []
): Element {
  const { kind } = deriveKind(NS.a, local, context);
  return {
    type: 'element',
    id: '',
    kind,
    ns: NS.a,
    local,
    prefix: 'a',
    nsDecls: [],
    attrs,
    children: kids(kind),
  };
}

/**
 * The blank cell that a table edit or the cell-count repair adds, formatted like `source`: the
 * source's cell properties, its body and list properties, and its first paragraph's properties
 * and character format as the end mark (21.1.3.16 tc, 21.1.2.2.3 endParaRPr). Every node has
 * the empty id: the cascade draws this element where the save adds a cell, and the editor
 * writes a copy of it, so both show the same cell. @internal
 */
export function addedTableCell(source: Element | undefined): Element {
  const body = child(source, 'a:txBody');
  const p = elements(body).find((c) => c.ns === NS.a && c.local === 'p');
  const pPr = child(p, 'a:pPr');
  const rPr =
    elements(p)
      .map((e) => child(e, 'a:rPr'))
      .find((e) => e) ?? child(p, 'a:endParaRPr');
  const bodyPr = child(body, 'a:bodyPr');
  const lstStyle = child(body, 'a:lstStyle');
  const tcPr = child(source, 'a:tcPr');
  // The end mark takes the format, not the links: text typed in the new cell is no link.
  const end = (k: NodeKind): Element =>
    rPr
      ? detached(rPr, k, 'endParaRPr', (c) => !RUN_STATE_CHILDREN.has(c.local))
      : made('endParaRPr', k, undefined, [
          { ns: '', local: 'lang', value: 'en-US' },
          { ns: '', local: 'dirty', value: '0' },
        ]);
  return made('tc', 'generic', (tc) => [
    made('txBody', tc, (b) => [
      bodyPr ? detached(bodyPr, b) : made('bodyPr', b),
      lstStyle ? detached(lstStyle, b) : made('lstStyle', b),
      made('p', b, (k) => [...(pPr ? [detached(pPr, k)] : []), end(k)]),
    ]),
    tcPr ? detached(tcPr, tc) : made('tcPr', tc),
  ]);
}

/** A table as the cascade draws it (`tableAsDrawn`). @internal */
export interface TableAsDrawn {
  /**
   * The cell at each grid position, `cells[row][column]`: the file's `a:tc`, a blank cell the
   * save adds (`addedTableCell`, with empty node ids), or undefined where no cell is drawn.
   */
  readonly cells: readonly (readonly (Element | undefined)[])[];
  /** Grid columns. */
  readonly columns: number;
  /** The width (EMU) of each grid column past the table's `a:gridCol` elements. */
  readonly addedColumnWidth: number;
  /** The merge grid (`tableGridOf`). */
  readonly grid: TableGrid;
}

/**
 * The cells and the merge grid of the table of `frame`, as the cascade draws it and find reads
 * it. A marked table reads as the save writes it: cell counts repaired (`tableCellCounts`) and
 * the merge grid of `tableGridOf`, so the canvas draws the grid of the saved deck. A table
 * without a mark reads with its file's cell counts: a short row draws no cell where it has none,
 * and cells past the grid are not drawn. In every table that the save repairs, a cell with two
 * text bodies reads as one (`cellAsSaved`). Grid columns stop at `maxColumns`. `tbl` is the frame's
 * `a:tbl`, or a table without a frame (never marked). Undefined when the table has no grid
 * column or no row. @internal
 */
export function tableAsDrawn(
  frame: Element | undefined,
  maxColumns: number,
  tbl = child(child(child(frame, 'a:graphic'), 'a:graphicData'), 'a:tbl')
): TableAsDrawn | undefined {
  const counts = tableCellCounts(tbl);
  if (counts && isStructureEdited(frame)) {
    const columns = Math.min(counts.columns, maxColumns);
    const cells = counts.rows.map((tcs) =>
      Array.from({ length: columns }, (_, c) =>
        c < tcs.length ? cellAsSaved(tcs[c]) : addedTableCell(tcs[tcs.length - 1])
      )
    );
    const grid = tableGridOf(frame, cells, columns);
    return { cells, columns, addedColumnWidth: counts.addedColumnWidth, grid };
  }
  const columns = Math.min(children(child(tbl, 'a:tblGrid'), 'a:gridCol').length, maxColumns);
  const cells = children(tbl, 'a:tr')
    .slice(0, TABLE_MAX_ROWS)
    .map((tr) => {
      const tcs = children(tr, 'a:tc');
      // The save joins the text bodies of a cell only in a table it repairs (`counts`).
      return counts ? tcs.map(cellAsSaved) : tcs;
    });
  if (columns === 0 || cells.length === 0) return undefined;
  return { cells, columns, addedColumnWidth: 0, grid: tableGridOf(frame, cells, columns) };
}

/** Whether a node is an `a:txBody` element. */
const isTextBody = (n: Node): n is Element =>
  n.type === 'element' && n.ns === NS.a && n.local === 'txBody';

const isParagraph = (n: Node): n is Element => n.type === 'element' && n.kind === 'paragraph';

/** Whether a text body holds content: a run, a line break, or a field (21.1.2.2.6 p). */
const holdsContent = (body: Element): boolean =>
  body.children.some(
    (p) =>
      isParagraph(p) &&
      p.children.some(
        (c) =>
          c.type === 'element' && (c.kind === 'run' || c.kind === 'break' || c.kind === 'field')
      )
  );

/**
 * The text bodies of a cell whose paragraphs the cell keeps when its bodies join (`cellAsSaved`,
 * and `joinTextBodies` in the table commands): each body that holds content, or the first body
 * when none does. A blank body adds no empty line, so a cell that two peers' table commands
 * each gave an empty body, or one peer's command while the other peer typed, reads as one body
 * would. @internal
 */
export function joinedTextBodies(bodies: readonly Element[]): readonly Element[] {
  const full = bodies.filter(holdsContent);
  return full.length > 0 ? full : bodies.slice(0, 1);
}

/**
 * A cell as the save writes it (`tableCellCounts` not undefined): a cell holds one `a:txBody`
 * (21.1.3.16 tc), and the save moves the paragraphs of a second and later text body to the end
 * of the first, so no text goes. Only the bodies of `joinedTextBodies` give paragraphs; the
 * first body keeps its body and list properties. A file can hold such a cell, and so can a
 * collaboration room where two peers typed into a cell with no text body. The view keeps every
 * node id, so the canvas draws, find reads, and a caret reaches each paragraph that the save
 * keeps, and an edit joins the bodies in the store first (`editor/table-cells.ts`,
 * `selectionBody`). The cell itself when it has one text body or none. @internal
 */
export function cellAsSaved(tc: Element): Element {
  const bodies = tc.children.filter(isTextBody);
  if (bodies.length < 2) return tc;
  const [first, ...rest] = bodies;
  const join = new Set(joinedTextBodies(bodies));
  const paragraphs: Node[] = first.children.filter((n) => join.has(first) || !isParagraph(n));
  for (const b of rest)
    if (join.has(b)) for (const p of b.children) if (isParagraph(p)) paragraphs.push(p);
  const joined: Element = { ...first, children: paragraphs };
  const kids: Node[] = [];
  for (const n of tc.children)
    if (n === first) kids.push(joined);
    else if (!isTextBody(n)) kids.push(n);
  return { ...tc, children: kids };
}
