// Table repair: make a table valid again when its structure came from somewhere the table
// commands do not control. Two peers who insert a row and a column at the same time converge on
// a table whose new row lacks the new column's cell; a merge beside a concurrent column insert
// leaves markers that no block claims. Such a table is not valid (ECMA-376 Part 1, 21.1.3.16
// tc: one `a:tc` per `a:gridCol` in every row), so the editor repairs it first. The cell counts
// come from one rule, `tableCellCounts`, which the cascade also draws a marked table by:
//
// - a short row gets empty cells at its end, formatted like its last cell (`addedTableCell`);
// - a long row gets grid columns at the table's end, as wide as the last one, and the other rows
//   get cells for them;
// - blank cells past the grid go first, so a repair that two peers make at once converges;
// - a table command also rewrites every cell's merge markers from the spans of the blocks
//   (`writeBlock`), so an orphan marker becomes a plain cell. Text that a concurrent merge left
//   in a covered cell stays there: the repair never moves text.
//
// A table command repairs the table it edits as part of its own transaction (also when it only
// probes, so `can` agrees with `exec`), and no other table; a selection writes nothing. Typing
// does not repair, so the store keeps the file's merge markup; text typed into a cell that the
// canvas draws and the store lacks adds that row's missing cells first (`selectionBody`). A save
// writes a copy and leaves the store as it is, so two peers that save never send each other the
// same repair:
//
// - every table gets its cell counts repaired, and a cell with two text bodies gets one. The
//   canvas draws a marked table with the same counts (`tableAsDrawn`), so the cells the save
//   adds show before the save, and every table with the joined bodies (`cellAsSaved`). A table
//   without a mark draws its file's cell counts: a short row shows no cell where the save adds
//   a blank one. Its grid gets no column: cells with text past the grid save as the file holds
//   them, and the canvas draws them neither before nor after the save;
// - a table that a structure command changed in a collaboration room (a structure mark,
//   `table-edit-marks.ts`) and whose merge markup fails the grid (`tableMergeFaults`) also gets
//   the merge markers the next table command would write. A merge and a concurrent delete of
//   one of its covered rows leave such a table: the span reaches the row below, which shows
//   again unless its cell is blank. As in the next table command, a span also ends before any
//   other unmarked cell with text that it covers (`trimSpans`). The cascade, the table commands,
//   and find already read a marked table by that grid (`tableGridOf`, `editedTableGrid`), so the
//   save writes what the canvas shows, and a copy of the table goes as a save writes it;
// - a table with a structure mark and no grid column or no row left goes, frame and all
//   (`isEmptiedTable`): peers who delete its last columns at the same time leave it, and a
//   row that one of them inserted can keep a cell of a column the other deleted. The cascade
//   does not draw it either;
// - every other table keeps its merge markup, after typing too.
//
// The marks are part of the room's state, so replicas that hold the same package save the same
// bytes.

import { child, children, elements, type Element } from '../../store/index.ts';
import {
  isEmptiedTable,
  resolveTableGrid,
  tableCellCounts,
  tableCellHasText,
  tableMergeFaults,
  tableSpanTrims,
} from '../../store/index.ts';
import type { PptxPackage } from '../../store/index.ts';
import { CommandRefusal, OpBatch } from './op-batch.ts';
import { isStructureEdited, markedFrames, unmarkStructureEdit } from './table-edit-marks.ts';
import { removeShapes } from './shape-removal.ts';
import {
  aSpec,
  anchorsIn,
  blockOf,
  joinTextBodies,
  newCellSpec,
  readTable,
  setAttr,
  tblOf,
  wholeTable,
  writeBlock,
} from './table-model.ts';

/**
 * End a span before the first cell it claims that has text and no merge marker
 * (`tableSpanTrims`, the rule that `editedTableGrid` reads too). A merge and a concurrent column
 * (or row) delete leave a span over a cell that was never merged; that cell stays visible. A
 * blank cell without a marker stays covered, as a span-only merge writes it.
 */
function trimSpans(batch: OpBatch, part: string, tblId: string): void {
  const rows = children(batch.element(part, tblId), 'a:tr').map((tr) => children(tr, 'a:tc'));
  for (const t of tableSpanTrims(rows, tableCellHasText))
    setAttr(batch, part, rows[t.row][t.col].id, t.name, t.value > 1 ? String(t.value) : null);
}

/**
 * Make the table of graphic frame `frameId` valid in `batch`. Returns whether it pushed ops.
 * `markers` also rewrites the merge markers from the spans (a table command does; a save
 * leaves them, so a plain save never changes how a table from elsewhere looks). A frame that
 * is not a table, a table without a grid or rows, and a table larger than the editor makes
 * (a file decides the work a repair does) are left as they are. `widen` adds grid columns for a
 * row with more cells than the grid (a table command and the save of a marked table do). Without
 * it, the cells with text past the grid stay as the file holds them, and the grid keeps its
 * columns: the canvas draws a table without a mark by its grid, so its save draws the same.
 */
export function repairTable(
  batch: OpBatch,
  part: string,
  frameId: string,
  markers = true,
  widen = true
): boolean {
  const before = batch.ops.length;
  const tbl = tblOf(batch.element(part, frameId));
  const grid = child(tbl, 'a:tblGrid');
  // The cell counts the cascade draws a marked table by (`tableCellCounts`).
  const counts = tableCellCounts(tbl);
  if (!tbl || !grid || !counts) return false;
  const gridCols = children(grid, 'a:gridCol');
  const columns = widen ? counts.columns : gridCols.length;
  // Blank cells past the grid go first: two peers that repair one short row at the same time
  // each add its missing cell, and the row then has one too many.
  children(tbl, 'a:tr').forEach((tr, r) => {
    for (const tc of children(tr, 'a:tc').slice(counts.rows[r].length))
      batch.push({ op: 'removeNode', part, node: tc.id });
  });
  dedupe(batch, part, batch.element(part, tbl.id));
  for (let c = gridCols.length; c < columns; c += 1)
    batch.insert({
      op: 'insertNode',
      part,
      parent: grid.id,
      index: grid.children.length + (c - gridCols.length),
      spec: aSpec('gridCol', { w: String(counts.addedColumnWidth) }),
    });
  for (const tr of children(batch.element(part, tbl.id), 'a:tr')) {
    const tcs = children(tr, 'a:tc');
    if (tcs.length >= columns) continue;
    const at = afterLastCell(tr);
    for (let c = tcs.length; c < columns; c += 1)
      batch.insert({
        op: 'insertNode',
        part,
        parent: tr.id,
        index: at + (c - tcs.length),
        spec: newCellSpec(tcs[tcs.length - 1], {}),
      });
  }
  if (markers) {
    trimSpans(batch, part, tbl.id);
    const m = readTable(batch, part, frameId);
    for (const a of anchorsIn(m, wholeTable(m))) writeBlock(batch, m, blockOf(m, a));
  }
  return batch.ops.length > before;
}

/**
 * Each slot of `a:tcPr`, `a:tc`, and `a:tblPr` that holds one element, by the element names
 * that fill it (CT_TableCellProperties, CT_TableCell, CT_TableProperties).
 */
const SINGLE: Readonly<Record<string, readonly (readonly string[])[]>> = {
  tcPr: [
    ['lnL'],
    ['lnR'],
    ['lnT'],
    ['lnB'],
    ['lnTlToBr'],
    ['lnBlToTr'],
    ['cell3D'],
    ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'],
    ['headers'],
    ['extLst'],
  ],
  tc: [['txBody'], ['tcPr'], ['extLst']],
  tblPr: [
    ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'],
    ['effectLst', 'effectDag'],
    ['tableStyle', 'tableStyleId'],
    ['extLst'],
  ],
};

/**
 * Remove the second and later element of a one-element slot. Two peers that set the same
 * border, fill, or style at once each insert one, and the room keeps both; the first stays,
 * as the cascade reads the first, so the table looks the same.
 */
function dedupe(batch: OpBatch, part: string, tbl: Element): void {
  const once = (el: Element | undefined, slots: readonly (readonly string[])[]): void => {
    for (const names of slots) {
      const hits = elements(el).filter((c) => names.includes(c.local));
      // The cascade prefers an inline `a:tableStyle` wherever it sits: that one stays.
      const keep = hits.find((c) => c.local === 'tableStyle') ?? hits[0];
      const extras = hits.filter((c) => c !== keep);
      // A second text body (a file or two peers' typing can leave one) gives its paragraphs to
      // the first: the repair never drops text, and the canvas draws the joined body
      // (`cellAsSaved`).
      if (keep?.local === 'txBody' && extras.length > 0) joinTextBodies(batch, part, keep, extras);
      else for (const extra of extras) batch.push({ op: 'removeNode', part, node: extra.id });
    }
  };
  once(child(tbl, 'a:tblPr'), SINGLE.tblPr);
  for (const tr of children(tbl, 'a:tr'))
    for (const tc of children(tr, 'a:tc')) {
      once(tc, SINGLE.tc);
      once(child(tc, 'a:tcPr'), SINGLE.tcPr);
      ordered(batch, part, tc.id, ['txBody', 'tcPr', 'extLst']);
      const tcPr = child(batch.element(part, tc.id), 'a:tcPr');
      if (tcPr) ordered(batch, part, tcPr.id, SINGLE.tcPr.flat());
    }
}

/**
 * Put the children of `parentId` named in `order` into that order, keeping the others where
 * they are: two peers that set a top and a left border at once can leave `a:lnT` before
 * `a:lnL`, which the schema refuses (CT_TableCellProperties is a sequence).
 */
function ordered(batch: OpBatch, part: string, parentId: string, order: readonly string[]): void {
  const rank = (c: Element): number => order.indexOf(c.local);
  const kids = elements(batch.element(part, parentId)).filter((c) => rank(c) >= 0);
  const sorted = [...kids].sort((a, b) => rank(a) - rank(b));
  if (sorted.every((c, i) => c === kids[i])) return;
  // Move each one to the end, in order: the named children end up in order after the others.
  // A move inside one parent counts the index without the moved node.
  for (const c of sorted)
    batch.push({
      op: 'moveNode',
      part,
      node: c.id,
      parent: parentId,
      index: batch.element(part, parentId).children.length - 1,
    });
}

/** Index in a row just after its last `a:tc` (before an `a:extLst`). */
function afterLastCell(tr: Element): number {
  let at = 0;
  tr.children.forEach((n, i) => {
    if (n.type === 'element' && n.local === 'tc') at = i + 1;
  });
  return at;
}

/** The table frames of a part, in groups too. */
function tableFrames(root: Element): string[] {
  const out: string[] = [];
  const walk = (el: Element): void => {
    for (const c of elements(el)) {
      if (c.kind === 'graphicFrame') {
        if (tblOf(c)) out.push(c.id);
      } else if (c.kind === 'spTree' || c.kind === 'grpSp' || c.local === 'cSld') walk(c);
    }
  };
  walk(root);
  return out;
}

/**
 * Write the merge markers of each table in a copy (a pasted or duplicated shape, a duplicated
 * slide) from the blocks the cascade draws (`writeBlock`), so the copy draws as its source and
 * holds no orphan marker. The store's part check refuses a merge violation that a transaction
 * adds, and a copy's cells are new nodes: without this, a table with an orphan marker could not
 * be copied. Only the copy changes; the source keeps its markup. `roots` are the copied shapes,
 * or the whole part when absent. A table the editor cannot read stays as it is.
 */
export function writeCopiedTableMarkers(
  batch: OpBatch,
  part: string,
  roots?: readonly string[]
): void {
  const frames: string[] = [];
  if (!roots) {
    const xml = batch.pkg.parts.get(part);
    if (xml?.kind === 'xml') for (const f of tableFrames(xml.xml.root)) frames.push(f);
  } else
    for (const id of roots) {
      const el = batch.element(part, id);
      if (el.kind === 'graphicFrame') {
        if (tblOf(el)) frames.push(el.id);
      } else for (const f of tableFrames(el)) frames.push(f);
    }
  for (const id of frames)
    try {
      const m = readTable(batch, part, id);
      for (const a of anchorsIn(m, wholeTable(m))) writeBlock(batch, m, blockOf(m, a));
    } catch (e) {
      if (!(e instanceof CommandRefusal)) throw e;
    }
}

/** Whether the merge markup of the table of `frameId` fails its grid. False when unreadable. */
function mergeFaulty(batch: OpBatch, part: string, frameId: string): boolean {
  try {
    const m = readTable(batch, part, frameId);
    return tableMergeFaults(m.tcs, resolveTableGrid(m.tcs, m.nC)).length > 0;
  } catch (e) {
    if (!(e instanceof CommandRefusal)) throw e;
    return false;
  }
}

/**
 * Write the table of frame `frameId` as a file holds it: cell counts repaired, and, when the
 * frame has a structure mark and its merge markup fails the grid, the merge markers of
 * `editedTableGrid`, the grid the editor draws it by. Then the marks go. A frame whose marked
 * table has no column or row left (`isEmptiedTable`) goes.
 */
function writeAsSaved(batch: OpBatch, part: string, frameId: string): void {
  const frame = batch.element(part, frameId);
  // Concurrent deletes took the last column or row: the frame goes, as a local delete of the
  // last column or row deletes the table, with its connector ends and animations.
  if (isEmptiedTable(frame)) {
    removeShapes(batch, part, [frameId]);
    return;
  }
  const edited = isStructureEdited(frame);
  try {
    repairTable(batch, part, frameId, false, edited);
    if (edited && mergeFaulty(batch, part, frameId)) repairTable(batch, part, frameId);
  } catch (e) {
    // A table the repair cannot read is saved as it is.
    if (!(e instanceof CommandRefusal)) throw e;
  }
  if (edited) unmarkStructureEdit(batch, part, frameId);
}

/**
 * The package to save: `pkg` itself when every table is valid and unmarked, else a copy with
 * each table repaired and without the structure marks of a collaboration room. A table with a
 * structure mark whose merge markup fails the grid also gets the merge markers the next table
 * command writes. The store does not change.
 */
export function repairTablesForSave(pkg: PptxPackage): PptxPackage {
  const batch = new OpBatch(pkg);
  for (const part of pkg.parts.values()) {
    if (part.kind !== 'xml') continue;
    // A marked frame outside the frames a save repairs (one in `mc:AlternateContent`) still
    // loses its mark, and gets its merge markers repaired.
    const frames = new Set([...tableFrames(part.xml.root), ...markedFrames(part.xml.root)]);
    for (const id of frames) writeAsSaved(batch, part.partName, id);
  }
  return batch.ops.length ? batch.pkg : pkg;
}

/**
 * A shape to copy, with each marked table in it written as a save writes it, so a pasted copy
 * draws as the canvas shows the source. A shape without a marked table is returned as it is;
 * undefined when the save leaves the shape out (`isEmptiedTable`). The store does not change.
 */
export function asSavedForCopy(
  pkg: PptxPackage,
  part: string,
  shape: Element
): Element | undefined {
  const frames = markedFrames(shape);
  if (frames.length === 0) return shape;
  if (isEmptiedTable(shape)) return undefined;
  const batch = new OpBatch(pkg);
  for (const id of frames) writeAsSaved(batch, part, id);
  return batch.element(part, shape.id);
}
