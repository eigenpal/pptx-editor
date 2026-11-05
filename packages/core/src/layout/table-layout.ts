// Table layout (design D5 of `table-rendering`): column positions from the grid, row heights
// from `a:tr/@h` grown to fit cell text, and one box with laid-out text per drawn cell. All
// lengths are EMU in the frame's local space.

import type { ResolvedTable, ResolvedTableCell } from '../cascade/index.ts';
import { heightIndependent, placeTextLayout, ROTATED_VERT } from './text/anchor.ts';
import { layoutText, type TextLayout } from './text/text-layout.ts';
import type { TextMeasurer } from './text/measurer.ts';

/** One drawn cell: its box in the frame and its text. */
export interface CellLayout {
  readonly cell: ResolvedTableCell;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly text: TextLayout;
}

/** A laid-out table. `columnX` and `rowY` hold every grid line, first to last. */
export interface TableLayout {
  /** Column lines from the frame's left edge, in visual order (mirrored for `rtl`). */
  readonly columnX: readonly number[];
  readonly rowY: readonly number[];
  /** The table runs right to left (`a:tblPr/@rtl`): grid column 0 is drawn at the right. */
  readonly rtl: boolean;
  readonly width: number;
  readonly height: number;
  /** Drawn cells only (merge anchors), in row-major order. */
  readonly cells: readonly CellLayout[];
}

/**
 * The height a cell's text needs, margins included. Rotated text does not grow rows. The
 * layout is kept in `sized` when its lines do not depend on the row height, so the cell is
 * laid out once: only its box and anchoring move to the final row.
 */
function neededHeight(
  cell: ResolvedTableCell,
  width: number,
  measurer: TextMeasurer,
  sized: Map<ResolvedTableCell, TextLayout>
): number {
  if (ROTATED_VERT.has(cell.text.body.vert)) return 0;
  const text = layoutText(cell.text, { x: 0, y: 0, width, height: 0 }, measurer);
  if (heightIndependent(cell.text.body)) sized.set(cell, text);
  return text.contentHeight;
}

/** Lay out a table. `substitutions` collects font substitutions for the report. */
export function layoutTable(
  table: ResolvedTable,
  measurer: TextMeasurer,
  substitutions?: Set<string>
): TableLayout {
  const nC = table.columns.length;
  const nR = table.rows.length;
  const colX = [0];
  for (const w of table.columns) colX.push(colX[colX.length - 1] + w);
  const width = colX[nC];
  const span = (c: number, g: number): number => colX[c + g] - colX[c];
  const drawn = table.cells.filter((c) => !c.covered);
  const sized = new Map<ResolvedTableCell, TextLayout>();

  const rowH = table.rows.map((r) => r.height);
  for (const cell of drawn)
    if (cell.rowSpan === 1)
      rowH[cell.row] = Math.max(
        rowH[cell.row],
        neededHeight(cell, span(cell.col, cell.gridSpan), measurer, sized)
      );
  // A cell spanning rows grows the last row it spans.
  for (const cell of drawn) {
    if (cell.rowSpan === 1) continue;
    const last = cell.row + cell.rowSpan - 1;
    let total = 0;
    for (let r = cell.row; r <= last; r += 1) total += rowH[r];
    const need = neededHeight(cell, span(cell.col, cell.gridSpan), measurer, sized);
    if (need > total) rowH[last] += need - total;
  }
  const rowY = [0];
  for (const h of rowH) rowY.push(rowY[rowY.length - 1] + h);

  const cells = drawn.map((cell): CellLayout => {
    const w = span(cell.col, cell.gridSpan);
    const h = rowY[cell.row + cell.rowSpan] - rowY[cell.row];
    // Right-to-left tables mirror their columns.
    const x = table.rtl ? width - colX[cell.col + cell.gridSpan] : colX[cell.col];
    const y = rowY[cell.row];
    const box = { x, y, width: w, height: h };
    const measured = sized.get(cell);
    const text = measured
      ? placeTextLayout(measured, cell.text.body, box)
      : layoutText(cell.text, box, measurer);
    for (const s of text.substitutions) substitutions?.add(s);
    return { cell, x, y, width: w, height: h, text };
  });
  return {
    columnX: table.rtl ? colX.map((v) => width - v).reverse() : colX,
    rowY,
    rtl: table.rtl,
    width,
    height: rowY[nR],
    cells,
  };
}
