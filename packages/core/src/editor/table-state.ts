// The table of the selection, for the snapshot: its size, the selected cell range, the style
// and its options, and the first selected cell's size, margins, and fill. The function the
// source returns keeps the object it returned last while nothing changes (the snapshot
// contract), so chrome that selects `snapshot.table` re-renders only when it moves.

import { child, nodeIndex, textOf, xmlPartOf, type PptxPackage } from '../store/index.ts';
import type { SlideLayout } from '../layout/index.ts';
import type { EditorSelection, TableContext, TableStyleOption } from '../contracts/index.ts';
import { fillName } from './selection-state.ts';
import { cellLayoutAt, cellRangeOf, expandLayoutRange, tableItem } from './table-cells.ts';
import { tableFrameOf } from './table-commands.ts';
import { styleOption } from '../automation/commands/table-format.ts';
import { tblOf } from '../automation/commands/table-model.ts';

const OPTIONS: readonly TableStyleOption[] = [
  'firstRow',
  'lastRow',
  'bandRow',
  'firstCol',
  'lastCol',
  'bandCol',
];

function stateOf(
  pkg: PptxPackage,
  part: string,
  layout: SlideLayout,
  sel: EditorSelection
): TableContext | null {
  const id = tableFrameOf(sel);
  const frame = id ? tableItem(layout, id) : undefined;
  if (!id || !frame) return null;
  // Only `a:tblPr` is read here; a table that a command must repair first still has chrome.
  const tree = xmlPartOf(pkg, part);
  const el = tree ? nodeIndex(tree).get(id)?.node : undefined;
  const tblPr = el?.type === 'element' ? child(tblOf(el), 'a:tblPr') : undefined;
  const t = frame.table;
  const nR = t.rowY.length - 1;
  const nC = t.columnX.length - 1;
  const raw =
    sel.kind === 'text' && sel.cell
      ? { r0: sel.cell.row, r1: sel.cell.row, c0: sel.cell.column, c1: sel.cell.column }
      : sel.kind === 'cells'
        ? cellRangeOf(sel)
        : { r0: 0, r1: nR - 1, c0: 0, c1: nC - 1 };
  const range = expandLayoutRange(t, raw);
  const first = cellLayoutAt(t, { row: range.r0, column: range.c0 });
  const focus =
    sel.kind === 'text' && sel.cell
      ? sel.cell
      : sel.kind === 'cells'
        ? sel.focus
        : { row: 0, column: 0 };
  const body = first?.cell.text.body;
  const styleId = textOf(child(tblPr, 'a:tableStyleId')).trim();
  return {
    nodeId: id,
    rows: nR,
    columns: nC,
    rowIndex: focus.row,
    columnIndex: focus.column,
    selectedRows: { from: range.r0, to: range.r1 },
    selectedColumns: { from: range.c0, to: range.c1 },
    ...(styleId ? { styleId } : {}),
    options: OPTIONS.filter((o) => styleOption(tblPr, o)),
    rowHeight: Math.round(t.rowY[range.r0 + 1] - t.rowY[range.r0]),
    // `columnX` is in visual order, mirrored for a right-to-left table.
    columnWidth: Math.round(
      t.rtl
        ? t.columnX[nC - range.c0] - t.columnX[nC - range.c0 - 1]
        : t.columnX[range.c0 + 1] - t.columnX[range.c0]
    ),
    margins: {
      left: body?.lIns ?? 91_440,
      right: body?.rIns ?? 91_440,
      top: body?.tIns ?? 45_720,
      bottom: body?.bIns ?? 45_720,
    },
    ...(first ? { cellFill: fillName(first.cell.fill) } : {}),
  };
}

/** A source of `EditorSnapshot.table` that keeps its last object while the state is equal. */
export function createTableStateSource(): (
  pkg: PptxPackage | undefined,
  part: string | undefined,
  layout: SlideLayout | undefined,
  sel: EditorSelection
) => TableContext | null {
  let last: TableContext | null = null;
  let key = 'null';
  return (pkg, part, layout, sel) => {
    const next = pkg && part && layout ? stateOf(pkg, part, layout, sel) : null;
    const k = JSON.stringify(next);
    if (k !== key) {
      key = k;
      last = next;
    }
    return last;
  };
}
