// A cell range on the clipboard. Copy writes the range as tab-separated text (one line for each
// row, a tab between cells; a cell with a tab, a line break, or a quote is quoted, as
// spreadsheets write it) and as this editor's cells format, so a paste into a table fills cells.
// Paste into a cell range, or of tab-separated text into a cell, fills the cells from the
// range's top-left cell (`pasteCells`). DOM-free: the editor passes the clipboard data.

import type { EditorSelection } from '../contracts/index.ts';
import type { SlideLayout } from '../layout/index.ts';
import { cellItem, cellRangeOf, expandLayoutRange, tableItem } from './table-cells.ts';

/** The clipboard type of a copied cell range: JSON rows of cell texts. */
export const CELLS_MIME = 'application/x-pptx-editor-cells+json';

/** The most rows, and cells in a row, a paste reads (the table limit, 75). */
const MAX_LINES = 75;
/** The most characters a pasted cell keeps. */
const MAX_CELL_CHARS = 32_768;

/** The plain text of each cell of a cell range, row by row; a covered position is ''. */
export function cellRangeTexts(
  layout: SlideLayout | undefined,
  sel: EditorSelection
): string[][] | null {
  if (sel.kind !== 'cells') return null;
  const frame = tableItem(layout, sel.nodeId);
  if (!frame) return null;
  const r = expandLayoutRange(frame.table, cellRangeOf(sel));
  const rows: string[][] = [];
  for (let row = r.r0; row <= r.r1; row += 1) {
    const line: string[] = [];
    for (let column = r.c0; column <= r.c1; column += 1) {
      const drawn = frame.table.cells.find((c) => c.cell.row === row && c.cell.col === column);
      const item = drawn ? cellItem(frame, { row, column }) : undefined;
      const paragraphs = item?.element.text?.paragraphs ?? [];
      line.push(
        paragraphs
          .map((p) => p.runs.map((run) => (run.kind === 'break' ? '\n' : run.text)).join(''))
          .join('\n')
      );
    }
    rows.push(line);
  }
  return rows;
}

const quote = (cell: string): string =>
  /[\t\n\r"]/.test(cell) ? `"${cell.replace(/"/g, '""')}"` : cell;

/** Cell texts as tab-separated text. */
export function toTabSeparated(rows: readonly (readonly string[])[]): string {
  return rows.map((row) => row.map(quote).join('\t')).join('\n');
}

/**
 * Tab-separated text as rows of cell texts: quoted cells keep their tabs, line breaks, and
 * doubled quotes; one trailing line break is not a row. At most 75 rows of 75 cells.
 */
export function parseTabSeparated(text: string): string[][] {
  const rows: string[][] = [[]];
  let cell = '';
  let quoted = false;
  let i = 0;
  const push = (): void => {
    rows[rows.length - 1].push(cell.slice(0, MAX_CELL_CHARS));
    cell = '';
  };
  while (i < text.length) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"' && text[i + 1] === '"') ((cell += '"'), (i += 2));
      else if (ch === '"') ((quoted = false), (i += 1));
      else ((cell += ch), (i += 1));
      continue;
    }
    if (ch === '"' && cell === '') {
      quoted = true;
      i += 1;
    } else if (ch === '\t') {
      push();
      i += 1;
    } else if (ch === '\n' || ch === '\r') {
      push();
      rows.push([]);
      i += ch === '\r' && text[i + 1] === '\n' ? 2 : 1;
    } else {
      cell += ch;
      i += 1;
    }
  }
  push();
  // A copy ends with a line break: its empty last row is not a row.
  const last = rows[rows.length - 1];
  if (rows.length > 1 && last.length === 1 && last[0] === '') rows.pop();
  return rows.slice(0, MAX_LINES).map((row) => row.slice(0, MAX_LINES));
}

/** Rows of cell texts from this editor's cells format, or null when it is not valid. */
export function parseCellsJson(raw: string): string[][] | null {
  if (raw.length > MAX_LINES * MAX_LINES * MAX_CELL_CHARS) return null;
  let value: unknown;
  try {
    value = JSON.parse(raw);
  } catch {
    return null;
  }
  if (!Array.isArray(value) || value.length === 0) return null;
  const rows: string[][] = [];
  for (const row of value.slice(0, MAX_LINES)) {
    if (!Array.isArray(row) || !row.every((c) => typeof c === 'string')) return null;
    rows.push((row as string[]).slice(0, MAX_LINES).map((c) => c.slice(0, MAX_CELL_CHARS)));
  }
  return rows;
}

/**
 * The cell texts a paste puts into a table, or null to paste as text. A cell range takes any
 * text; a caret in a cell takes a copied range or tab-separated text, and leaves other text to
 * the cell's own text paste.
 */
export function pastedCells(
  sel: EditorSelection,
  data: { readonly cells?: string | null; readonly text?: string | null }
): string[][] | null {
  const inTable = sel.kind === 'cells' || (sel.kind === 'text' && !!sel.cell);
  if (!inTable) return null;
  const copied = data.cells ? parseCellsJson(data.cells) : null;
  if (copied) return copied;
  const text = data.text ?? '';
  if (!text) return null;
  if (sel.kind === 'cells' || text.includes('\t')) return parseTabSeparated(text);
  return null;
}

/**
 * Write a cell range to a clipboard event's data: tab-separated text, and the cells format that
 * a paste in this editor reads. Returns false when the selection is not a cell range.
 */
export function copyCellsTo(
  data: { setData(format: string, value: string): void },
  layout: SlideLayout | undefined,
  sel: EditorSelection
): boolean {
  const rows = cellRangeTexts(layout, sel);
  if (!rows) return false;
  data.setData(CELLS_MIME, JSON.stringify(rows));
  data.setData('text/plain', toTabSeparated(rows));
  return true;
}

/** The cell texts that a clipboard event pastes into the table of `sel`, or null. */
export function cellsFrom(
  data: { getData(format: string): string },
  sel: EditorSelection
): string[][] | null {
  return pastedCells(sel, { cells: data.getData(CELLS_MIME), text: data.getData('text/plain') });
}
