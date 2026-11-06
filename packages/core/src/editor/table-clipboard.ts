// A cell range on the clipboard. Copy writes the range as tab-separated text (one line for each
// row, a tab between cells; a cell with a tab, a line break, or a quote is quoted, as
// spreadsheets write it) and as this editor's cells format, so a paste into a table fills cells.
// Paste into a cell range, or of tab-separated text into a cell, fills the cells from the
// range's top-left cell (`pasteCells`). DOM-free: the editor passes the clipboard data.

import type { EditorSelection } from '../contracts/index.ts';
import type { SlideLayout } from '../layout/index.ts';
import { cellItem, cellRangeOf, selectedLayoutRange, tableItem } from './table-cells.ts';

/** The clipboard type of a copied cell range: JSON rows of cell texts. */
export const CELLS_MIME = 'application/x-pptx-editor-cells+json';

/** The most rows, and cells in a row, a paste reads (the table limit, 75). */
const MAX_LINES = 75;
/** The most characters a pasted cell keeps. */
const MAX_CELL_CHARS = 32_768;

/** The most clipboard characters a paste reads, before it parses them. */
const MAX_RAW_CHARS = 4_000_000;

/** Control characters a cell cannot hold (all but tab and line break); a paste drops them. */
const CONTROL = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g;

/** A pasted cell's text: no control characters, at most the cell limit. */
const cellText = (text: string): string => text.replace(CONTROL, '').slice(0, MAX_CELL_CHARS);

/** The plain text of each cell of a cell range, row by row; a covered position is ''. */
export function cellRangeTexts(
  layout: SlideLayout | undefined,
  sel: EditorSelection
): string[][] | null {
  if (sel.kind !== 'cells') return null;
  const frame = tableItem(layout, sel.nodeId);
  if (!frame) return null;
  const r = selectedLayoutRange(frame.table, cellRangeOf(sel)).range;
  // The drawn cells by grid position, read once: a 75 by 75 range looks each up at once.
  const drawn = new Set(frame.table.cells.map((c) => `${c.cell.row}:${c.cell.col}`));
  const rows: string[][] = [];
  for (let row = r.r0; row <= r.r1; row += 1) {
    const line: string[] = [];
    for (let column = r.c0; column <= r.c1; column += 1) {
      const item = drawn.has(`${row}:${column}`) ? cellItem(frame, { row, column }) : undefined;
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
export function parseTabSeparated(raw: string): string[][] {
  const text = raw.slice(0, MAX_RAW_CHARS);
  const rows: string[][] = [[]];
  let cell = '';
  let quoted = false;
  let i = 0;
  const push = (): void => {
    // Cells past the limit in a row are never kept.
    const row = rows[rows.length - 1];
    if (row.length < MAX_LINES) row.push(cellText(cell));
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
      // Rows past the limit are never read.
      if (rows.length > MAX_LINES) break;
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
  if (raw.length > MAX_RAW_CHARS) return null;
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
    rows.push((row as string[]).slice(0, MAX_LINES).map(cellText));
  }
  return rows;
}

/** What a paste does in a table: fill cells, insert one cell's text at the caret, or neither. */
export type TablePaste = { readonly cells: string[][] } | { readonly text: string } | null;

/** Text in quotes, with only doubled quotes inside: each step has one choice, so linear. */
const QUOTED = /^"(?:[^"]|"")*"$/;
/** What made a copy quote a cell: a tab, a line break, or a quote (doubled). */
const NEEDS_QUOTES = /[\t\r\n]|""/;

/**
 * Whether `text` is one quoted cell, as a copy of one cell writes it: quoted only because it
 * holds a tab, a line break, or a quote. Other text in quotes ("Hello") is text, and keeps its
 * quotes. Two linear tests on text within the paste limit: untrusted text never backtracks.
 */
const oneQuotedCell = (text: string): boolean =>
  text.length <= MAX_RAW_CHARS && QUOTED.test(text) && NEEDS_QUOTES.test(text.slice(1, -1));

/**
 * What a paste does in a table. A cell range takes any text as cells. A caret in a cell takes a
 * copied range of more than one cell, or tab-separated text, as cells; one copied cell (from the
 * cells format, or as one quoted cell of plain text) goes in as text at the caret, unquoted.
 * Other text is null: the cell's own text paste. Outside a table, null.
 */
export function tablePaste(
  sel: EditorSelection,
  data: { readonly cells?: string | null; readonly text?: string | null }
): TablePaste {
  const inTable = sel.kind === 'cells' || (sel.kind === 'text' && !!sel.cell);
  if (!inTable) return null;
  const copied = data.cells ? parseCellsJson(data.cells) : null;
  const text = data.text ?? '';
  const rows =
    copied ??
    (text && (sel.kind === 'cells' || text.includes('\t') || oneQuotedCell(text))
      ? parseTabSeparated(text)
      : null);
  if (!rows) return null;
  const one = rows.length === 1 && rows[0].length === 1;
  return sel.kind === 'text' && one ? { text: rows[0][0] } : { cells: rows };
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

/** What a clipboard event's paste does in the table of `sel` (`tablePaste`). */
export function tablePasteFrom(
  data: { getData(format: string): string },
  sel: EditorSelection
): TablePaste {
  return tablePaste(sel, { cells: data.getData(CELLS_MIME), text: data.getData('text/plain') });
}
