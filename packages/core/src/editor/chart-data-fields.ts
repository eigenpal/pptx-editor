// Field rules of the Edit Data dialog: a grid of categories (rows) by series (columns), as text
// the user types. What the dialog opens with, the row and column edits, what blocks OK, the
// command OK applies, and where the arrow keys move in the grid. Framework-free, so the React
// and Vue dialogs agree. Every export is adapter glue, not a stable API: `@internal`.

import type {
  ChartData,
  ChartDataKind,
  ChartWorkbookState,
  EditorCommand,
} from '../contracts/index.ts';
import type { DialogErrors } from './dialog-draft.ts';
import {
  MAX_CHART_LABEL,
  MAX_CHART_POINTS,
  MAX_CHART_SERIES,
} from '../automation/commands/chart-data-model.ts';

export { MAX_CHART_POINTS, MAX_CHART_SERIES };

/** One row of the grid: a category (or X value) and its cell in each series. @internal */
export interface ChartDraftRow {
  readonly label: string;
  /** The category this row continues, or undefined for a new row. */
  readonly source?: number;
}

/** One column of the grid: a series. @internal */
export interface ChartDraftColumn {
  readonly name: string;
  /** The series this column continues, or undefined for a new column. */
  readonly source?: number;
  /** One cell text for each row. */
  readonly cells: readonly string[];
}

/** The draft of the Edit Data dialog. @internal */
export interface ChartDataDraft {
  readonly kind: ChartDataKind;
  readonly rows: readonly ChartDraftRow[];
  readonly columns: readonly ChartDraftColumn[];
  /** Where the chart's workbook is, for the dialog's note. */
  readonly workbook: ChartWorkbookState;
  /** Whether the first column holds numbers: the X values of a scatter chart that are not text. */
  readonly numericX: boolean;
}

/** The fields of the dialog's draft. @internal */
export interface ChartDataDraftFields {
  readonly grid: ChartDataDraft;
}

/** A number as a cell shows it. */
const cellText = (v: number | null): string => (v === null ? '' : String(v));

/** The draft the dialog opens with: the chart's data as text. @internal */
export function chartDraftFrom(data: ChartData): ChartDataDraftFields {
  return {
    grid: {
      kind: data.kind,
      workbook: data.workbook ?? 'none',
      numericX:
        data.kind === 'xy' && data.categories.every((c) => typeof parseCell(c.label) === 'number'),
      rows: data.categories.map((c) => ({ label: c.label, ...source(c.source) })),
      columns: data.series.map((s) => ({
        name: s.name,
        ...source(s.source),
        cells: s.values.map(cellText),
      })),
    },
  };
}

const source = (s: number | undefined): { source?: number } =>
  s === undefined ? {} : { source: s };

/** The draft with the text of one grid cell changed. Row 0 is the series names; column 0 the labels. @internal */
export function setChartDraftCell(
  draft: ChartDataDraft,
  row: number,
  column: number,
  text: string
): ChartDataDraft {
  if (column === 0) {
    if (row === 0) return draft;
    return {
      ...draft,
      rows: draft.rows.map((r, i) => (i === row - 1 ? { ...r, label: text } : r)),
    };
  }
  return {
    ...draft,
    columns: draft.columns.map((c, j) => {
      if (j !== column - 1) return c;
      if (row === 0) return { ...c, name: text };
      return { ...c, cells: c.cells.map((v, i) => (i === row - 1 ? text : v)) };
    }),
  };
}

/** The text of one grid cell. @internal */
export function chartDraftCell(draft: ChartDataDraft, row: number, column: number): string {
  if (column === 0) return row === 0 ? '' : (draft.rows[row - 1]?.label ?? '');
  const col = draft.columns[column - 1];
  if (!col) return '';
  return row === 0 ? col.name : (col.cells[row - 1] ?? '');
}

/**
 * The draft with a new, empty category before row `at` (1-based; past the end appends), with
 * the label `label`. A new scatter row gets the next X value instead. @internal
 */
export function insertChartDraftRow(
  draft: ChartDataDraft,
  at: number,
  label: string
): ChartDataDraft {
  if (draft.rows.length >= MAX_CHART_POINTS) return draft;
  const index = Math.max(0, Math.min(at - 1, draft.rows.length));
  const text = draft.kind === 'xy' ? String(nextX(draft.rows)) : label;
  const rows = [...draft.rows];
  rows.splice(index, 0, { label: text });
  return {
    ...draft,
    rows,
    columns: draft.columns.map((c) => {
      const cells = [...c.cells];
      cells.splice(index, 0, '');
      return { ...c, cells };
    }),
  };
}

function nextX(rows: readonly ChartDraftRow[]): number {
  const xs = rows.map((r) => Number(r.label)).filter(Number.isFinite);
  return xs.length ? Math.max(...xs) + 1 : 1;
}

/** The draft without category row `row` (1-based). The last category stays. @internal */
export function removeChartDraftRow(draft: ChartDataDraft, row: number): ChartDataDraft {
  if (draft.rows.length <= 1 || row < 1 || row > draft.rows.length) return draft;
  return {
    ...draft,
    rows: draft.rows.filter((_, i) => i !== row - 1),
    columns: draft.columns.map((c) => ({ ...c, cells: c.cells.filter((_, i) => i !== row - 1) })),
  };
}

/**
 * The draft with a new, empty series before column `at` (1-based; past the end appends), named
 * `name`. @internal
 */
export function insertChartDraftColumn(
  draft: ChartDataDraft,
  at: number,
  name: string
): ChartDataDraft {
  if (draft.columns.length >= MAX_CHART_SERIES) return draft;
  const index = Math.max(0, Math.min(at - 1, draft.columns.length));
  const columns = [...draft.columns];
  columns.splice(index, 0, { name, cells: draft.rows.map(() => '') });
  return { ...draft, columns };
}

/** The draft without series column `column` (1-based). The last series stays. @internal */
export function removeChartDraftColumn(draft: ChartDataDraft, column: number): ChartDataDraft {
  if (draft.columns.length <= 1 || column < 1 || column > draft.columns.length) return draft;
  return { ...draft, columns: draft.columns.filter((_, j) => j !== column - 1) };
}

/** A cell's number, null for an empty cell, or undefined for text that is not a number. */
function parseCell(text: string): number | null | undefined {
  const t = text.trim();
  if (t === '') return null;
  const n = Number(t);
  return Number.isFinite(n) ? n : undefined;
}

/**
 * The translation key of the problem in each grid cell, by `row,column`: a value or a scatter X
 * value that is not a number, or text longer than a chart takes. @internal
 */
export function chartDraftCellErrors(draft: ChartDataDraft): ReadonlyMap<string, string> {
  const out = new Map<string, string>();
  const long = (s: string) => s.length > MAX_CHART_LABEL;
  draft.rows.forEach((r, i) => {
    if (long(r.label)) out.set(`${i + 1},0`, 'chartData.errors.tooLong');
    else if (draft.numericX && typeof parseCell(r.label) !== 'number')
      out.set(`${i + 1},0`, 'chartData.errors.xNumber');
  });
  draft.columns.forEach((c, j) => {
    if (long(c.name)) out.set(`0,${j + 1}`, 'chartData.errors.tooLong');
    c.cells.forEach((v, i) => {
      if (parseCell(v) === undefined) out.set(`${i + 1},${j + 1}`, 'chartData.errors.number');
    });
  });
  return out;
}

/** Errors that block OK: the first cell problem, as the `grid` field's error. @internal */
export function validateChartDraft(
  fields: ChartDataDraftFields
): DialogErrors<ChartDataDraftFields> {
  const first = chartDraftCellErrors(fields.grid).values().next();
  return first.done ? {} : { grid: first.value };
}

/** The command OK applies, or null when the draft has a cell problem. @internal */
export function chartDraftCommand(fields: ChartDataDraftFields): EditorCommand | null {
  const draft = fields.grid;
  if (chartDraftCellErrors(draft).size > 0) return null;
  const data: Omit<ChartData, 'workbook'> = {
    kind: draft.kind,
    categories: draft.rows.map((r) => ({
      label: draft.numericX ? String(parseCell(r.label)) : r.label,
      ...source(r.source),
    })),
    series: draft.columns.map((c) => ({
      name: c.name,
      values: c.cells.map((v) => parseCell(v) ?? null),
      ...source(c.source),
    })),
  };
  return { type: 'setChartData', data };
}

/** A position in the grid: row 0 is the series names, column 0 the category labels. @internal */
export interface ChartGridPosition {
  readonly row: number;
  readonly column: number;
}

/**
 * Where a key moves the focus in a grid of `rows` × `columns` cells (headers included), or
 * null when the key stays in the cell. The arrow keys move between cells, except Left and Right
 * inside text, which move the caret first; Enter moves down and Shift+Enter up; Home and End
 * with Ctrl (Cmd) go to the first and last cell. The corner cell has no field, so a move
 * never lands on it. @internal
 */
export function chartGridMove(
  key: { readonly key: string; readonly shiftKey?: boolean; readonly ctrlKey?: boolean },
  at: ChartGridPosition,
  size: { readonly rows: number; readonly columns: number },
  caret: { readonly atStart: boolean; readonly atEnd: boolean } = { atStart: true, atEnd: true }
): ChartGridPosition | null {
  let { row, column } = at;
  switch (key.key) {
    case 'ArrowUp':
      row -= 1;
      break;
    case 'ArrowDown':
      row += 1;
      break;
    case 'ArrowLeft':
      if (!caret.atStart) return null;
      column -= 1;
      break;
    case 'ArrowRight':
      if (!caret.atEnd) return null;
      column += 1;
      break;
    case 'Enter':
      row += key.shiftKey ? -1 : 1;
      break;
    case 'Home':
      if (!key.ctrlKey) return null;
      [row, column] = [0, 1];
      break;
    case 'End':
      if (!key.ctrlKey) return null;
      [row, column] = [size.rows - 1, size.columns - 1];
      break;
    default:
      return null;
  }
  if (row < 0 || column < 0 || row >= size.rows || column >= size.columns) return null;
  if (row === 0 && column === 0) return null;
  return { row, column };
}
