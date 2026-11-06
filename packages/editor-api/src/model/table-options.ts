/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The options of `ShapeCollection.addTable`: Office.js `TableAddOptions` and the data types it
// names (`TableCellProperties`, `TableCellBorders`, `BorderProperties`, `TableCellMargins`,
// `FillProperties`, `TableColumnProperties`, `TableRowProperties`, `TableMergedAreaProperties`).
// Every option is checked before anything is queued, so a bad option adds no table.

import { failArgument, nearest, oneOf } from '../runtime/errors.ts';
import {
  DASH_STYLE,
  HORIZONTAL_ALIGNMENT,
  tableStyleName,
  UNDERLINE,
  VERTICAL_ALIGNMENT,
  type ParagraphHorizontalAlignment,
  type ShapeLineDashStyle,
  type TableStyle,
  type TextVerticalAlignment,
} from './enums.ts';
import { fraction } from './fill-line.ts';
import { ObjectPath } from '../runtime/object-path.ts';
import { finite, pathOf } from './model-object.ts';
import { TableCell, type FontProperties, type TextRun } from './table-cell.ts';
import type { Table } from './tables.ts';

/**
 * One line of a cell, for `TableCellBorders`.
 *
 * @public
 */
export interface BorderProperties {
  /** `#RRGGBB` or a preset color name. */
  color?: string | undefined;
  dashStyle?: ShapeLineDashStyle | `${ShapeLineDashStyle}` | undefined;
  /** From 0 (opaque) to 1 (clear). */
  transparency?: number | undefined;
  /** Points. */
  weight?: number | undefined;
}

/**
 * The lines of a cell, for `TableCellProperties`.
 *
 * @public
 */
export interface TableCellBorders {
  bottom?: BorderProperties | undefined;
  diagonalDown?: BorderProperties | undefined;
  diagonalUp?: BorderProperties | undefined;
  left?: BorderProperties | undefined;
  right?: BorderProperties | undefined;
  top?: BorderProperties | undefined;
}

/**
 * The margins of a cell in points, for `TableCellProperties`.
 *
 * @public
 */
export interface TableCellMargins {
  bottom?: number | undefined;
  left?: number | undefined;
  right?: number | undefined;
  top?: number | undefined;
}

/**
 * A solid fill, for `TableCellProperties`.
 *
 * @public
 */
export interface FillProperties {
  /** `#RRGGBB` or a preset color name. */
  color?: string | undefined;
  /** From 0 (opaque) to 1 (clear). */
  transparency?: number | undefined;
}

/**
 * The format and content of a cell, for `TableAddOptions`.
 *
 * @public
 */
export interface TableCellProperties {
  borders?: TableCellBorders | undefined;
  fill?: FillProperties | undefined;
  font?: FontProperties | undefined;
  horizontalAlignment?:
    ParagraphHorizontalAlignment | `${ParagraphHorizontalAlignment}` | undefined;
  indentLevel?: number | undefined;
  margins?: TableCellMargins | undefined;
  text?: string | undefined;
  textRuns?: TextRun[] | undefined;
  verticalAlignment?: TextVerticalAlignment | `${TextVerticalAlignment}` | undefined;
}

/**
 * One column of a new table, for `TableAddOptions`.
 *
 * @public
 */
export interface TableColumnProperties {
  /** Points. */
  columnWidth?: number | undefined;
}

/**
 * One row of a new table, for `TableAddOptions`.
 *
 * @public
 */
export interface TableRowProperties {
  /** Points. */
  rowHeight?: number | undefined;
}

/**
 * One merged area of a new table, for `TableAddOptions`.
 *
 * @public
 */
export interface TableMergedAreaProperties {
  columnCount: number;
  columnIndex: number;
  rowCount: number;
  rowIndex: number;
}

/**
 * Options for `ShapeCollection.addTable`. They apply in this order: `values`, `style`, the
 * column widths and row heights, `mergedAreas`, `uniformCellProperties`, then
 * `specificCellProperties`.
 *
 * @public
 */
export interface TableAddOptions {
  /** One entry for each column. */
  columns?: TableColumnProperties[] | undefined;
  height?: number | undefined;
  left?: number | undefined;
  mergedAreas?: TableMergedAreaProperties[] | undefined;
  /** One entry for each row. */
  rows?: TableRowProperties[] | undefined;
  /** The properties of each cell, by row then column. Its size must equal the table's. */
  specificCellProperties?: TableCellProperties[][] | undefined;
  style?: TableStyle | `${TableStyle}` | undefined;
  top?: number | undefined;
  /** The properties of every cell. */
  uniformCellProperties?: TableCellProperties | undefined;
  /** The text of each cell, by row then column. Its size must equal the table's. */
  values?: string[][] | undefined;
  width?: number | undefined;
}

/**
 * A cell of a table that the same batch adds. Its path names the new shape directly, so the
 * writes need no read first: a read of a new object cannot go before the batch's writes.
 */
function newCell(table: Table, row: number, column: number): TableCell {
  const at = pathOf(table);
  const label = `${at.label}.getCellOrNullObject(${row}, ${column})`;
  return new TableCell(table.context, ObjectPath.derived(label, at), row, column);
}

/** The checked options, ready to apply after the table exists. */
export interface TablePlan {
  readonly values: string[][] | undefined;
  readonly apply: (table: Table) => void;
}

/** An options object with only `names`, or `InvalidArgument` naming the first other key. */
function record(value: unknown, names: readonly string[], target: string): Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value))
    failArgument(target, 'an object of options');
  for (const key of Object.keys(value))
    if (!names.includes(key))
      failArgument(`${target}.${key}`, `an option name: ${oneOf(names)}`, nearest(key, names));
  return value as Record<string, unknown>;
}

function string(value: unknown, target: string): string {
  if (typeof value !== 'string') failArgument(target, 'a string');
  return value;
}

function positive(value: unknown, target: string): number {
  const n = finite(value, target);
  if (n <= 0) failArgument(target, 'a number greater than 0');
  return n;
}

function nonNegative(value: unknown, target: string): number {
  const n = finite(value, target);
  if (n < 0) failArgument(target, 'a number of 0 or more');
  return n;
}

function wholeIn(value: unknown, min: number, max: number, target: string): number {
  if (typeof value !== 'number' || !Number.isInteger(value) || value < min || value > max)
    failArgument(target, `a whole number from ${min} to ${max}`);
  return value;
}

/** An array of `length` entries, or `InvalidArgument`. */
function list(value: unknown, length: number | undefined, target: string): unknown[] {
  if (!Array.isArray(value) || (length !== undefined && value.length !== length))
    failArgument(
      target,
      length === undefined
        ? 'an array'
        : `an array of ${length} entries, one for each ${target.endsWith('values') || target.endsWith('rows') || target.endsWith('specificCellProperties') ? 'row' : 'entry'}`
    );
  return value;
}

const BORDER_KEYS = ['color', 'dashStyle', 'transparency', 'weight'];
const SIDES = ['bottom', 'diagonalDown', 'diagonalUp', 'left', 'right', 'top'] as const;
const MARGIN_KEYS = ['bottom', 'left', 'right', 'top'] as const;
const FONT_KEYS = [
  'allCaps',
  'bold',
  'color',
  'doubleStrikethrough',
  'italic',
  'name',
  'size',
  'smallCaps',
  'strikethrough',
  'subscript',
  'superscript',
  'underline',
] as const;
const CELL_KEYS = [
  'borders',
  'fill',
  'font',
  'horizontalAlignment',
  'indentLevel',
  'margins',
  'text',
  'textRuns',
  'verticalAlignment',
];

/** Check one `FontProperties`, as the `ShapeFont` setters check each value. */
function checkFont(value: unknown, target: string): void {
  const font = record(value, FONT_KEYS, target);
  for (const key of FONT_KEYS) {
    const v = font[key];
    const at = `${target}.${key}`;
    if (v === undefined) continue;
    if (key === 'underline') UNDERLINE.toFile(v, at);
    else if (key === 'name' || key === 'color') string(v, at);
    else if (key === 'size') positive(v, at);
    else if (typeof v !== 'boolean') failArgument(at, 'true or false');
  }
}

/** Check one `TableCellProperties`; answer the function that applies it to a cell. */
function checkCell(value: unknown, target: string): (cell: TableCell) => void {
  const props = record(value, CELL_KEYS, target);
  const steps: ((cell: TableCell) => void)[] = [];
  const { borders, fill, font, horizontalAlignment, indentLevel, margins, text, textRuns } = props;
  if (text !== undefined && textRuns !== undefined)
    failArgument(`${target}.textRuns`, 'text or textRuns, not both');
  if (text !== undefined) {
    const t = string(text, `${target}.text`);
    steps.push((cell) => (cell.text = t));
  }
  if (textRuns !== undefined) {
    const runs = list(textRuns, undefined, `${target}.textRuns`).map((run, i) => {
      const r = record(run, ['font', 'text'], `${target}.textRuns[${i}]`);
      string(r.text, `${target}.textRuns[${i}].text`);
      if (r.font !== undefined) checkFont(r.font, `${target}.textRuns[${i}].font`);
      return run as TextRun;
    });
    steps.push((cell) => (cell.textRuns = runs));
  }
  if (font !== undefined) {
    checkFont(font, `${target}.font`);
    const f = font as Record<string, unknown>;
    steps.push((cell) => {
      for (const key of FONT_KEYS)
        if (f[key] !== undefined) (cell.font as unknown as Record<string, unknown>)[key] = f[key];
    });
  }
  if (fill !== undefined) {
    const p = record(fill, ['color', 'transparency'], `${target}.fill`);
    const color = p.color === undefined ? undefined : string(p.color, `${target}.fill.color`);
    const t =
      p.transparency === undefined
        ? undefined
        : fraction(p.transparency, `${target}.fill.transparency`);
    steps.push((cell) => {
      if (color !== undefined) cell.fill.setSolidColor(color);
      if (t !== undefined) cell.fill.transparency = t;
    });
  }
  if (horizontalAlignment !== undefined) {
    HORIZONTAL_ALIGNMENT.toFile(horizontalAlignment, `${target}.horizontalAlignment`);
    const a = horizontalAlignment as `${ParagraphHorizontalAlignment}`;
    steps.push((cell) => (cell.horizontalAlignment = a));
  }
  if (indentLevel !== undefined) {
    const n = finite(indentLevel, `${target}.indentLevel`);
    steps.push((cell) => (cell.indentLevel = n));
  }
  if (props.verticalAlignment !== undefined) {
    VERTICAL_ALIGNMENT.toFile(props.verticalAlignment, `${target}.verticalAlignment`);
    const a = props.verticalAlignment as `${TextVerticalAlignment}`;
    steps.push((cell) => (cell.verticalAlignment = a));
  }
  if (margins !== undefined) {
    const m = record(margins, MARGIN_KEYS, `${target}.margins`);
    for (const key of MARGIN_KEYS)
      if (m[key] !== undefined) {
        const n = nonNegative(m[key], `${target}.margins.${key}`);
        steps.push((cell) => (cell.margins[key] = n));
      }
  }
  if (borders !== undefined) {
    const b = record(borders, SIDES, `${target}.borders`);
    for (const side of SIDES) {
      if (b[side] === undefined) continue;
      const at = `${target}.borders.${side}`;
      const line = record(b[side], BORDER_KEYS, at);
      if (line.color !== undefined) string(line.color, `${at}.color`);
      if (line.dashStyle !== undefined) DASH_STYLE.toFile(line.dashStyle, `${at}.dashStyle`);
      if (line.transparency !== undefined) fraction(line.transparency, `${at}.transparency`);
      if (line.weight !== undefined) nonNegative(line.weight, `${at}.weight`);
      const l = line as BorderProperties;
      steps.push((cell) => {
        const border = cell.borders[side];
        if (l.color !== undefined) border.color = l.color;
        if (l.weight !== undefined) border.weight = l.weight;
        if (l.dashStyle !== undefined) border.dashStyle = l.dashStyle;
        if (l.transparency !== undefined) border.transparency = l.transparency;
      });
    }
  }
  return (cell) => {
    for (const step of steps) step(cell);
  };
}

const TABLE_KEYS = [
  'columns',
  'height',
  'left',
  'mergedAreas',
  'rows',
  'specificCellProperties',
  'style',
  'top',
  'uniformCellProperties',
  'values',
  'width',
];

/** Check `TableAddOptions` for a `rows` by `columns` table, before anything is queued. */
export function tablePlan(
  options: TableAddOptions | undefined,
  rows: number,
  columns: number,
  target: string
): TablePlan {
  if (options === undefined) return { values: undefined, apply: () => undefined };
  const o = record(options, TABLE_KEYS, target);
  const steps: ((table: Table) => void)[] = [];
  let values: string[][] | undefined;
  if (o.values !== undefined)
    values = list(o.values, rows, `${target}.values`).map((row) =>
      list(row, columns, `${target}.values`).map((v) => string(v, `${target}.values`))
    );
  if (o.style !== undefined) {
    tableStyleName(o.style, `${target}.style`);
    const style = o.style as `${TableStyle}`;
    steps.push((table) => (table.styleSettings.style = style));
  }
  if (o.columns !== undefined)
    list(o.columns, columns, `${target}.columns`).forEach((entry, i) => {
      const c = record(entry, ['columnWidth'], `${target}.columns[${i}]`);
      if (c.columnWidth === undefined) return;
      const w = positive(c.columnWidth, `${target}.columns[${i}].columnWidth`);
      steps.push((table) => (table.columns.getItemAt(i).width = w));
    });
  if (o.rows !== undefined)
    list(o.rows, rows, `${target}.rows`).forEach((entry, i) => {
      const r = record(entry, ['rowHeight'], `${target}.rows[${i}]`);
      if (r.rowHeight === undefined) return;
      const h = positive(r.rowHeight, `${target}.rows[${i}].rowHeight`);
      steps.push((table) => (table.rows.getItemAt(i).height = h));
    });
  if (o.mergedAreas !== undefined)
    list(o.mergedAreas, undefined, `${target}.mergedAreas`).forEach((entry, i) => {
      const at = `${target}.mergedAreas[${i}]`;
      const a = record(entry, ['columnCount', 'columnIndex', 'rowCount', 'rowIndex'], at);
      const r0 = wholeIn(a.rowIndex, 0, rows - 1, `${at}.rowIndex`);
      const c0 = wholeIn(a.columnIndex, 0, columns - 1, `${at}.columnIndex`);
      const rc = wholeIn(a.rowCount, 1, rows - r0, `${at}.rowCount`);
      const cc = wholeIn(a.columnCount, 1, columns - c0, `${at}.columnCount`);
      steps.push((table) => table.mergeCells(r0, c0, rc, cc));
    });
  if (o.uniformCellProperties !== undefined) {
    const apply = checkCell(o.uniformCellProperties, `${target}.uniformCellProperties`);
    steps.push((table) => {
      for (let r = 0; r < rows; r += 1)
        for (let c = 0; c < columns; c += 1) apply(newCell(table, r, c));
    });
  }
  if (o.specificCellProperties !== undefined) {
    const at = `${target}.specificCellProperties`;
    const grid = list(o.specificCellProperties, rows, at).map((row, r) =>
      list(row, columns, `${at}[${r}]`).map((cell, c) => checkCell(cell, `${at}[${r}][${c}]`))
    );
    steps.push((table) =>
      grid.forEach((row, r) => row.forEach((apply, c) => apply(newCell(table, r, c))))
    );
  }
  return {
    values,
    apply: (table) => {
      for (const step of steps) step(table);
    },
  };
}
