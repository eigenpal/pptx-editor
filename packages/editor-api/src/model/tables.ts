/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Tables: Office.js `Table`, `TableCell`, `TableRow`, `TableRowCollection`, `TableColumn`, and
// `TableColumnCollection` over `a:tbl` in a graphic frame (ECMA-376 Part 1, 21.1.3.13). Row
// heights and column widths are in points.

import type { AutomationTableRead } from '@pptx-editor.dev/core/automation';
import { ClientResult } from '../runtime/client-result.ts';
import { fail, failArgument, nearest } from '../runtime/errors.ts';
import { hydrated } from '../runtime/hydrate.ts';
import type { ResolvedLoadOptions } from '../runtime/load-options.ts';
import { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import { finite, ModelObject, type RefOf } from './model-object.ts';
import type { Shape } from './shapes.ts';
import { CELL_SPANS, TableCell } from './table-cell.ts';
import { TableStyleSettings } from './table-format.ts';

/** Options for `Table.clear`. @public */
export interface TableClearOptions {
  all?: boolean;
  format?: boolean;
  text?: boolean;
}

/** A positive whole number, or `InvalidArgument`. */
function count(value: unknown, target: string): number {
  if (typeof value !== 'number' || !Number.isInteger(value) || value < 1)
    failArgument(target, 'a whole number of 1 or more');
  return value;
}

/** A zero-based index, or `InvalidArgument`. */
function index(value: unknown, target: string): number {
  if (typeof value !== 'number' || !Number.isInteger(value) || value < 0)
    failArgument(target, 'a whole number of 0 or more');
  return value;
}

/**
 * The table of a table shape.
 *
 * @public
 */
export class Table extends ModelObject {
  readonly #shape: Shape;
  #rows: TableRowCollection | undefined;
  #columns: TableColumnCollection | undefined;
  #styleSettings: TableStyleSettings | undefined;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, shape: Shape) {
    super(context, path);
    this.#shape = shape;
  }

  /** The number of columns. */
  get columnCount(): number {
    return this.loadedProperty('columnCount');
  }
  /** The number of rows. */
  get rowCount(): number {
    return this.loadedProperty('rowCount');
  }
  /** The text of each cell, by row then column. A merged area's text is in its first cell. */
  get values(): string[][] {
    return this.loadedProperty('values');
  }
  /** The rows. */
  get rows(): TableRowCollection {
    return (this.#rows ??= new TableRowCollection(
      this.context,
      ObjectPath.derived(`${this.path.label}.rows`, this.path)
    ));
  }
  /** The columns. */
  get columns(): TableColumnCollection {
    return (this.#columns ??= new TableColumnCollection(
      this.context,
      ObjectPath.derived(`${this.path.label}.columns`, this.path)
    ));
  }

  /** The style options and the style of the table (`a:tblPr`). */
  get styleSettings(): TableStyleSettings {
    return (this.#styleSettings ??= new TableStyleSettings(
      this.context,
      ObjectPath.derived(`${this.path.label}.styleSettings`, this.path),
      this.path
    ));
  }

  /** The merged areas of the table: one cell for each, its first cell, with its span. */
  getMergedAreas(): TableCellCollection {
    return new TableCellCollection(
      this.context,
      ObjectPath.derived(`${this.path.label}.getMergedAreas()`, this.path)
    );
  }

  /** The table shape. */
  getShape(): Shape {
    return this.#shape;
  }

  /** A cell by zero-based row and column, or a null object outside the table. */
  getCellOrNullObject(rowIndex: number, columnIndex: number): TableCell {
    const row = index(rowIndex, `${this.path.label}.getCellOrNullObject`);
    const column = index(columnIndex, `${this.path.label}.getCellOrNullObject`);
    const path = ObjectPath.pending(`${this.path.label}.getCellOrNullObject(${row}, ${column})`);
    const cell = new TableCell(this.context, path, row, column);
    const label = path.label;
    this.enqueue(
      'read',
      `getCellOrNullObject(${row}, ${column})`,
      (ref) => ({ op: 'readTableCell', shape: ref(this.path), row, column }),
      (value) => {
        if (value.kind === 'null') return path.resolveNull();
        hydrated(value, 'tableCell', label);
        path.resolveTo(this.path.handle());
      },
      path,
      false
    );
    return cell;
  }

  /** Merge a rectangle of cells into its first cell. */
  mergeCells(rowIndex: number, columnIndex: number, rowCount: number, columnCount: number): void {
    const target = `${this.path.label}.mergeCells`;
    const op = {
      op: 'mergeTableCells' as const,
      row: index(rowIndex, target),
      column: index(columnIndex, target),
      rowCount: count(rowCount, target),
      columnCount: count(columnCount, target),
    };
    this.command('mergeCells', (ref) => ({ ...op, shape: ref(this.path) }));
  }

  /**
   * Clear the text, the direct format, or both, of every cell. With no options, or `all`, both
   * go. The direct format is each cell's `a:tcPr` and its paragraph and run properties;
   * hyperlinks and the table style stay.
   */
  clear(options?: TableClearOptions): void {
    const target = `${this.path.label}.clear`;
    if (options !== undefined && (typeof options !== 'object' || options === null))
      failArgument(target, 'a TableClearOptions object');
    for (const key of Object.keys(options ?? {}))
      if (key !== 'all' && key !== 'format' && key !== 'text')
        failArgument(
          `${target}.${key}`,
          'the options all, format, or text',
          nearest(key, ['all', 'format', 'text'])
        );
    const { all, format, text: words } = options ?? {};
    for (const value of [all, format, words])
      if (value !== undefined && typeof value !== 'boolean') failArgument(target, 'true or false');
    const everything =
      all === true || (all === undefined && format === undefined && words === undefined);
    const op = {
      op: 'clearTable' as const,
      text: everything || words === true,
      format: everything || format === true,
    };
    if (!op.text && !op.format) return;
    this.command('clear', (ref) => ({ ...op, shape: ref(this.path) }));
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, ['columnCount', 'rowCount', 'values'], {
      rows: () => this.rows,
      columns: () => this.columns,
      styleSettings: () => this.styleSettings,
    });
    this.loadScalars(
      names,
      'table',
      (ref) => ({ op: 'readTable', shape: ref(this.path) }),
      (v) => ({
        columnCount: v.table.columnCount,
        rowCount: v.table.rowCount,
        values: v.table.values.map((row) => [...row]),
      })
    );
  }
}

/** The names a row or a column uses. */
interface LineKind {
  readonly target: 'row' | 'column';
  readonly index: string;
  readonly size: string;
}
const ROW: LineKind = { target: 'row', index: 'rowIndex', size: 'height' };
const COLUMN: LineKind = { target: 'column', index: 'columnIndex', size: 'width' };

/** The index of each line and the way to fill its size. Module-private. */
const LINES = new WeakMap<object, { readonly at: number; size(value: number): void }>();

/** A row or a column of a table. */
abstract class TableLine extends ModelObject {
  readonly #at: number;
  readonly #kind: LineKind;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, at: number, size: number, kind: LineKind) {
    super(context, path);
    this.#at = at;
    this.#kind = kind;
    this.setLoadedProperty(kind.index, at);
    this.setLoadedProperty(kind.size, size);
    LINES.set(this, { at, size: (value) => this.setLoadedProperty(kind.size, value) });
  }

  /** @internal */
  protected get target(): 'row' | 'column' {
    return this.#kind.target;
  }

  /** Delete this line. */
  delete(): void {
    const op = this.target === 'row' ? 'deleteTableRows' : 'deleteTableColumns';
    this.command('delete', (ref) => ({ op, shape: ref(this.path), index: this.#at, count: 1 }));
  }

  /** @internal Set the height or width, in points. */
  protected writeSize(name: string, value: number): void {
    const size = finite(value, name);
    if (size <= 0) failArgument(`${this.path.label}.${name}`, 'a size greater than 0 points');
    this.command(name, (ref) => ({
      op: 'setTableSize',
      shape: ref(this.path),
      target: this.target,
      index: this.#at,
      size,
    }));
  }

  /** @internal Values come with the line; a load has nothing more to read. */
  protected override onLoad(request: ResolvedLoadOptions): void {
    this.selection(request, [this.#kind.index, this.#kind.size, 'currentHeight']);
  }
}

/**
 * One row of a table.
 *
 * @public
 */
export class TableRow extends TableLine {
  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, at: number, size: number) {
    super(context, path, at, size, ROW);
  }
  /** The zero-based row index. */
  get rowIndex(): number {
    return this.loadedProperty('rowIndex');
  }
  /** The row height, in points (`a:tr/@h`). */
  get height(): number {
    return this.loadedProperty('height');
  }
  set height(value: number) {
    this.writeSize('height', value);
  }
  /** The height of the row as stored, in points. */
  get currentHeight(): number {
    return this.loadedProperty('height');
  }
}

/**
 * One column of a table.
 *
 * @public
 */
export class TableColumn extends TableLine {
  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, at: number, size: number) {
    super(context, path, at, size, COLUMN);
  }
  /** The zero-based column index. */
  get columnIndex(): number {
    return this.loadedProperty('columnIndex');
  }
  /** The column width, in points (`a:gridCol/@w`). */
  get width(): number {
    return this.loadedProperty('width');
  }
  set width(value: number) {
    this.writeSize('width', value);
  }
}

/** The rows or the columns of a table. */
abstract class TableLines<T extends TableLine> extends ModelObject {
  /** @internal */
  protected abstract get target(): 'row' | 'column';
  /** @internal */
  protected abstract make(path: ObjectPath, at: number, size: number): T;

  /** The loaded lines. */
  get items(): T[] {
    return this.loadedProperty<T[]>('items');
  }

  /** Insert `count` lines at `index`, or at the end. They copy the line before them. */
  add(index_?: number | null | undefined, count_?: number | undefined): void {
    const at = index_ === null || index_ === undefined ? undefined : index(index_, 'index');
    const n = count_ === undefined ? 1 : count(count_, 'count');
    const op = this.target === 'row' ? 'insertTableRows' : 'insertTableColumns';
    this.command('add', (ref) => ({
      op,
      shape: ref(this.path),
      ...(at === undefined ? {} : { index: at }),
      count: n,
    }));
  }

  /** The number of lines, after the next `sync()`. */
  getCount(): ClientResult<number> {
    const label = `${this.path.label}.getCount`;
    const { result, fill } = ClientResult.create<number>(label);
    this.enqueue(
      'read',
      'getCount',
      (ref) => this.#read(ref),
      (value) => fill(this.#sizes(hydrated(value, 'table', label).table).length)
    );
    return result;
  }

  /** The line at a zero-based position; `ItemNotFound` at the sync when there is none. */
  getItemAt(index_: number): T {
    const at = index(index_, `${this.path.label}.getItemAt`);
    const line = this.make(
      ObjectPath.derived(`${this.path.label}.getItemAt(${at})`, this.path),
      at,
      0
    );
    const label = `${this.path.label}.getItemAt(${at})`;
    this.enqueue(
      'read',
      `getItemAt(${at})`,
      (ref) => this.#read(ref),
      (value) => {
        const size = this.#sizes(hydrated(value, 'table', label).table)[at];
        if (size === undefined) fail({ code: 'ItemNotFound', target: label });
        LINES.get(line)?.size(size);
      }
    );
    return line;
  }

  /** @internal Delete lines, last first so the indices stay valid. */
  protected deleteLines(lines: readonly T[], name: string): void {
    if (!Array.isArray(lines))
      failArgument(`${this.path.label}.${name}`, 'an array of rows or columns of this table');
    const at = (line: T): number => LINES.get(line)?.at ?? 0;
    const sorted = [...lines].sort((a, b) => at(b) - at(a));
    for (const line of sorted) line.delete();
  }

  #read(ref: RefOf): { op: 'readTable'; shape: ReturnType<RefOf> } {
    return { op: 'readTable', shape: ref(this.path) };
  }

  #sizes(table: AutomationTableRead): readonly number[] {
    return this.target === 'row' ? table.rowHeights : table.columnWidths;
  }

  /** @internal */
  protected override get loadsItemPaths(): boolean {
    return true;
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    void request;
    const label = `${this.path.label}.items`;
    this.enqueue(
      'read',
      'items',
      (ref) => this.#read(ref),
      (value) => {
        const sizes = this.#sizes(hydrated(value, 'table', label).table);
        this.setLoadedProperty(
          'items',
          sizes.map((size, i) =>
            this.make(ObjectPath.derived(`${label}[${i}]`, this.path), i, size)
          )
        );
      }
    );
  }
}

/**
 * The rows of a table.
 *
 * @public
 */
export class TableRowCollection extends TableLines<TableRow> {
  /** @internal */
  constructor(context: RequestContext, path: ObjectPath) {
    super(context, path);
  }
  /** Delete rows. */
  deleteRows(rows: TableRow[]): void {
    this.deleteLines(rows, 'deleteRows');
  }
  /** @internal */
  protected get target(): 'row' {
    return 'row';
  }
  /** @internal */
  protected make(path: ObjectPath, at: number, size: number): TableRow {
    return new TableRow(this.context, path, at, size);
  }
}

/**
 * The columns of a table.
 *
 * @public
 */
export class TableColumnCollection extends TableLines<TableColumn> {
  /** @internal */
  constructor(context: RequestContext, path: ObjectPath) {
    super(context, path);
  }
  /** Delete columns. */
  deleteColumns(columns: TableColumn[]): void {
    this.deleteLines(columns, 'deleteColumns');
  }
  /** @internal */
  protected get target(): 'column' {
    return 'column';
  }
  /** @internal */
  protected make(path: ObjectPath, at: number, size: number): TableColumn {
    return new TableColumn(this.context, path, at, size);
  }
}

/** One merged area of a table, as `readTable` gives it. */
type MergedArea = AutomationTableRead['mergedAreas'][number];

/**
 * The merged areas of a table, from `Table.getMergedAreas`. Each item is the first cell of one
 * area, and its `rowCount` and `columnCount` give the span.
 *
 * @public
 */
export class TableCellCollection extends ModelObject {
  /** @internal */
  constructor(context: RequestContext, path: ObjectPath) {
    super(context, path);
  }

  /** The loaded cells. */
  get items(): TableCell[] {
    return this.loadedProperty<TableCell[]>('items');
  }

  /** The number of merged areas, after the next `sync()`. */
  getCount(): ClientResult<number> {
    const label = `${this.path.label}.getCount`;
    const { result, fill } = ClientResult.create<number>(label);
    this.enqueue(
      'read',
      'getCount',
      (ref) => this.#read(ref),
      (value) => fill(hydrated(value, 'table', label).table.mergedAreas.length)
    );
    return result;
  }

  /**
   * The merged area that holds the cell at a zero-based row and column of the table, as its
   * first cell, or a null object when no merged area holds that cell.
   */
  getItemAtOrNullObject(row: number, column: number): TableCell {
    const r = index(row, `${this.path.label}.getItemAtOrNullObject`);
    const c = index(column, `${this.path.label}.getItemAtOrNullObject`);
    const path = ObjectPath.pending(`${this.path.label}.getItemAtOrNullObject(${r}, ${c})`);
    const cell = new TableCell(this.context, path, r, c);
    const label = path.label;
    this.enqueue(
      'read',
      `getItemAtOrNullObject(${r}, ${c})`,
      (ref) => this.#read(ref),
      (value) => {
        const area = hydrated(value, 'table', label).table.mergedAreas.find(
          (a) =>
            r >= a.rowIndex &&
            r < a.rowIndex + a.rowCount &&
            c >= a.columnIndex &&
            c < a.columnIndex + a.columnCount
        );
        if (!area) return path.resolveNull();
        CELL_SPANS.get(cell)?.(spanOf(area));
        path.resolveTo(this.path.handle());
      },
      path,
      false
    );
    return cell;
  }

  #read(ref: RefOf): { op: 'readTable'; shape: ReturnType<RefOf> } {
    return { op: 'readTable', shape: ref(this.path) };
  }

  /** @internal */
  protected override get loadsItemPaths(): boolean {
    return true;
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    void request;
    const label = `${this.path.label}.items`;
    this.enqueue(
      'read',
      'items',
      (ref) => this.#read(ref),
      (value) => {
        const areas = hydrated(value, 'table', label).table.mergedAreas;
        this.setLoadedProperty(
          'items',
          areas.map((area, i) => {
            const cell = new TableCell(
              this.context,
              ObjectPath.derived(`${label}[${i}]`, this.path),
              area.rowIndex,
              area.columnIndex
            );
            CELL_SPANS.get(cell)?.(spanOf(area));
            return cell;
          })
        );
      }
    );
  }
}

function spanOf(area: MergedArea): {
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
} {
  return {
    rowIndex: area.rowIndex,
    columnIndex: area.columnIndex,
    rowCount: area.rowCount,
    columnCount: area.columnCount,
  };
}
