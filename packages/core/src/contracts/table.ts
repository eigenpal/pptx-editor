// The table types of the public contract (`TableCommand` addresses and the table of the
// selection). Split from `index.ts`, which re-exports them, for the line cap.

import type { ColorChoice, LineDash } from './neutral.ts';

/**
 * A table cell by its row and column, zero-based: the shape of the `cell` that DOCX Editor's
 * `getSelectedTable()` returns. A merged cell is addressed by its top-left grid position.
 * @public
 */
export interface TableCellAddress {
  /** The row, zero-based. */
  readonly row: number;
  /** The grid column, zero-based. */
  readonly column: number;
}

/** Which cell edges a border command sets, as DOCX Editor's `TableBorderTarget`. @public */
export type TableBorderTarget =
  'all' | 'outside' | 'inside' | 'none' | 'top' | 'bottom' | 'left' | 'right';

/** The edges a border is set on: every target but `none`, as DOCX Editor's type. @public */
export type TableBorderEdgeTarget = Exclude<TableBorderTarget, 'none'>;

/**
 * A cell border: color, width in points, and dash. DOCX Editor's spec counts eighths of a point
 * and names a WordprocessingML border style; a DrawingML line (`a:lnL`) has a width and a preset dash, so
 * the fields follow `setLineWidth` and `setLineDash` (and Office.js `BorderProperties.weight`).
 * @public
 */
export interface TableBorderSpec {
  /** The line color. */
  readonly color: ColorChoice;
  /** The line width in points (at most 1,584, `ST_LineWidth`). */
  readonly width: number;
  /** The preset dash; solid when absent. */
  readonly dash?: LineDash;
}

/**
 * Table style options, the `a:tblPr` flags: Header Row, Total Row, Banded Rows, First Column,
 * Last Column, Banded Columns. @public
 */
export type TableStyleOption =
  'firstRow' | 'lastRow' | 'bandRow' | 'firstCol' | 'lastCol' | 'bandCol';

/**
 * A table style the Table Styles picker offers. A deck style carries its `styleName`; the
 * built-in default is `builtIn` with no name, so the chrome names it in the user's language.
 * @public
 */
export interface TableStyleChoice {
  /** The style GUID, `{…}`, as `a:tableStyleId` holds it. */
  readonly id: string;
  /** The style's `styleName`, which can be empty. */
  readonly name: string;
  /** The editor's own Medium Style 2 - Accent 1 (no name): the chrome shows its label. */
  readonly builtIn?: true;
}

/**
 * The table of the selection, for the Table Design and Table Layout chrome: DOCX Editor's
 * `TableContext` (`rows`, `columns`, `rowIndex`, `columnIndex`), with the selected cell range
 * that a slide table selection adds. @public
 */
export interface TableContext {
  /** The table's `p:graphicFrame`. */
  readonly nodeId: string;
  /** The number of rows (`a:tr`). */
  readonly rows: number;
  /** The number of grid columns (`a:gridCol`). */
  readonly columns: number;
  /** The row of the caret, or of the focus cell of a cell range (0 for a selected table). */
  readonly rowIndex: number;
  /** The column of the caret, or of the focus cell of a cell range (0 for a selected table). */
  readonly columnIndex: number;
  /**
   * The selected rows and columns, first to last, inclusive: the shape of DOCX Editor's
   * `getTableCellSelection()` range (`from`, `to`).
   */
  readonly selectedRows: { readonly from: number; readonly to: number };
  /** The selected grid columns, first to last, inclusive. */
  readonly selectedColumns: { readonly from: number; readonly to: number };
  /** The table style id, when the table names one. */
  readonly styleId?: string;
  /** The style options that are on. */
  readonly options: readonly TableStyleOption[];
  /** The first selected row's laid-out height, EMU. */
  readonly rowHeight: number;
  /** The first selected grid column's width, EMU. */
  readonly columnWidth: number;
  /** The first selected cell's margins, EMU. */
  readonly margins: {
    readonly left: number;
    readonly right: number;
    readonly top: number;
    readonly bottom: number;
  };
  /**
   * The first selected cell's fill as drawn (its own or the style's): `#RRGGBB`, `none`,
   * `gradient`, or `other`, as `SelectionGeometry.fill`.
   */
  readonly cellFill?: string;
}

/**
 * Table commands (`table-commands.ts`). Each acts on the table of the selection: a caret in a
 * cell, a cell range, or the selected table frame. The structure commands have DOCX Editor's
 * names; three differ (the style options, the shared Align buttons,
 * and No Border). @public
 */
export type TableCommand =
  /** Insert > Table: a `rows` × `cols` table with the theme's default table style. */
  | { readonly type: 'insertTable'; readonly rows: number; readonly cols: number }
  /** Insert a row above or below the selected rows. */
  | { readonly type: 'insertRow'; readonly where: 'above' | 'below' }
  /** Insert a column left or right of the selected columns. */
  | { readonly type: 'insertColumn'; readonly where: 'left' | 'right' }
  | { readonly type: 'deleteRow' }
  | { readonly type: 'deleteColumn' }
  | { readonly type: 'deleteTable' }
  /** Select the whole table, or the rows or columns of the selected cells. Not an edit. */
  | { readonly type: 'selectTableRegion'; readonly region: 'table' | 'row' | 'column' }
  /** Merge the selected cell range into one cell; the text of every cell is kept. */
  | { readonly type: 'mergeCells' }
  /** Split the selected cell into `rows` × `cols` cells. */
  | { readonly type: 'splitCell'; readonly rows: number; readonly cols: number }
  /** Cell shading: a color choice, or null for No Fill (`a:noFill`, over the style's fill). */
  | { readonly type: 'setCellFill'; readonly color: ColorChoice | null }
  /** Borders of the selected cells, drawn with `spec`; `scope: 'none'` removes them. */
  | { readonly type: 'setTableBorders'; readonly scope: 'none' }
  | {
      readonly type: 'setTableBorders';
      readonly scope: TableBorderEdgeTarget;
      readonly spec: TableBorderSpec;
    }
  /** The table style (`a:tableStyleId`), by `styleId`. */
  | { readonly type: 'setTableStyle'; readonly styleId: string }
  /** Turn one table style option (`a:tblPr` flag) on or off. */
  | { readonly type: 'toggleTableStyleOption'; readonly option: TableStyleOption }
  /** Cell margins in EMU; an absent side keeps its value. */
  | {
      readonly type: 'setCellMargins';
      readonly margins: {
        readonly left?: number;
        readonly right?: number;
        readonly top?: number;
        readonly bottom?: number;
      };
    }
  /** The height of the selected rows in EMU (`a:tr/@h`, a minimum). */
  | { readonly type: 'setRowHeight'; readonly height: number }
  /** The width of the selected columns in EMU (`a:gridCol/@w`). */
  | { readonly type: 'setColumnWidth'; readonly width: number }
  /** Give the selected rows (or every row) the same height. */
  | { readonly type: 'distributeRows' }
  /** Give the selected columns (or every column) the same width. */
  | { readonly type: 'distributeColumns' }
  /**
   * Set every column width and row height of the selected table, in EMU (a grid line drag).
   * Run the steps of one drag with one `historyGroup` (`exec` options) for one undo step.
   */
  | {
      readonly type: 'setTableGrid';
      readonly columns?: readonly number[];
      readonly rows?: readonly number[];
    };
