// The chart data types of the public contract: the values a chart plots, as `chartData` reads
// them and `setChartData` writes them. DOM-free.

/**
 * How a chart plots its data. `category`: each series has one value per category (bar, column,
 * line, area, pie, doughnut, and radar charts). `xy`: each series has one Y value per X value
 * (scatter charts); the categories are the X values. @public
 */
export type ChartDataKind = 'category' | 'xy';

/**
 * Where the chart's workbook is, and what an edit does to it. `embedded`: an edit writes the
 * new values into the embedded workbook too. `kept`: the embedded workbook has a structure that
 * the editor does not change, so it keeps its old values. `external`: the workbook is a linked
 * file, which the editor never opens; an edit turns off its automatic update. `none`: the chart
 * has no workbook. @public
 */
export type ChartWorkbookState = 'embedded' | 'kept' | 'external' | 'none';

/** One category (or X value) of a chart. @public */
export interface ChartDataCategory {
  /** The label, or the X value of a scatter chart as text. */
  readonly label: string;
  /**
   * On read: the position of the category in the chart. On write: the category this one
   * continues, so its data point formatting and data labels stay with it. Absent: a new
   * category.
   */
  readonly source?: number;
}

/** One series of a chart. @public */
export interface ChartDataSeries {
  /** The series name, as the legend shows it. */
  readonly name: string;
  /** One value for each category; `null` is a blank cell. */
  readonly values: readonly (number | null)[];
  /**
   * On read: the position of the series in the chart. On write: the series this one continues,
   * so its formatting stays with it. Absent: a new series, with automatic formatting.
   */
  readonly source?: number;
}

/** The data of one chart: its categories and its series. @public */
export interface ChartData {
  readonly kind: ChartDataKind;
  readonly categories: readonly ChartDataCategory[];
  readonly series: readonly ChartDataSeries[];
  /** On read only: where the chart's workbook is. A write ignores it. */
  readonly workbook?: ChartWorkbookState;
}

/**
 * The data of a chart, or why it cannot be edited. `reason` is a translation key, such as
 * `errors.chartData.unsupportedType`. @public
 */
export type ChartDataRead =
  { readonly ok: true; readonly data: ChartData } | { readonly ok: false; readonly reason: string };

/** The selected chart, while exactly one chart frame is selected. @public */
export interface SelectedChartState {
  /** The node id of the chart's graphic frame. */
  readonly nodeId: string;
  /** Whether `setChartData` can edit the chart's data. */
  readonly editable: boolean;
  /** Why the data cannot be edited, as a translation key, when `editable` is false. */
  readonly reason?: string;
}

/** The chart data commands. @public */
export type ChartCommand = {
  /**
   * Replace the data of the selected chart, or of the chart frame `nodeId` on the current slide:
   * its categories, series names, and values, with series and categories added or removed. The
   * values caches, the point counts, and the cell references of the chart change together, and
   * an embedded workbook gets the same values when its structure allows it
   * (`ChartData.workbook`). An edit that an `embedded` workbook cannot take is refused with
   * `errors.chartData.workbook`. Formatting stays with each series and point by its `source`.
   * The chart keeps at least one series and one category.
   */
  readonly type: 'setChartData';
  readonly nodeId?: string;
  readonly data: Omit<ChartData, 'workbook'>;
};
