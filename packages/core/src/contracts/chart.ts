// The chart types of the public contract: the values a chart plots, as `chartData` reads them
// and `setChartData` writes them, and the chart's type and elements, as `chartFormat` reads
// them and `insertChart`, `setChartType`, `setChartFormat`, and `setChartColor` write them.
// DOM-free.

import type { ColorChoice } from './neutral.ts';

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

/**
 * A chart type that `insertChart` and `setChartType` write (ECMA-376 Part 1, 21.2.2): the plot
 * element, and its direction, grouping, or style. Column and bar charts are `c:barChart` with
 * vertical or horizontal bars; `lineMarkers` shows a marker at each point; `scatter` shows
 * markers only, and `scatterLines` joins them with lines. @public
 */
export type ChartType =
  | 'columnClustered'
  | 'columnStacked'
  | 'columnPercentStacked'
  | 'barClustered'
  | 'barStacked'
  | 'barPercentStacked'
  | 'line'
  | 'lineMarkers'
  | 'area'
  | 'areaStacked'
  | 'areaPercentStacked'
  | 'pie'
  | 'doughnut'
  | 'radar'
  | 'scatter'
  | 'scatterLines';

/** Where a chart shows its legend (`c:legendPos`), or `none` for no legend. @public */
export type ChartLegendPosition = 'none' | 'top' | 'bottom' | 'left' | 'right';

/** The color of one series, or of one data point of a pie or doughnut chart. @public */
export interface ChartColorEntry {
  /** The series name, or the category label of a pie or doughnut point. */
  readonly name: string;
  /** The color the chart part sets, or null for the automatic color. */
  readonly color: ColorChoice | null;
  /** The theme color that the automatic color starts from, for a swatch. */
  readonly automatic: ColorChoice;
}

/**
 * The type and elements of a chart, as `chartFormat` reads them. A field is null when the
 * chart type has no such element: a pie chart has no axes. @public
 */
export interface ChartFormat {
  /** The chart type, or null for a type that `setChartType` does not write (3-D, stock). */
  readonly type: ChartType | null;
  /** The title text, or null for no title. An automatic title reads as its series name. */
  readonly title: string | null;
  readonly legend: ChartLegendPosition;
  /** Whether the chart shows a value label at each data point. */
  readonly dataLabels: boolean;
  /** Whether the value axis shows major gridlines; null for a chart with no axes. */
  readonly gridlines: boolean | null;
  /** Whether the chart shows its axes; null for a chart with no axes. */
  readonly axes: boolean | null;
  /** The axis titles, or null for no title; null for a chart with no axes. */
  readonly axisTitles: {
    readonly category: string | null;
    readonly value: string | null;
  } | null;
  /** What `colors` and `setChartColor` color: each series, or each point of a pie chart. */
  readonly colorsBy: 'series' | 'point';
  readonly colors: readonly ChartColorEntry[];
}

/**
 * The type and elements of a chart, or why the editor does not edit them. `reason` is a
 * translation key, such as `errors.chartData.combination`. @public
 */
export type ChartFormatRead =
  | { readonly ok: true; readonly format: ChartFormat }
  | { readonly ok: false; readonly reason: string };

/**
 * The chart elements that `setChartFormat` changes. An absent field keeps its value. A title of
 * null removes the title; a string sets its text and keeps its formatting. @public
 */
export interface ChartFormatChange {
  readonly title?: string | null;
  readonly legend?: ChartLegendPosition;
  readonly dataLabels?: boolean;
  readonly gridlines?: boolean;
  readonly axes?: boolean;
  readonly categoryAxisTitle?: string | null;
  readonly valueAxisTitle?: string | null;
}

/** The data of a new chart: its categories (or X values) and its series. @public */
export interface NewChartData {
  readonly categories: readonly string[];
  readonly series: readonly {
    readonly name: string;
    readonly values: readonly (number | null)[];
  }[];
}

/** The chart commands. @public */
export type ChartCommand =
  | {
      /**
       * Replace the data of the selected chart, or of the chart frame `nodeId` on the current
       * slide: its categories, series names, and values, with series and categories added or
       * removed. The values caches, the point counts, and the cell references of the chart
       * change together, and an embedded workbook gets the same values when its structure
       * allows it (`ChartData.workbook`). An edit that an `embedded` workbook cannot take is
       * refused with `errors.chartData.workbook`. Formatting stays with each series and point
       * by its `source`. The chart keeps at least one series and one category.
       */
      readonly type: 'setChartData';
      readonly nodeId?: string;
      readonly data: Omit<ChartData, 'workbook'>;
    }
  | {
      /**
       * Insert a chart of `chartType` on the current slide, with an embedded workbook that holds
       * the same data. It fills the empty content or chart placeholder that is selected or
       * holds the caret, or else the first one; with none, it is centered on the slide. `data`
       * absent: three series over four categories of sample values. The new chart is selected.
       */
      readonly type: 'insertChart';
      readonly chartType: ChartType;
      readonly data?: NewChartData;
    }
  | {
      /**
       * Change the type of the selected chart, or of the chart frame `nodeId`. Each series
       * keeps its name, data, and color. Formatting that the new type does not have, such as
       * markers, trendlines, and 3-D views, goes.
       */
      readonly type: 'setChartType';
      readonly nodeId?: string;
      readonly chartType: ChartType;
    }
  | {
      /** Show, hide, or change the title, legend, data labels, gridlines, and axes. */
      readonly type: 'setChartFormat';
      readonly nodeId?: string;
      readonly format: ChartFormatChange;
    }
  | {
      /**
       * Set the color of one series (`index` in plot order), or of one point of a pie or
       * doughnut chart (`ChartFormat.colorsBy`). Null gives back the automatic color.
       */
      readonly type: 'setChartColor';
      readonly nodeId?: string;
      readonly index: number;
      readonly color: ColorChoice | null;
    };
