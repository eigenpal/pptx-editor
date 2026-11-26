/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Charts, an SDK extension: the chart part (`c:chartSpace`, ECMA-376 Part 1, §21.2.2.29) that a
// chart frame names (`c:chart`, `CT_RelId`, §21.2.2.26). The pinned Office.js reference has no
// chart members. `Chart.chartType` and the `ChartType` values take the Office.js chart names;
// the data members are plain values, not ranges.

import type { AutomationChartData, AutomationChartRead } from '@pptx-editor.dev/core/automation';
import { failArgument, nearest, oneOf } from '../runtime/errors.ts';
import type { ResolvedLoadOptions } from '../runtime/load-options.ts';
import { pathOf } from './model-object.ts';
import type { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import { ModelObject } from './model-object.ts';
import type { Shape } from './shapes.ts';

/** A chart type as the engine names it (`columnClustered`, `scatter`, …). */
type ChartTypeOfFile = NonNullable<AutomationChartRead['type']>;

/**
 * A chart type that the engine reads and writes: the plot element (`c:barChart`, `c:lineChart`,
 * `c:areaChart`, `c:pieChart`, `c:doughnutChart`, `c:radarChart`, `c:scatterChart`) with its
 * direction, grouping, or style (ECMA-376 Part 1, §21.2.2). An SDK extension.
 *
 * @public
 */
export enum ChartType {
  columnClustered = 'ColumnClustered',
  columnStacked = 'ColumnStacked',
  columnStacked100 = 'ColumnStacked100',
  barClustered = 'BarClustered',
  barStacked = 'BarStacked',
  barStacked100 = 'BarStacked100',
  line = 'Line',
  lineMarkers = 'LineMarkers',
  area = 'Area',
  areaStacked = 'AreaStacked',
  areaStacked100 = 'AreaStacked100',
  pie = 'Pie',
  doughnut = 'Doughnut',
  radar = 'Radar',
  xyscatter = 'XYScatter',
  xyscatterLines = 'XYScatterLines',
}

/**
 * Where the workbook of a chart is (`c:externalData`, ECMA-376 Part 1, §21.2.2.63), and what
 * `Chart.setData` does to it. An SDK extension.
 *
 * - `Embedded`: the chart has an embedded workbook, and `setData` writes the same values into
 *   it.
 * - `Kept`: the embedded workbook has a structure that the engine does not change, so it keeps
 *   its old values.
 * - `External`: the workbook is a linked file, which the engine never opens. `setData` turns off
 *   its automatic update (`c:autoUpdate`), so the old values do not come back.
 * - `None`: the chart has no workbook. Its data is in the chart part only.
 *
 * @public
 */
export enum ChartWorkbookState {
  embedded = 'Embedded',
  kept = 'Kept',
  external = 'External',
  none = 'None',
}

/**
 * Why the engine cannot edit the data of a chart. An SDK extension.
 *
 * - `NoData`: the chart plots no series.
 * - `Combination`: the chart has more than one plot, such as columns and a line.
 * - `UnsupportedType`: a plot or a data reference that the engine does not edit, such as a
 *   stock, bubble, or surface chart.
 * - `Filtered`: the chart hides some of its series.
 * - `MultiLevel`: the categories have more than one level (`c:multiLvlStrRef`).
 * - `MixedCategories`: the series use different categories.
 * - `TooLarge`: the chart has more than 255 series or 4000 categories.
 * - `Unreadable`: the chart part is missing, or it is not a chart part.
 *
 * @public
 */
export enum ChartReadOnlyReason {
  noData = 'NoData',
  combination = 'Combination',
  unsupportedType = 'UnsupportedType',
  filtered = 'Filtered',
  multiLevel = 'MultiLevel',
  mixedCategories = 'MixedCategories',
  tooLarge = 'TooLarge',
  unreadable = 'Unreadable',
}

/**
 * One series of a chart: its name (`c:tx`) and one value for each category (`c:val`, or `c:yVal`
 * of a scatter chart). `null` is a blank point. An SDK extension.
 *
 * @public
 */
export interface ChartSeriesData {
  /** The series name, as the legend shows it. At most 255 characters. */
  name: string;
  /** One value for each category, in category order. `null` is a blank point. */
  values: (number | null)[];
}

/**
 * The data of a chart: the category labels (`c:cat`), or the X values of a scatter chart
 * (`c:xVal`) as text, and the series in plot order. An SDK extension.
 *
 * @public
 */
export interface ChartData {
  /**
   * The category labels, or the X values of a scatter chart as text. At most 4000. The chart
   * XML of the data must also fit in 16 MiB. That limit depends on the length of the labels and
   * values.
   */
  categories: string[];
  /** The series in plot order. At most 255. */
  series: ChartSeriesData[];
}

/** Null-prototype maps: a type from a consumer never selects an inherited key. */
function pairs<V extends string>(entries: readonly (readonly [string, V])[]): Record<string, V> {
  return Object.assign(Object.create(null) as Record<string, V>, Object.fromEntries(entries));
}

/** The engine's chart type (`ChartType` of the core contract) of each Office.js value. */
const TYPE_TO_FILE = pairs<ChartTypeOfFile>([
  ['ColumnClustered', 'columnClustered'],
  ['ColumnStacked', 'columnStacked'],
  ['ColumnStacked100', 'columnPercentStacked'],
  ['BarClustered', 'barClustered'],
  ['BarStacked', 'barStacked'],
  ['BarStacked100', 'barPercentStacked'],
  ['Line', 'line'],
  ['LineMarkers', 'lineMarkers'],
  ['Area', 'area'],
  ['AreaStacked', 'areaStacked'],
  ['AreaStacked100', 'areaPercentStacked'],
  ['Pie', 'pie'],
  ['Doughnut', 'doughnut'],
  ['Radar', 'radar'],
  ['XYScatter', 'scatter'],
  ['XYScatterLines', 'scatterLines'],
] as const);
const TYPE_FROM_FILE = pairs(
  Object.entries(TYPE_TO_FILE).map(([api, file]) => [file, api as `${ChartType}`] as const)
);

/** The engine's chart type of an Office.js `ChartType` value, or `InvalidArgument`. */
export function chartTypeToFile(value: unknown, target: string): ChartTypeOfFile {
  const names = Object.keys(TYPE_TO_FILE);
  if (typeof value !== 'string' || !(value in TYPE_TO_FILE))
    failArgument(target, oneOf(names), nearest(value, names));
  return TYPE_TO_FILE[value]!;
}

const WORKBOOK: Readonly<Record<AutomationChartRead['workbook'], `${ChartWorkbookState}`>> = {
  embedded: 'Embedded',
  kept: 'Kept',
  external: 'External',
  none: 'None',
};

/** The Office.js value of a read-only reason of the protocol (`combination`). */
const reasonOf = (reason: string): `${ChartReadOnlyReason}` =>
  (reason.charAt(0).toUpperCase() + reason.slice(1)) as `${ChartReadOnlyReason}`;

/**
 * The data of a chart, checked and copied, as the protocol sends it. A value that is not of the
 * form refuses with `InvalidArgument` naming `target`. The engine checks the rest at the sync:
 * the label and name text, the counts, and the X values of a scatter chart.
 */
export function chartDataOf(data: unknown, target: string): AutomationChartData {
  const form = '{ categories: string[], series: [{ name: string, values: (number | null)[] }] }';
  if (typeof data !== 'object' || data === null) failArgument(target, form);
  const { categories, series } = data as Partial<ChartData>;
  if (!Array.isArray(categories) || !Array.isArray(series)) failArgument(target, form);
  // By index, so a hole in a sparse array refuses as a value that is not a string.
  const labels: string[] = [];
  for (let i = 0; i < categories.length; i += 1) {
    const label: unknown = categories[i];
    if (typeof label !== 'string') failArgument(`${target}.categories[${i}]`, 'a string');
    labels.push(label);
  }
  const rows: AutomationChartData['series'][number][] = [];
  for (let i = 0; i < series.length; i += 1) {
    const at = `${target}.series[${i}]`;
    const s: unknown = series[i];
    if (typeof s !== 'object' || s === null) failArgument(at, '{ name: string, values: [] }');
    const { name, values } = s as Partial<ChartSeriesData>;
    if (typeof name !== 'string') failArgument(`${at}.name`, 'a string');
    if (!Array.isArray(values) || values.length !== labels.length)
      failArgument(`${at}.values`, `an array of ${labels.length} values, one for each category`);
    const copy: (number | null)[] = [];
    for (let k = 0; k < values.length; k += 1) {
      const v: unknown = values[k];
      if (v !== null && (typeof v !== 'number' || !Number.isFinite(v)))
        failArgument(`${at}.values[${k}]`, 'a finite number or null');
      copy.push(v as number | null);
    }
    rows.push({ name, values: copy });
  }
  return { categories: labels, series: rows };
}

const CHART_NAMES = [
  'chartType',
  'title',
  'categories',
  'series',
  'workbook',
  'readOnlyReason',
] as const;

/**
 * The chart of a chart frame (`p:graphicFrame` with a `c:chart` reference). An SDK extension:
 * the pinned Office.js reference has no chart object.
 *
 * Read `Shape.chart`. It is a null object for a shape that is not a chart frame, so check
 * `isNullObject` after a sync.
 *
 * @example
 * ```ts
 * const chart = shape.chart.load('categories,series');
 * await context.sync();
 * if (!chart.isNullObject) console.log(chart.categories, chart.series);
 * ```
 *
 * @public
 */
export class Chart extends ModelObject {
  readonly #shape: Shape;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, shape: Shape) {
    super(context, path, { nullable: true });
    this.#shape = shape;
  }

  /**
   * The chart type, or null for a type that has no `ChartType` value, such as a 3-D, stock,
   * bubble, or combination chart, or a stacked line chart. A chart whose type reads null can
   * still take new data when `readOnlyReason` is null.
   */
  get chartType(): ChartType | `${ChartType}` | null {
    return this.loadedProperty('chartType');
  }
  /**
   * The title as it shows (`c:title`, ECMA-376 Part 1, §21.2.2.210), or null for no title. An
   * automatic title of a chart with one series reads as the series name. For a chart whose
   * `readOnlyReason` is not null, an automatic title reads as `Chart Title`, and a chart with no
   * `c:title` reads null.
   */
  get title(): string | null {
    return this.loadedProperty('title');
  }
  /**
   * The category labels (`c:cat`, §21.2.2.24), or the X values of a scatter chart (`c:xVal`,
   * §21.2.2.234) as text. The values come from the caches of the chart part (`c:strCache`,
   * §21.2.2.199; `c:numCache`, §21.2.2.120) as the cache holds them: a number or a date category
   * reads as its cached value, such as `44927`, without its number format. Empty when
   * `readOnlyReason` is not null, except for a `Combination` chart whose plots share one set of
   * categories: it reads the data of every plot.
   */
  get categories(): string[] {
    return this.loadedProperty('categories');
  }
  /**
   * The series in plot order (`c:ser`, `c:order`): each name (`c:tx`, §21.2.2.215) and one value
   * for each category (`c:val`, §21.2.2.224). A blank point reads as `null`. Empty when
   * `categories` is empty.
   */
  get series(): ChartSeriesData[] {
    return this.loadedProperty('series');
  }
  /** Where the chart's workbook is, and what `setData` does to it. */
  get workbook(): ChartWorkbookState | `${ChartWorkbookState}` {
    return this.loadedProperty('workbook');
  }
  /** Why the engine cannot edit the chart's data, or null when `setData` can edit it. */
  get readOnlyReason(): ChartReadOnlyReason | `${ChartReadOnlyReason}` | null {
    return this.loadedProperty('readOnlyReason');
  }

  /** The chart frame. */
  getShape(): Shape {
    return this.#shape;
  }

  /**
   * Replace the categories and the series of the chart, by position. Category `i` and series
   * `j` keep the formatting, the data point overrides, and the data labels of the ones at the
   * same position. A series past the old count is new, and it takes the automatic look of its
   * position. The write changes the caches (`c:numCache`, `c:strCache`), the point counts, and
   * the cell references (`c:f`) of each series together. An embedded workbook gets the same
   * values when `workbook` is `Embedded`.
   *
   * Give at least one category and one series, at most 4000 categories and 255 series, and one
   * value for each category. Labels and names have at most 255 characters. The X values of a
   * scatter chart are numbers as text, unless the chart already names them as text.
   *
   * Data that is not of this form refuses with `InvalidArgument`. A chart whose data the
   * engine cannot edit (`readOnlyReason`) refuses with `NotSupported`, and so does a change
   * that its cell references or its embedded workbook cannot take, such as two series with the
   * same name over a workbook table. Nothing is written then. The write is one undo step.
   *
   * @example
   * ```ts
   * shape.chart.setData({
   *   categories: ['Q1', 'Q2', 'Q3', 'Q4'],
   *   series: [{ name: 'Revenue', values: [12, 15, 14, 18] }],
   * });
   * await context.sync();
   * ```
   */
  setData(data: ChartData): void {
    const checked = chartDataOf(data, this.at('setData'));
    // On the shape itself: the write waits for no read, so it goes with the other writes of the
    // sync, all or nothing. The host refuses a shape that is not a chart (`not-chart`).
    const shape = pathOf(this.#shape);
    this.command('setData', (ref) => ({ op: 'setChartData', shape: ref(shape), data: checked }));
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, CHART_NAMES);
    this.loadScalars(
      names,
      'chart',
      // On the shape, answering null for another shape: the load waits for no read.
      (ref) => ({ op: 'readChart', shape: ref(pathOf(this.#shape)), orNull: true }),
      (v) => ({
        chartType: v.chart.type ? (TYPE_FROM_FILE[v.chart.type] ?? null) : null,
        title: v.chart.title,
        categories: [...v.chart.categories],
        series: v.chart.series.map((s) => ({ name: s.name, values: [...s.values] })),
        workbook: WORKBOOK[v.chart.workbook],
        readOnlyReason: v.chart.readOnlyReason ? reasonOf(v.chart.readOnlyReason) : null,
      })
    );
  }
}
