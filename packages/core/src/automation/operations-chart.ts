// Chart data of the automation operations `readChart` and `setChartData` (see `operations.ts`).
/**
 * One series of a chart: its name (`c:tx`, ECMA-376 Part 1, §21.2.2.215) and one value for each
 * category (`c:val`, §21.2.2.224, or `c:yVal`, §21.2.2.237). `null` is a blank point.
 */
export interface AutomationChartSeries {
  readonly name: string;
  readonly values: readonly (number | null)[];
}

/**
 * The data of a chart: the category labels (`c:cat`, §21.2.2.24), or the X values of a scatter
 * chart (`c:xVal`, §21.2.2.234) as text, and the series in plot order.
 */
export interface AutomationChartData {
  readonly categories: readonly string[];
  readonly series: readonly AutomationChartSeries[];
}
