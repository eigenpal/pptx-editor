// The Chart Data tab of the side panel, headless: the text of each field, the check of what the
// user types, and the one `setChartData` command that each committed field, or each add or
// remove, sends. The React and Vue panes (`ChartDataPane.tsx`) render from these, so both check
// the same input with the same keys.

import type { ChartData, EditorCommand, EditorSnapshot } from '../contracts/index.ts';
import {
  MAX_CHART_LABEL,
  MAX_CHART_POINTS,
  MAX_CHART_SERIES,
} from '../automation/commands/chart-data-model.ts';

export { MAX_CHART_POINTS, MAX_CHART_SERIES };

/** The chart identity and grid labels used to guard local field drafts. @internal */
export function chartPaneDraftKey(snapshot: EditorSnapshot, data: ChartData): string {
  return JSON.stringify([
    snapshot.currentSlide,
    snapshot.chart?.nodeId,
    data.kind,
    data.categories.map((category) => category.label),
    data.series.map((series) => series.name),
  ]);
}

/**
 * Keep drafts when the chart and existing grid labels stay in place, including appends.
 * Source indexes are renumbered on read, so other topology changes clear drafts. @internal
 */
export function chartPaneDraftContinues(previous: string, next: string): boolean {
  if (previous === next) return true;
  type Scope = [number, string | null, string, string[], string[]];
  const before = JSON.parse(previous) as Scope;
  const after = JSON.parse(next) as Scope;
  return (
    before[0] === after[0] &&
    before[1] === after[1] &&
    before[2] === after[2] &&
    before[3].length <= after[3].length &&
    before[4].length <= after[4].length &&
    before[3].every((label, index) => label === after[3][index]) &&
    before[4].every((name, index) => name === after[4][index])
  );
}

/** One change of the Chart Data tab: a committed field, or an add or a remove. @public */
export type ChartPaneChange =
  | {
      readonly type: 'value';
      readonly series: number;
      readonly category: number;
      readonly text: string;
    }
  | { readonly type: 'seriesName'; readonly series: number; readonly text: string }
  | { readonly type: 'label'; readonly category: number; readonly text: string }
  | { readonly type: 'addSeries'; readonly name: string }
  | { readonly type: 'removeSeries'; readonly series: number }
  | { readonly type: 'addCategory'; readonly label: string }
  | { readonly type: 'removeCategory'; readonly category: number };

/** The command of a change, or the i18n key of why the typed text is refused. @public */
export type ChartPaneResult =
  | { readonly ok: true; readonly command: EditorCommand }
  | {
      readonly ok: false;
      readonly error:
        | 'chartData.errors.number'
        | 'chartData.errors.xNumber'
        | 'chartData.errors.tooLong'
        | 'chartData.errors.limit';
    };

/** A value as its field shows it: empty for no value. @public */
export const chartValueText = (v: number | null | undefined): string =>
  v === null || v === undefined ? '' : String(v);

/** The number a value field holds, null when it is empty, or undefined for other text. @public */
export function parseChartValue(text: string): number | null | undefined {
  const t = text.trim().replace(',', '.');
  if (t === '') return null;
  const n = Number(t);
  return Number.isFinite(n) ? n : undefined;
}

/**
 * Whether the categories of `data` are numbers that the tab checks: the X values of a scatter
 * chart, unless the chart names them as text. @public
 */
export function chartPaneNumericX(data: Pick<ChartData, 'kind' | 'categories'>): boolean {
  return (
    data.kind === 'xy' && data.categories.every((c) => typeof parseChartValue(c.label) === 'number')
  );
}

/**
 * The label of a new category: the next X value of a scatter chart with number X values, else
 * `fallback`. @public
 */
export function nextChartCategoryLabel(
  data: Pick<ChartData, 'kind' | 'categories'>,
  fallback: string
): string {
  if (!chartPaneNumericX(data)) return fallback;
  let max = 0;
  for (const c of data.categories) max = Math.max(max, Number(parseChartValue(c.label)));
  return String(Math.floor(max) + 1);
}

const fail = (error: Extract<ChartPaneResult, { ok: false }>['error']): ChartPaneResult => ({
  ok: false,
  error,
});

/**
 * The `setChartData` command of one change to the chart's current data, or why the typed text
 * is refused. Every series and category keeps its `source`, so formatting stays with it. A
 * value that is the same as before still gives a command; the pane skips unchanged fields. @public
 */
export function chartPaneCommand(data: ChartData, change: ChartPaneChange): ChartPaneResult {
  const numericX = chartPaneNumericX(data);
  let categories: { label: string; source?: number }[] = data.categories.map((c) => ({ ...c }));
  let series: { name: string; values: (number | null)[]; source?: number }[] = data.series.map(
    (s) => ({ ...s, values: [...s.values] })
  );
  switch (change.type) {
    case 'value': {
      const v = parseChartValue(change.text);
      const target = series[change.series];
      if (v === undefined || !target || change.category >= categories.length)
        return fail('chartData.errors.number');
      target.values[change.category] = v;
      break;
    }
    case 'seriesName': {
      const target = series[change.series];
      if (!target) return fail('chartData.errors.tooLong');
      if (change.text.length > MAX_CHART_LABEL) return fail('chartData.errors.tooLong');
      target.name = change.text;
      break;
    }
    case 'label': {
      const target = categories[change.category];
      if (!target || change.text.length > MAX_CHART_LABEL) return fail('chartData.errors.tooLong');
      const x = parseChartValue(change.text);
      if (numericX && typeof x !== 'number') return fail('chartData.errors.xNumber');
      target.label = numericX ? String(x) : change.text;
      break;
    }
    case 'addSeries':
      if (series.length >= MAX_CHART_SERIES) return fail('chartData.errors.limit');
      series = [...series, { name: change.name, values: categories.map(() => null) }];
      break;
    case 'removeSeries':
      if (series.length <= 1 || !series[change.series]) return fail('chartData.errors.limit');
      series = series.filter((_, i) => i !== change.series);
      break;
    case 'addCategory':
      if (categories.length >= MAX_CHART_POINTS) return fail('chartData.errors.limit');
      categories = [...categories, { label: change.label }];
      series = series.map((s) => ({ ...s, values: [...s.values, null] }));
      break;
    case 'removeCategory':
      if (categories.length <= 1 || !categories[change.category])
        return fail('chartData.errors.limit');
      categories = categories.filter((_, i) => i !== change.category);
      series = series.map((s) => ({
        ...s,
        values: s.values.filter((_, i) => i !== change.category),
      }));
      break;
  }
  return {
    ok: true,
    command: { type: 'setChartData', data: { kind: data.kind, categories, series } },
  };
}
