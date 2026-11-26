// Charts: the chart part (`c:chartSpace`, `CT_ChartSpace`, ECMA-376 Part 1, §21.2.2.29) that a
// chart frame names (`c:chart`, `CT_RelId`, §21.2.2.26, in a `p:graphicFrame`,
// `CT_GraphicalObjectFrame`, §19.3.1.21). The read and the data edit are the editor's own
// (`commands/chart-data.ts`, `commands/chart-format.ts`): an edit writes the series caches
// (`c:numCache`, `CT_NumData`, §21.2.2.120; `c:strCache`, `CT_StrData`, §21.2.2.199), the cell
// references (`c:f`, §21.2.2.65), and an embedded workbook (`c:externalData`,
// `CT_ExternalData`, §21.2.2.63) in one transaction, so the chart and its workbook keep the
// same values.

import type { ChartData } from '../contracts/neutral.ts';
import { attr, elements, NS, type Element } from '../store/index.ts';
import {
  chartFrameOf,
  readChartData,
  setChartData as writeChartData,
} from './commands/chart-data.ts';
import { MAX_CHART_SERIES, readChartModel } from './commands/chart-data-model.ts';
import { chartTitleText, readChartFormat } from './commands/chart-format.ts';
import { checkChartData } from './commands/chart-insert.ts';
import { chartTypeOf } from './commands/chart-type.ts';
import { cChild } from './commands/chart-xml.ts';
import { chartWorkbook } from './commands/chart-workbook.ts';
import { CommandRefusal } from './commands/op-batch.ts';
import type { AutomationChartData } from './operations-chart.ts';
import type {
  AutomationChartRead,
  AutomationChartReadOnlyReason,
  AutomationRef,
  AutomationValue,
} from './protocol.ts';
import { shapeType } from './plan-reads.ts';
import {
  AutomationFailure,
  badArgument,
  shapeAt,
  type PlanContext,
  type ShapeLocation,
} from './targets.ts';

const READ_ONLY_REASONS: ReadonlySet<string> = new Set<AutomationChartReadOnlyReason>([
  'noData',
  'combination',
  'unsupportedType',
  'filtered',
  'multiLevel',
  'mixedCategories',
  'tooLarge',
]);

/** The read-only reason of an engine refusal key (`errors.chartData.combination`). */
function readOnlyReasonOf(key: string): AutomationChartReadOnlyReason {
  const name = key.startsWith('errors.chartData.') ? key.slice('errors.chartData.'.length) : '';
  return READ_ONLY_REASONS.has(name) ? (name as AutomationChartReadOnlyReason) : 'unsupportedType';
}

/** A chart frame (`a:graphicData/@uri` of the chart namespace), or a `not-chart` failure. */
function chartShapeAt(ctx: PlanContext, ref: AutomationRef): ShapeLocation {
  const at = shapeAt(ctx, ref);
  if (shapeType(at.el) !== 'chart')
    throw new AutomationFailure('unsupported-content', 'that shape is not a chart', 'not-chart');
  return at;
}

/**
 * The type of a chart whose data the editor does not edit: the type of its one plot, or null
 * for a chart with no plot or more than one (`c:plotArea`, §21.2.2.145).
 */
function readOnlyType(chartSpace: Element): AutomationChartRead['type'] {
  const plotArea = cChild(cChild(chartSpace, 'chart'), 'plotArea');
  const plots = elements(plotArea).filter((e) => e.ns === NS.c && e.local.endsWith('Chart'));
  return plots.length === 1 ? chartTypeOf(plots[0]!) : null;
}

/** The most plots of a combination chart that the read combines. */
const MAX_COMBINED_PLOTS = 8;

/** A number of a cache point, or null for a blank or non-numeric point. */
function pointValue(v: string | undefined): number | null {
  const n = v === undefined || v.trim() === '' ? NaN : Number(v);
  return Number.isFinite(n) ? n : null;
}

/**
 * The data of a combination chart, which the editor does not edit: each plot of the plot area
 * read on its own (`c:plotArea`, `CT_PlotArea`, §21.2.2.145), when every plot has category data and all of them
 * share one set of categories. The series follow `c:order` across the plots. Undefined for any
 * other chart.
 */
function combinationData(
  chartSpace: Element
): Pick<AutomationChartRead, 'categories' | 'series'> | undefined {
  const chart = cChild(chartSpace, 'chart');
  const plotArea = cChild(chart, 'plotArea');
  if (!chart || !plotArea) return undefined;
  const isPlot = (e: Element): boolean => e.ns === NS.c && e.local.endsWith('Chart');
  const plots = elements(plotArea).filter(isPlot);
  // A combination has a few plots; a part with more is not read, so its cost stays bounded.
  if (plots.length < 2 || plots.length > MAX_COMBINED_PLOTS) return undefined;
  // The categories of the plots that name some; a plot without `c:cat` takes them.
  let categories: readonly string[] | undefined;
  let count = 0;
  const series: { order: number; name: string; points: readonly (string | undefined)[] }[] = [];
  for (const plot of plots) {
    // The chart part with this plot alone in its plot area: the read of a one-plot chart.
    const only: Element = {
      ...plotArea,
      children: plotArea.children.filter((k) => k.type !== 'element' || !isPlot(k) || k === plot),
    };
    const alone: Element = {
      ...chartSpace,
      children: chartSpace.children.map((k) =>
        k === chart
          ? { ...chart, children: chart.children.map((g) => (g === plotArea ? only : g)) }
          : k
      ),
    };
    const read = readChartModel(alone);
    if (!read.ok || read.model.kind !== 'category') return undefined;
    const labels = read.model.categories;
    if (read.model.hasCategories) {
      if (categories && categories.join('\u0000') !== labels.join('\u0000')) return undefined;
      categories = labels;
    }
    count = Math.max(count, labels.length);
    if (series.length + read.model.series.length > MAX_CHART_SERIES) return undefined;
    for (const s of read.model.series) {
      const order = Number(attr(cChild(s.el, 'order'), 'val'));
      series.push({
        order: Number.isFinite(order) ? order : series.length,
        name: s.name,
        points: s.val?.values ?? [],
      });
    }
  }
  const labels = categories ?? Array.from({ length: count }, (_, i) => String(i + 1));
  // A plot with more points than the shared categories has data that no category names.
  if (count > labels.length) return undefined;
  series.sort((a, b) => a.order - b.order);
  return {
    categories: [...labels],
    series: series.map(({ name, points }) => ({
      name,
      values: labels.map((_, i) => pointValue(points[i])),
    })),
  };
}

/**
 * The chart of a chart frame. `orNull`: a shape that is not a chart answers null, so the object
 * model loads a chart and tells a null object in one read.
 */
export function readChart(ctx: PlanContext, ref: AutomationRef, orNull?: boolean): AutomationValue {
  if (orNull && shapeType(shapeAt(ctx, ref).el) !== 'chart') return { kind: 'null' };
  const at = chartShapeAt(ctx, ref);
  const pkg = ctx.batch.pkg;
  const frame = chartFrameOf(pkg, at.part, at.el.id);
  if (!frame)
    return {
      kind: 'chart',
      chart: {
        type: null,
        title: null,
        categories: [],
        series: [],
        workbook: 'none',
        readOnlyReason: 'unreadable',
      },
    };
  const data = readChartData(pkg, at.part, at.el.id);
  const format = readChartFormat(pkg, at.part, at.el.id);
  if (!data.ok || !format.ok) {
    const model = readChartModel(frame.chartSpace);
    const combined =
      !model.ok && model.reason === 'errors.chartData.combination'
        ? combinationData(frame.chartSpace)
        : undefined;
    return {
      kind: 'chart',
      chart: {
        type: readOnlyType(frame.chartSpace),
        title: chartTitleText(frame.chartSpace),
        categories: combined?.categories ?? [],
        series: combined?.series ?? [],
        workbook: chartWorkbook(pkg, frame.part, frame.chartSpace).state,
        readOnlyReason: model.ok ? 'unsupportedType' : readOnlyReasonOf(model.reason),
      },
    };
  }
  return {
    kind: 'chart',
    chart: {
      type: format.format.type,
      title: format.format.title,
      categories: data.data.categories.map((c) => c.label),
      series: data.data.series.map((s) => ({ name: s.name, values: [...s.values] })),
      workbook: data.data.workbook ?? 'none',
      readOnlyReason: null,
    },
  };
}

/** The details of an `errors.unsupported` refusal that name the form of the data. */
const DATA_FORM = new Set(['data', 'series', 'values', 'source', 'kind']);

/** Engine refusals that name the caller's data, not the chart. */
const DATA_REFUSALS = new Set([
  'errors.chartData.empty',
  'errors.chartData.tooLarge',
  'errors.chartData.invalidText',
  'errors.chartData.notNumber',
]);

/**
 * Run a chart edit. A refusal of the data is `invalid-argument`; a chart whose data, cell
 * references, or workbook the edit cannot write is `unsupported-content`. The detail is the
 * engine's reason, which selects the public diagnostic.
 */
export function chartEdit<T>(run: () => T): T {
  try {
    return run();
  } catch (e) {
    if (!(e instanceof CommandRefusal)) throw e;
    // The form of the data: an array, a series, a value count, a source index.
    if (e.reason === 'errors.unsupported' && DATA_FORM.has(e.detail ?? ''))
      badArgument('chartData');
    const code = DATA_REFUSALS.has(e.reason) ? 'invalid-argument' : 'unsupported-content';
    // A store refusal keeps its detail; the diagnostic reads the reason before the ':'.
    const detail =
      e.reason === 'errors.unsupported' && e.detail ? `${e.reason}:${e.detail}` : e.reason;
    throw new AutomationFailure(code, 'the chart data was refused', detail);
  }
}

/**
 * Replace the categories and series of a chart frame, by position: category `i` and series `j`
 * continue the ones at the same position, with their formatting, data point overrides, and
 * labels; the ones past the old counts are new and take automatic formatting.
 */
export function setChartData(
  ctx: PlanContext,
  ref: AutomationRef,
  data: AutomationChartData
): AutomationValue {
  const at = chartShapeAt(ctx, ref);
  const pkg = ctx.batch.pkg;
  // The form of the data first: a refusal names the data, whatever the chart is.
  const checked = chartEdit(() => checkChartData(data, false));
  const frame = chartFrameOf(pkg, at.part, at.el.id);
  if (!frame)
    throw new AutomationFailure(
      'unsupported-content',
      'the chart part is unreadable',
      'chart-unreadable'
    );
  const model = readChartModel(frame.chartSpace);
  // One detail for every read-only chart: `readChart` names the reason.
  if (!model.ok)
    throw new AutomationFailure(
      'unsupported-content',
      'the chart data is read-only',
      `chart-read-only:${model.reason}`
    );
  const old = model.model;
  const next: Omit<ChartData, 'workbook'> = {
    kind: old.kind,
    categories: checked.categories.map((label, i) =>
      i < old.categories.length ? { label, source: i } : { label }
    ),
    series: checked.series.map((s, i) =>
      i < old.series.length
        ? { name: s.name, values: s.values, source: i }
        : { name: s.name, values: s.values }
    ),
  };
  chartEdit(() => writeChartData(ctx.batch, at.part, at.el.id, next));
  return { kind: 'applied' };
}
