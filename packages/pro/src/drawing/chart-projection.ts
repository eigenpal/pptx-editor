/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// After DOCX Editor `packages/pro/src/drawing/chart-projection.ts` at commit ba742d7d1.
// The chart renderer: a chart frame's shapes and labels from its chart part.
//
// The chart lays out from the values cached in its part, and every primitive and label lowers to
// a resolved shape that the engine lays out and both painters draw. A chart renders whole or not
// at all: a chart type outside the supported set, or one past the bounds, renders null and the
// frame keeps its labeled placeholder.

import type { ResolvedElement } from '@pptx-editor.dev/core/cascade';
import type { ChartRenderInput, TextMeasurer } from '@pptx-editor.dev/core/layout';
import { chartElements } from './chart-elements.ts';
import { readChartModel, type ChartModel } from './chart-model.ts';
import { renderChart } from './chart-render.ts';
import { withChartText } from './chart-text.ts';
import type { OoxmlElement } from './chart-xml.ts';

const EMU_PER_POINT = 12_700;

/**
 * Whether every plot of a chart is a type this renderer draws: bar, column, line, area,
 * scatter, pie, doughnut, and radar, in two dimensions, with no trendlines or error bars.
 * 3-D, of-pie, stock, bubble, and surface charts keep the placeholder.
 */
export function chartInScope(model: ChartModel): boolean {
  if (model.plots.length === 0) return false;
  for (const plot of model.plots) {
    if (plot.threeD) return false;
    if (plot.element.local === 'stockChart') return false;
    if (!['bar', 'line', 'area', 'scatter', 'pie', 'doughnut', 'radar'].includes(plot.kind))
      return false;
    for (const series of plot.series)
      if (series.trendlines.length > 0 || series.errorBars.length > 0) return false;
  }
  return true;
}

interface ChartMemo {
  readonly cx: number;
  readonly cy: number;
  readonly measurer: TextMeasurer;
  readonly owner: string;
  readonly theme: string;
  readonly result: readonly ResolvedElement[] | null;
}

/** One result per chart part tree: layout runs on every edit, and the chart rarely changes. */
const MEMO = new WeakMap<OoxmlElement, ChartMemo>();

/** What the theme contributes to a chart: its colors, its color map, and its fonts. */
function themeKey({ chart }: ChartRenderInput): string {
  const colors = [...chart.colors.scheme].map(([k, c]) => `${k}=${c.r},${c.g},${c.b},${c.a}`);
  const map = [...chart.colors.colorMap].map(([k, v]) => `${k}=${v}`);
  const fonts = [chart.theme.majorFont.latin, chart.theme.minorFont.latin];
  return [...colors, ...map, ...fonts].join('|');
}

/**
 * The shapes and labels of a chart frame in its box, or null when the chart is outside the
 * supported set or past the bounds.
 */
export function renderChartFrame(input: ChartRenderInput): readonly ResolvedElement[] | null {
  const { frame, chart, measurer } = input;
  const { cx, cy } = frame.xfrm;
  if (!(cx > 0) || !(cy > 0)) return null;
  const owner = `${frame.partName}|${frame.layer}|${frame.nodeId}`;
  const theme = themeKey(input);
  const memo = MEMO.get(chart.chartSpace);
  if (
    memo &&
    memo.cx === cx &&
    memo.cy === cy &&
    memo.measurer === measurer &&
    memo.owner === owner &&
    memo.theme === theme
  )
    return memo.result;
  const result = renderUncached(input);
  MEMO.set(chart.chartSpace, {
    cx,
    cy,
    measurer,
    owner,
    theme,
    result,
  });
  return result;
}

function renderUncached({ frame, chart, measurer }: ChartRenderInput) {
  const model = readChartModel(chart.chartSpace);
  if (!model || !chartInScope(model)) return null;
  return withChartText({ measurer, theme: chart.theme }, () => {
    const rendered = renderChart(
      model,
      frame.xfrm.cx / EMU_PER_POINT,
      frame.xfrm.cy / EMU_PER_POINT,
      chart.colors
    );
    if (!rendered) return null;
    return chartElements(rendered.scene.components, rendered.labels, frame, chart.theme);
  });
}
