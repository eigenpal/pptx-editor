/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Shared helpers for the chart tests: the chart frames of `fixtures/charts.pptx` (or another
// chart fixture), and a chart laid out with the deterministic fixed measurer.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  resolveSlide,
  type ResolvedChartPart,
  type ResolvedGraphicFrame,
} from '@pptx-editor.dev/core/cascade';
import { createFixedMeasurer } from '@pptx-editor.dev/core/layout';
import { parseXmlPart, readPptx, type Element } from '@pptx-editor.dev/core/store';
import { readChartModel, type ChartModel } from '../chart-model.ts';
import { renderChart, type ChartRender } from '../chart-render.ts';
import { withChartText } from '../chart-text.ts';

export const FIXTURES = join(import.meta.dir, '../../../../../fixtures');

export function chartsDeck(deck = 'charts'): Uint8Array {
  return new Uint8Array(readFileSync(join(FIXTURES, `${deck}.pptx`)));
}

export interface FixtureChart {
  readonly frame: ResolvedGraphicFrame;
  readonly chart: ResolvedChartPart;
}

/** The chart frame of each slide of `<deck>.pptx`, in slide order. */
export function fixtureCharts(deck = 'charts', slides = 8): FixtureChart[] {
  const read = readPptx(chartsDeck(deck));
  if (!read.ok) throw new Error(`${deck}.pptx does not open`);
  const out: FixtureChart[] = [];
  for (let index = 0; index < slides; index += 1) {
    const slide = resolveSlide(read.pkg, index)!;
    const frame = slide.elements.find(
      (e): e is ResolvedGraphicFrame => e.kind === 'graphicFrame' && e.frameKind === 'chart'
    );
    if (!frame?.chart) throw new Error(`slide ${index + 1} has no chart`);
    out.push({ frame, chart: frame.chart });
  }
  return out;
}

/** A chart of the fixture laid out with the fixed measurer, at its frame's size. */
export function renderFixture(
  index: number,
  charts: readonly FixtureChart[] = fixtureCharts()
): { model: ChartModel; rendered: ChartRender } {
  const { frame, chart } = charts[index]!;
  const model = readChartModel(chart.chartSpace)!;
  const rendered = withChartText({ measurer: createFixedMeasurer(), theme: chart.theme }, () =>
    renderChart(model, frame.xfrm.cx / 12_700, frame.xfrm.cy / 12_700, chart.colors)
  );
  if (!rendered) throw new Error(`chart ${index + 1} does not render`);
  return { model, rendered };
}

/** A `c:chartSpace` from markup, for tests of hostile or unusual charts. */
export function chartSpace(plotArea: string, chartExtra = ''): Element {
  const xml =
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    `<c:chart>${chartExtra}<c:plotArea><c:layout/>${plotArea}</c:plotArea></c:chart></c:chartSpace>`;
  const parsed = parseXmlPart('/ppt/charts/test.xml', xml);
  if (!parsed.ok) throw new Error(parsed.reason);
  return parsed.part.root;
}

/** A numeric series with `count` cached points. */
export function numericSeries(values: readonly number[], index = 0): string {
  const points = values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('');
  return (
    `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>` +
    `<c:val><c:numRef><c:numCache><c:ptCount val="${values.length}"/>${points}` +
    '</c:numCache></c:numRef></c:val></c:ser>'
  );
}
