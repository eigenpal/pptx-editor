/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Charts that the editor inserts or changes (`insertChart`, `setChartType`, `setChartFormat`)
// draw with the drawing module, with no placeholder, for every type the editor writes.
import { describe, expect, test } from 'bun:test';
import { createPptxEditor } from '@pptx-editor.dev/core/editor';
import type { ChartType } from '@pptx-editor.dev/core/contracts';
import type { LayoutItem } from '@pptx-editor.dev/core/layout';
import { drawingModule } from '../index.ts';
import { chartsDeck } from './chart-support.ts';

type Frame = Extract<LayoutItem, { kind: 'frame' }>;

const TYPES: readonly ChartType[] = [
  'columnClustered',
  'columnStacked',
  'columnPercentStacked',
  'barClustered',
  'barStacked',
  'barPercentStacked',
  'line',
  'lineMarkers',
  'area',
  'areaStacked',
  'areaPercentStacked',
  'pie',
  'doughnut',
  'radar',
  'scatter',
  'scatterLines',
];

async function editor() {
  const e = createPptxEditor({
    fixedMeasurer: true,
    substitutes: false,
    modules: [drawingModule({ licenseKey: 'test' })],
  });
  expect((await e.load(chartsDeck())).ok).toBe(true);
  return e;
}

/** The laid-out chart of the selected frame. */
function drawn(e: Awaited<ReturnType<typeof editor>>): Frame {
  const sel = e.snapshot().selection;
  const id = sel.kind === 'shapes' ? sel.nodeIds[0] : undefined;
  const layout = e.session!.layout(e.snapshot().currentSlide)!;
  expect(layout.diagnostics).not.toContain('graphic frame rendered as placeholder: chart');
  return layout.items.find((i): i is Frame => i.kind === 'frame' && i.element.nodeId === id)!;
}

/** The text of the drawn chart's labels. */
const texts = (frame: Frame): string[] =>
  frame.chart!.filter((i) => i.kind === 'shape' && i.text).map((i) => JSON.stringify(i));

describe('inserted and changed charts', () => {
  test('every inserted type draws its series, with no placeholder', async () => {
    const e = await editor();
    for (const chartType of TYPES) {
      e.setSelection({ kind: 'none' });
      expect(e.exec({ type: 'insertChart', chartType }).ok).toBe(true);
      expect(drawn(e).chart?.length ?? 0).toBeGreaterThan(4);
    }
    e.destroy?.();
  });

  test('a column chart changed to a pie colors each slice, with any number of series', async () => {
    const e = await editor();
    expect(e.exec({ type: 'insertChart', chartType: 'columnClustered' }).ok).toBe(true);
    expect(e.exec({ type: 'setChartType', chartType: 'pie' }).ok).toBe(true);
    const fills = new Set(
      drawn(e)
        .chart!.filter((i) => i.kind === 'shape' && !i.text)
        .map((i) => JSON.stringify(i.kind === 'shape' ? i.element.fill : null))
    );
    // Four slices, each in its own color, and the chart area.
    expect(fills.size).toBeGreaterThanOrEqual(4);
    const legend = texts(drawn(e)).filter((t) => t.includes('Category'));
    expect(legend.length).toBeGreaterThanOrEqual(4);
    e.destroy?.();
  });

  test('a title, labels, and axis titles add drawn text', async () => {
    const e = await editor();
    expect(e.exec({ type: 'insertChart', chartType: 'columnClustered' }).ok).toBe(true);
    const before = texts(drawn(e));
    const format = {
      title: 'Quarterly sales',
      dataLabels: true,
      categoryAxisTitle: 'Quarter',
      valueAxisTitle: 'Units',
    };
    expect(e.exec({ type: 'setChartFormat', format }).ok).toBe(true);
    const after = texts(drawn(e));
    expect(after.length).toBeGreaterThan(before.length + 12);
    for (const word of ['Quarterly sales', 'Quarter', 'Units'])
      expect(after.some((t) => t.includes(word))).toBe(true);
    e.destroy?.();
  });
});
