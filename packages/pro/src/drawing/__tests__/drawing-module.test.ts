/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { describe, expect, test } from 'bun:test';
import { resolveSlide } from '@pptx-editor.dev/core/cascade';
import { createPptxEditor } from '@pptx-editor.dev/core/editor';
import { createFixedMeasurer, layoutSlide, type LayoutItem } from '@pptx-editor.dev/core/layout';
import { readPptx } from '@pptx-editor.dev/core/store';
import { drawingModule, installDrawing } from '../index.ts';
import { renderChartFrame } from '../chart-projection.ts';
import { chartsDeck, chartSpace, fixtureCharts, numericSeries } from './chart-support.ts';

type Frame = Extract<LayoutItem, { kind: 'frame' }>;
const frameOf = (items: readonly LayoutItem[]) =>
  items.find((i): i is Frame => i.kind === 'frame')!;

describe('drawingModule', () => {
  test('draws every fixture chart as laid-out shapes and labels in its frame', async () => {
    const editor = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      modules: [drawingModule({ licenseKey: 'test' })],
    });
    expect((await editor.load(chartsDeck())).ok).toBe(true);
    for (let index = 0; index < 8; index += 1) {
      const layout = editor.session!.layout(index)!;
      const frame = frameOf(layout.items);
      expect(frame.chart?.length ?? 0).toBeGreaterThan(5);
      expect(layout.diagnostics).not.toContain('graphic frame rendered as placeholder: chart');
      const labels = frame.chart!.filter((i) => i.kind === 'shape' && i.text);
      expect(labels.length).toBeGreaterThan(0);
      // Every item is under the frame's id, so the frame stays the one object.
      for (const item of frame.chart!)
        expect(item.element.nodeId.startsWith(`${frame.element.nodeId}#chart-`)).toBe(true);
    }
    editor.destroy?.();
  });

  test('keeps the placeholder for a chart type outside the supported set', () => {
    installDrawing();
    const { frame, chart } = fixtureCharts()[0]!;
    const threeD = chartSpace(`<c:bar3DChart>${numericSeries([1, 2])}</c:bar3DChart>`);
    const measurer = createFixedMeasurer();
    expect(renderChartFrame({ frame, chart: { ...chart, chartSpace: threeD }, measurer })).toBe(
      null
    );
    expect(renderChartFrame({ frame, chart, measurer })).not.toBe(null);
  });

  test('installDrawing serves headless layout', () => {
    installDrawing();
    const read = readPptx(chartsDeck());
    if (!read.ok) throw new Error('deck');
    const layout = layoutSlide(resolveSlide(read.pkg, 5)!, createFixedMeasurer());
    expect(frameOf(layout.items).chart?.length ?? 0).toBeGreaterThan(4);
  });
});
