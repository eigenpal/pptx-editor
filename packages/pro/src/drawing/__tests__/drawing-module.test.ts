/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { describe, expect, test } from 'bun:test';
import { resolveSlide } from '@pptx-editor.dev/core/cascade';
import { createPptxEditor } from '@pptx-editor.dev/core/editor';
import { createFixedMeasurer, layoutSlide, type LayoutItem } from '@pptx-editor.dev/core/layout';
import { readPptx } from '@pptx-editor.dev/core/store';
import { unzipSync } from 'fflate';
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

  test('a resized chart lays out again at its new size, and save keeps the chart part', async () => {
    installDrawing();
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const deck = chartsDeck();
    expect((await editor.load(deck)).ok).toBe(true);
    // The first shape is the chart area, as wide as the frame.
    const area = (items: readonly LayoutItem[]) => items[0]!.element.xfrm.cx;
    const before = frameOf(editor.session!.layout(5)!.items);
    const { nodeId, xfrm } = before.element;
    expect(area(before.chart!)).toBe(xfrm.cx);
    const half = Math.round(xfrm.cx / 2);
    editor.goToSlide(5);
    editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] });
    const moved = {
      type: 'setTransform',
      nodeId,
      x: xfrm.x,
      y: xfrm.y,
      cy: xfrm.cy,
      rot: 0,
    } as const;
    expect(editor.exec({ ...moved, cx: half }).ok).toBe(true);
    const after = frameOf(editor.session!.layout(5)!.items);
    expect(after.element.xfrm.cx).toBe(half);
    expect(area(after.chart!)).toBe(half);
    const saved = editor.save()!;
    const name = 'ppt/charts/chart6.xml';
    expect([...unzipSync(saved)[name]!]).toEqual([...unzipSync(deck)[name]!]);
    editor.destroy?.();
  });

  test('the chart cache serves one frame again and lays out again for another theme', () => {
    const { frame, chart } = fixtureCharts()[0]!;
    const measurer = createFixedMeasurer();
    const first = renderChartFrame({ frame, chart, measurer });
    expect(renderChartFrame({ frame, chart, measurer })).toBe(first);
    const scheme = new Map(chart.colors.scheme);
    scheme.set('lt1', { r: 0, g: 0, b: 0, a: 1 });
    const dark = { ...chart, colors: { ...chart.colors, scheme } };
    const again = renderChartFrame({ frame, chart: dark, measurer });
    expect(again).not.toBe(first);
    expect(renderChartFrame({ frame, chart, measurer: createFixedMeasurer() })).not.toBe(again);
  });
});
