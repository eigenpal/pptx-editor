// The drawing renderer seam: a chart frame keeps its placeholder until a chart renderer is
// installed, and a renderer's elements lay out in the frame's box.

import { describe, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { resolveSlide, type ResolvedGraphicFrame } from '../../cascade/index.ts';
import { readPptx } from '../../store/index.ts';
import {
  createFixedMeasurer,
  drawingRenderers,
  installDrawingRenderers,
  layoutSlide,
  type ChartRenderer,
  type LayoutItem,
} from '../index.ts';

const deck = new Uint8Array(
  readFileSync(join(import.meta.dir, '../../../../../fixtures/charts.pptx'))
);

function firstSlide() {
  const read = readPptx(deck);
  if (!read.ok) throw new Error('charts.pptx does not open');
  return resolveSlide(read.pkg, 0)!;
}

const frameOf = (items: readonly LayoutItem[]) =>
  items.find((i): i is Extract<LayoutItem, { kind: 'frame' }> => i.kind === 'frame')!;

describe('drawing renderers', () => {
  test('a chart without the drawing module keeps its placeholder', () => {
    const slide = firstSlide();
    const frame = slide.elements.find((e) => e.kind === 'graphicFrame') as ResolvedGraphicFrame;
    expect(frame.frameKind).toBe('chart');
    expect(frame.chart?.partName).toBe('/ppt/charts/chart1.xml');
    expect(frame.chart?.chartSpace.local).toBe('chartSpace');
    const layout = layoutSlide(slide, createFixedMeasurer());
    expect(frameOf(layout.items).chart).toBeUndefined();
    expect(layout.diagnostics).toContain('graphic frame rendered as placeholder: chart');
  });

  test('an installed chart renderer lays its elements out in the frame box', () => {
    const slide = firstSlide();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    const renderer: ChartRenderer = ({ frame }) => [
      {
        kind: 'shape',
        nodeId: `${frame.nodeId}#bar`,
        partName: frame.partName,
        layer: frame.layer,
        name: '',
        hidden: false,
        xfrm: { x: 100, y: 200, cx: 300, cy: 400, rot: 0, flipH: false, flipV: false },
        geometry: { kind: 'preset', preset: 'rect', adjust: new Map() },
        fill: { type: 'solid', color: { r: 255, g: 0, b: 0, a: 1 } },
      },
    ];
    installDrawingRenderers({ chart: renderer });
    // The first layout above read the registry, so this install is late: one warning.
    expect(warn).toHaveBeenCalledTimes(1);
    installDrawingRenderers({ chart: renderer });
    expect(warn).toHaveBeenCalledTimes(1);
    warn.mockRestore();
    const layout = layoutSlide(slide, createFixedMeasurer());
    const frame = frameOf(layout.items);
    expect(frame.chart).toHaveLength(1);
    const bar = frame.chart![0]!;
    // The bar's box sits at its offset inside the frame.
    const x = frame.element.xfrm.x;
    expect(bar.matrix[4]).toBe(x + 100);
    expect(layout.diagnostics).not.toContain('graphic frame rendered as placeholder: chart');
  });

  test('a renderer that throws or declines leaves the placeholder', () => {
    installDrawingRenderers({
      chart: () => {
        throw new Error('bad chart');
      },
    });
    let layout = layoutSlide(firstSlide(), createFixedMeasurer());
    expect(frameOf(layout.items).chart).toBeUndefined();
    installDrawingRenderers({ chart: () => null });
    layout = layoutSlide(firstSlide(), createFixedMeasurer());
    expect(layout.diagnostics).toContain('graphic frame rendered as placeholder: chart');
    expect(typeof drawingRenderers().chart).toBe('function');
  });
});
