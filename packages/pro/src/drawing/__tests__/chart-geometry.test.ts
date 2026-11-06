/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The geometry of each supported chart type, from the charts of `fixtures/charts.pptx`. Values
// are read back from the painted shapes, so a change to the scale, the placement, or the paint
// of one type fails here.

import { describe, expect, test } from 'bun:test';
import type { ChartComponent, Point } from '../chart-scene.ts';
import { outsideLabelRoom } from '../chart-pie.ts';
import { renderFixture } from './chart-support.ts';

interface Box {
  l: number;
  t: number;
  r: number;
  b: number;
}

function boxOf(points: readonly Point[]): Box {
  return {
    l: Math.min(...points.map((p) => p.x)),
    t: Math.min(...points.map((p) => p.y)),
    r: Math.max(...points.map((p) => p.x)),
    b: Math.max(...points.map((p) => p.y)),
  };
}

const filled = (components: readonly ChartComponent[], hex: string) =>
  components.filter((c) => c.closed && c.fill?.hex.toUpperCase() === hex);

const stroked = (components: readonly ChartComponent[], hex: string) =>
  components.filter((c) => !c.closed && c.stroke?.hex.toUpperCase() === hex);

/** The signed area of a polygon (shoelace). */
function area(points: readonly Point[]): number {
  let sum = 0;
  for (let i = 0; i < points.length; i += 1) {
    const a = points[i]!;
    const b = points[(i + 1) % points.length]!;
    sum += a.x * b.y - b.x * a.y;
  }
  return Math.abs(sum) / 2;
}

const labelTexts = (index: number) =>
  renderFixture(index).rendered.labels.map((l) =>
    l.paragraphs
      .flat()
      .map((r) => r.text)
      .join('')
  );

describe('chart geometry', () => {
  test('clustered column: bar heights follow the values from one baseline', () => {
    const { components } = renderFixture(0).rendered.scene;
    // The legend key is the small square; the bars are taller than any key.
    const plan = filled(components, '4472C4')
      .map((c) => boxOf(c.subpaths[0]!))
      .filter((bar) => bar.b - bar.t > 20);
    expect(plan).toHaveLength(4);
    const baseline = plan[0]!.b;
    for (const bar of plan) expect(bar.b).toBeCloseTo(baseline, 3);
    const heights = plan.map((bar) => bar.b - bar.t);
    [42, 35, 51, 28].forEach((v, i) => expect(heights[i]! / heights[0]!).toBeCloseTo(v / 42, 3));
    // The three series of one category sit side by side, left to right.
    const north = ['4472C4', 'ED7D31', 'A5A5A5'].map((hex) =>
      boxOf(filled(components, hex)[0]!.subpaths[0]!)
    );
    expect(north[0]!.r).toBeLessThan(north[1]!.l);
    expect(north[1]!.r).toBeLessThan(north[2]!.l);
    const texts = labelTexts(0);
    expect(texts).toContain('Regional sales');
    expect(texts).toEqual(expect.arrayContaining(['0', '60', 'North', 'West', 'Plan', 'Forecast']));
  });

  test('stacked bar: segments of one category join end to end, with value labels', () => {
    const { rendered } = renderFixture(1);
    const segments = ['4472C4', 'ED7D31', 'A5A5A5'].map((hex) =>
      boxOf(filled(rendered.scene.components, hex)[0]!.subpaths[0]!)
    );
    expect(segments[0]!.r).toBeCloseTo(segments[1]!.l, 3);
    expect(segments[1]!.r).toBeCloseTo(segments[2]!.l, 3);
    const widths = segments.map((s) => s.r - s.l);
    expect(widths[1]! / widths[0]!).toBeCloseTo(38 / 42, 3);
    expect(labelTexts(1)).toEqual(expect.arrayContaining(['42', '38', '45', '28']));
  });

  test('line with markers: one open polyline per series through every category', () => {
    const { components } = renderFixture(2).rendered.scene;
    const line = stroked(components, '4472C4').find((c) => c.subpaths[0]!.length === 4)!;
    const ys = line.subpaths[0]!.map((p) => p.y);
    // Higher values sit higher: 51 (East) is the top point, 28 (West) the lowest.
    expect(Math.min(...ys)).toBe(ys[2]!);
    expect(Math.max(...ys)).toBe(ys[3]!);
    const markers = filled(components, '4472C4').filter((c) => c.subpaths.length >= 1);
    expect(markers.length).toBeGreaterThan(0);
  });

  test('stacked area: each series is a closed band over the one below', () => {
    const { components } = renderFixture(3).rendered.scene;
    const bands = ['4472C4', 'ED7D31', 'A5A5A5'].map((hex) => filled(components, hex)[0]!);
    for (const band of bands) expect(band.subpaths[0]!.length).toBeGreaterThanOrEqual(6);
    const tops = bands.map((band) => boxOf(band.subpaths[0]!).t);
    expect(tops[2]!).toBeLessThan(tops[1]!);
    expect(tops[1]!).toBeLessThan(tops[0]!);
  });

  test('scatter: markers sit at their x values, evenly spaced for even steps', () => {
    const { model, rendered } = renderFixture(4);
    const series = model.plots[0]!.series[0]!;
    expect(series.xValues).toEqual([1, 2.5, 4, 5.5, 7]);
    const centres = filled(rendered.scene.components, '4472C4')
      .flatMap((c) => c.subpaths)
      .filter((path) => path.length >= 8)
      .map((path) => boxOf(path))
      .map((box) => (box.l + box.r) / 2)
      .sort((a, b) => a - b)
      .slice(0, 5);
    expect(centres).toHaveLength(5);
    const step = centres[1]! - centres[0]!;
    expect(step).toBeGreaterThan(0);
    for (let i = 2; i < 5; i += 1) expect(centres[i]! - centres[i - 1]!).toBeCloseTo(step, 3);
  });

  test('pie: slice areas follow the shares, with percentage labels', () => {
    const { model, rendered } = renderFixture(5);
    const values = model.plots[0]!.series[0]!.values as number[];
    const total = values.reduce((a, b) => a + b, 0);
    const slices = rendered.scene.components.filter(
      (c) => c.closed && c.fill && c.subpaths[0]!.length > 10
    );
    expect(slices).toHaveLength(values.length);
    const areas = slices.map((s) => area(s.subpaths[0]!));
    const sum = areas.reduce((a, b) => a + b, 0);
    values.forEach((v, i) => expect(areas[i]! / sum).toBeCloseTo(v / total, 2));
    const percents = labelTexts(5).filter((t) => t.endsWith('%'));
    expect(percents).toHaveLength(values.length);
  });

  test('doughnut: rings leave a hole of the hole size', () => {
    const { model, rendered } = renderFixture(6);
    const plot = model.plots[0]!;
    const rings = rendered.scene.components.filter(
      (c) => c.closed && c.fill && c.subpaths[0]!.length > 10
    );
    expect(rings).toHaveLength(plot.series[0]!.values.length);
    const all = rings.flatMap((r) => r.subpaths[0]!);
    const box = boxOf(all);
    const cx = (box.l + box.r) / 2;
    const cy = (box.t + box.b) / 2;
    const outer = (box.r - box.l) / 2;
    const inner = Math.min(...all.map((p) => Math.hypot(p.x - cx, p.y - cy)));
    expect(inner / outer).toBeCloseTo(plot.holeSize / 100, 2);
  });

  test('pie and doughnut: room for outside labels grows with the plot area, not the label size', () => {
    // The reference renderings of `charts.pdf` draw the pie 320 pt wide and the doughnut 325 pt.
    const diameter = (index: number) => {
      const slices = renderFixture(index).rendered.scene.components.filter(
        (c) => c.closed && c.fill && c.subpaths[0]!.length > 10
      );
      const box = boxOf(slices.flatMap((c) => c.subpaths[0]!));
      return box.r - box.l;
    };
    expect(Math.abs(diameter(5) - 320)).toBeLessThan(3);
    expect(Math.abs(diameter(6) - 325)).toBeLessThan(3);
    // A 1.5-inch-tall plot keeps 17 pt, a 6.8-inch-tall one 64 pt.
    expect(outsideLabelRoom(86)).toBeCloseTo(17, 0);
    expect(outsideLabelRoom(467.6)).toBeCloseTo(64, 0);
  });

  test('radar: one polygon per series with a vertex on each category spoke', () => {
    const { model, rendered } = renderFixture(7);
    const count = model.plots[0]!.series[0]!.values.length;
    const outline = rendered.scene.components.find(
      (c) => c.stroke?.hex.toUpperCase() === '4472C4' && c.subpaths[0]!.length >= count
    )!;
    expect(outline).toBeDefined();
    const labels = labelTexts(7);
    expect(labels).toEqual(expect.arrayContaining(['North', 'South', 'East', 'West']));
  });
});
