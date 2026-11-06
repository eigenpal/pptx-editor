/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The legend and radar spacing rules at several text and frame sizes. The first group checks
// the rules against values measured on the `chart-sizes` and `charts` reference PDFs (Carlito,
// line 1.221 sizes). The second group renders `fixtures/chart-sizes.pptx` and checks that the
// painted charts follow the rules at every size.

import { describe, expect, test } from 'bun:test';
import { keyMetrics, type LegendEntry } from '../chart-legend.ts';
import { labelRoom } from '../chart-radar.ts';
import type { ChartComponent, Point } from '../chart-scene.ts';
import type { ChartLabel } from '../chart-text.ts';
import { fixtureCharts, renderFixture } from './chart-support.ts';

const CARLITO_LINE = 1.221;
const square: LegendEntry = {
  text: 'Plan',
  look: { fill: null, line: null, marker: null },
  line: false,
};
const lined: LegendEntry = {
  text: 'Plan',
  look: {
    fill: null,
    line: { hex: '4472C4', alpha: 1, widthPt: 2.25, cap: 'round', dash: null },
    marker: null,
  },
  line: true,
};

describe('chart spacing rules match the reference PDFs', () => {
  test('a square key starts its text 0.825 sizes less 0.4 pt past the key start', () => {
    // [text size, key start to text start]
    for (const [size, measured] of [
      [9, 7.03],
      [11.97, 9.46],
      [14, 11.16],
      [20, 16.1],
    ] as const)
      expect([size, Math.abs(keyMetrics(square, size).text - measured) < 0.1]).toEqual([
        size,
        true,
      ]);
  });

  test('a line key is 19.2 pt long at every size, with its text 21.23 pt from its start', () => {
    for (const size of [9, 11.97, 14, 20]) {
      expect(keyMetrics(lined, size).key).toBe(19.2);
      expect(keyMetrics(lined, size).text).toBeCloseTo(21.23, 6);
    }
  });

  test('a vertical legend keeps room for half a square key, or 15.66 pt for a line key', () => {
    // [text size, measured gap from the plot to the square keys]
    for (const [size, measured] of [
      [9, 15.72],
      [14, 17.09],
      [20, 18.75],
    ] as const)
      expect([size, Math.abs(keyMetrics(square, size).gap - measured) < 0.1]).toEqual([size, true]);
    for (const size of [9, 11.97, 14, 20]) expect(keyMetrics(lined, size).gap).toBe(15.66);
  });

  test('a radar keeps half a category line and the label room of its plot height', () => {
    // [text size, chart height, measured radius]
    for (const [size, height, measured] of [
      [9, 237.6, 99.21],
      [11.97, 424.8, 188.25],
      [14, 316.8, 131.12],
      [20, 424.8, 175.89],
    ] as const) {
      const radius = (height - 22) / 2 - ((CARLITO_LINE * size) / 2 + labelRoom(size));
      expect([size, Math.abs(radius - measured) < 0.1]).toEqual([size, true]);
    }
  });

  test('labelRoom is never negative', () => {
    expect(labelRoom(1)).toBe(0);
    expect(labelRoom(0)).toBe(0);
  });
});

// --- The painted charts of `chart-sizes.pptx` ---------------------------------------------------

const charts = fixtureCharts('chart-sizes', 12);
const render = (index: number) => renderFixture(index, charts).rendered;

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

const label = (labels: readonly ChartLabel[], text: string): ChartLabel =>
  labels.find((l) => l.paragraphs[0]?.[0]?.text === text)!;

const sizeOf = (l: ChartLabel): number => l.paragraphs[0]![0]!.style.sizePt;

/** The gridlines' box: the plot of a column chart, or the outer ring of a radar. */
const gridBox = (components: readonly ChartComponent[]): Box =>
  boxOf(
    components
      .filter((c) => c.stroke?.hex.toUpperCase() === 'D9D9D9')
      .flatMap((c) => c.subpaths.flat())
  );

/** The square legend key of the first series: 0.55 sizes wide. */
const squareKey = (components: readonly ChartComponent[], size: number): Box =>
  components
    .filter((c) => c.closed && c.fill?.hex.toUpperCase() === '4472C4')
    .map((c) => boxOf(c.subpaths[0]!))
    .find(
      (b) => Math.abs(b.r - b.l - 0.55 * size) < 0.01 && Math.abs(b.b - b.t - 0.55 * size) < 0.01
    )!;

/** The line legend key of the first series: a 19.2 pt level stroke. */
const lineKey = (components: readonly ChartComponent[]): Box =>
  components
    .filter((c) => !c.closed && c.stroke?.hex.toUpperCase() === '4472C4')
    .map((c) => boxOf(c.subpaths[0]!))
    .find((b) => Math.abs(b.r - b.l - 19.2) < 0.01 && b.b === b.t)!;

const height = (index: number): number => charts[index]!.frame.xfrm.cy / 12_700;
const width = (index: number): number => charts[index]!.frame.xfrm.cx / 12_700;

/** A label's text width: its frame is 1.3 text widths and two sizes wide. */
const textWidth = (l: ChartLabel): number => (l.width - 2 * sizeOf(l)) / 1.3;

describe('the chart-sizes charts follow the rules', () => {
  test('right legends: rows one line and 5.85 pt apart, and the plot ends the gap before the keys', () => {
    for (const index of [0, 1, 2]) {
      const { scene, labels } = render(index);
      const plan = label(labels, 'Plan');
      const actual = label(labels, 'Actual');
      const size = sizeOf(plan);
      expect(actual.y - plan.y).toBeCloseTo(plan.height + 5.85, 6);
      const key = squareKey(scene.components, size);
      expect(key.l - gridBox(scene.components).r).toBeCloseTo(13.25 + 0.275 * size, 6);
      expect(plan.x - key.l).toBeCloseTo(0.825 * size - 0.4, 6);
    }
  });

  test('bottom legends: the keys centre 9 pt and 0.6 sizes above the chart bottom', () => {
    for (const index of [3, 5, 7]) {
      const { scene, labels } = render(index);
      const first = index === 7 ? 'North' : 'Plan';
      const size = sizeOf(label(labels, first));
      const key = squareKey(scene.components, size);
      expect((key.t + key.b) / 2).toBeCloseTo(height(index) - (9 + 0.6 * size), 6);
      // From the first key to the end of the last text, the row centres 0.18 sizes right.
      const last = label(labels, index === 7 ? 'West' : 'Forecast');
      const end = last.x + textWidth(last);
      expect((key.l + end) / 2 - width(index) / 2).toBeCloseTo(0.18 * size, 6);
    }
  });

  test('line keys in a row: 19.2 pt long, 6.43 pt and 0.386 sizes apart, centred 1.76 pt right', () => {
    for (const [index, side] of [
      [4, 'b'],
      [8, 't'],
    ] as const) {
      const { scene, labels } = render(index);
      const plan = label(labels, 'Plan');
      const size = sizeOf(plan);
      const key = lineKey(scene.components);
      const center = side === 'b' ? height(index) - (9 + 0.6 * size) : 9 + 0.6 * size;
      expect(key.t).toBeCloseTo(center, 6);
      expect(plan.x - key.l).toBeCloseTo(21.23, 6);
      const actual = label(labels, 'Actual');
      expect(actual.x - 21.23 - (plan.x + textWidth(plan))).toBeCloseTo(6.43 + 0.386 * size, 6);
      const forecast = label(labels, 'Forecast');
      const end = forecast.x + textWidth(forecast);
      expect((key.l + end) / 2 - width(index) / 2).toBeCloseTo(1.76, 6);
    }
  });

  test('radars: the radius keeps the label room, and side labels sit 0.04 radii out', () => {
    for (const index of [9, 10, 11]) {
      const { scene, labels } = render(index);
      const south = label(labels, 'South');
      const size = sizeOf(south);
      const ring = gridBox(scene.components);
      const radius = (ring.b - ring.t) / 2;
      const cx = (ring.l + ring.r) / 2;
      expect(radius).toBeCloseTo(
        (height(index) - 22) / 2 - (south.height / 2 + labelRoom(size)),
        6
      );
      // A label frame is 1.3 text widths and two sizes wide (`paintRadar`).
      const textWidth = (south.width - 2 * size) / 1.3;
      expect(south.x + south.width / 2 - textWidth / 2 - (cx + radius)).toBeCloseTo(
        0.04 * radius,
        6
      );
      // The plot box ends 15.66 pt before the line keys, and the ring centres in it.
      const key = lineKey(scene.components);
      expect(key.l - (2 * cx - 11)).toBeCloseTo(15.66, 6);
    }
  });
});
