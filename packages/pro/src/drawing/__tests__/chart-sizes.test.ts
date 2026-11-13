/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The legend and radar spacing rules at several text and frame sizes. The first group checks
// the rules against values measured on the `chart-sizes` and `charts` reference PDFs (Carlito,
// line 1.221 sizes). The second group renders `fixtures/chart-sizes.pptx` and checks that the
// painted charts follow the rules at every size.

import { describe, expect, test } from 'bun:test';
import { keyMetrics, rowKeyInset, rowPlotGap, type LegendEntry } from '../chart-legend.ts';
import { bestFitInside } from '../chart-pie.ts';
import { labelRoom } from '../chart-radar.ts';
import { autoValueScale } from '../chart-scale.ts';
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
    // [text size, measured key length, measured key start to text start]
    for (const [size, key, text] of [
      [9, 19.2, 21.22],
      [11.97, 19.2, 21.23],
      [14, 19.2, 21.23],
      [20, 19.2, 21.23],
    ] as const) {
      const metrics = keyMetrics(lined, size);
      expect([size, Math.abs(metrics.key - key) < 0.05]).toEqual([size, true]);
      expect([size, Math.abs(metrics.text - text) < 0.05]).toEqual([size, true]);
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
    // [text size, measured gap from the plot to the line keys]
    for (const [size, measured] of [
      [9, 15.65],
      [11.97, 15.65],
      [14, 15.66],
      [20, 15.66],
    ] as const)
      expect([size, Math.abs(keyMetrics(lined, size).gap - measured) < 0.05]).toEqual([size, true]);
  });

  test('a radar keeps half a category line and the label room of its plot height', () => {
    // [text size, chart height, measured radius]
    for (const [size, height, measured] of [
      [9, 237.6, 99.21],
      [9, 424.8, 192.82],
      [11.97, 424.8, 188.25],
      [14, 316.8, 131.12],
      [20, 424.8, 175.89],
    ] as const) {
      const radius = (height - 22) / 2 - ((CARLITO_LINE * size) / 2 + labelRoom(size));
      expect([size, Math.abs(radius - measured) < 0.1]).toEqual([size, true]);
    }
  });

  test('a row legend centres its keys 9 pt and 0.6045 sizes inside the chart edge', () => {
    // [text size, measured key centre from the bottom or top edge]
    for (const [size, measured] of [
      [9, 14.44],
      [11.97, 16.23],
      [14, 17.46],
      [20, 21.08],
    ] as const)
      expect([size, Math.abs(rowKeyInset(size) - measured) < 0.05]).toEqual([size, true]);
  });

  test('the plot ends 14.2 pt and 0.59 sizes past the key centre line of a bottom legend', () => {
    // [text size, measured gap from the key centre line to the bottom of the plot box]. The
    // plot's bottom comes from where the reference puts the category labels.
    for (const [size, measured] of [
      [9, 19.51],
      [11.97, 21.16],
      [20, 26.0],
    ] as const)
      expect([size, Math.abs(rowPlotGap(size) - measured) < 0.15]).toEqual([size, true]);
  });

  test('row entries leave the measured room after their text, and centre the row', () => {
    const point: LegendEntry = { ...square, point: true };
    // [entry, text size, measured room from the text end to the next key, measured row shift]
    for (const [name, entry, size, after, shift] of [
      ['square', square, 9, 7.84, 1.79],
      ['square', square, 11.97, 10.3, 2.2],
      ['square', square, 20, 16.94, 3.31],
      ['point', point, 11.97, 9.87, 2.2],
      ['point', point, 14, 11.48, 2.48],
      ['line', lined, 9, 9.87, 1.76],
      ['line', lined, 11.97, 11.03, 1.77],
      ['line', lined, 20, 14.15, 1.77],
    ] as const) {
      const metrics = keyMetrics(entry, size);
      expect([name, size, Math.abs(metrics.after - after) < 0.05]).toEqual([name, size, true]);
      expect([name, size, Math.abs(metrics.shift - shift) < 0.05]).toEqual([name, size, true]);
    }
  });

  test('vertical legend rows are one line and 5.85 pt apart', () => {
    // [text size, measured row pitch]
    for (const [size, measured] of [
      [9, 16.92],
      [11.97, 20.45],
      [14, 22.92],
      [20, 30.24],
    ] as const)
      expect([size, Math.abs(CARLITO_LINE * size + 5.85 - measured) < 0.1]).toEqual([size, true]);
  });

  test('radar labels sit 0.04 radii past a side spoke, and 0.16 sizes more above the top one', () => {
    // [text size, radius, measured side gap, measured gap below the top label's line]. The two
    // 9 pt radars keep 3.9 pt and 7.7 pt, so the gaps follow the radius, not the text size.
    for (const [size, radius, side, top] of [
      [9, 99.21, 3.93, 5.36],
      [9, 192.82, 7.67, 9.2],
      [11.97, 188.25, 7.49, 9.3],
      [14, 131.12, 5.2, 7.37],
      [20, 175.89, 6.99, 10.22],
    ] as const) {
      expect([size, Math.abs(0.04 * radius - side) < 0.1]).toEqual([size, true]);
      // 0.08 sizes of `RISE_GAP`, and 0.08 sizes of `LABEL_LIFT`.
      expect([size, Math.abs(0.04 * radius + 0.16 * size - top) < 0.15]).toEqual([size, true]);
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

/** The square legend key in `hex` (the first series by default): 0.55 sizes wide. */
const squareKey = (components: readonly ChartComponent[], size: number, hex = '4472C4'): Box =>
  components
    .filter((c) => c.closed && c.fill?.hex.toUpperCase() === hex)
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

/** The chart's box before the legend takes its room (`CHART_MARGIN` on each side). */
const MARGIN = 11;
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

  test('bottom legends: keys at rowKeyInset, room after each text, and the row shift', () => {
    for (const index of [3, 5, 6]) {
      const { scene, labels } = render(index);
      const first = index === 6 ? 'North' : 'Plan';
      const size = sizeOf(label(labels, first));
      const key = squareKey(scene.components, size);
      expect((key.t + key.b) / 2).toBeCloseTo(height(index) - rowKeyInset(size), 6);
      const metrics = keyMetrics(index === 6 ? { ...square, point: true } : square, size);
      const firstLabel = label(labels, first);
      const next = squareKey(scene.components, size, 'ED7D31');
      expect(next.l - (firstLabel.x + textWidth(firstLabel))).toBeCloseTo(metrics.after, 6);
      // From the first key to the end of the last text, the row centres `shift` right.
      const last = label(labels, index === 6 ? 'West' : 'Forecast');
      const end = last.x + textWidth(last);
      expect((key.l + end) / 2 - width(index) / 2).toBeCloseTo(metrics.shift, 6);
    }
  });

  test('line keys in a row: 19.2 pt long, 6.43 pt and 0.386 sizes apart, centred 1.76 pt right', () => {
    for (const [index, side] of [
      [4, 'b'],
      [7, 't'],
    ] as const) {
      const { scene, labels } = render(index);
      const plan = label(labels, 'Plan');
      const size = sizeOf(plan);
      const key = lineKey(scene.components);
      const center = side === 'b' ? height(index) - rowKeyInset(size) : rowKeyInset(size);
      expect(key.t).toBeCloseTo(center, 6);
      expect(plan.x - key.l).toBeCloseTo(21.23, 6);
      const actual = label(labels, 'Actual');
      expect(actual.x - 21.23 - (plan.x + textWidth(plan))).toBeCloseTo(6.43 + 0.386 * size, 6);
      const forecast = label(labels, 'Forecast');
      const end = forecast.x + textWidth(forecast);
      expect((key.l + end) / 2 - width(index) / 2).toBeCloseTo(1.76, 6);
    }
  });

  test('radars: the radius keeps the label room, and the labels keep their gaps', () => {
    for (const index of [8, 9, 10, 11]) {
      const { scene, labels } = render(index);
      const south = label(labels, 'South');
      const size = sizeOf(south);
      const ring = gridBox(scene.components);
      const radius = (ring.b - ring.t) / 2;
      const cx = (ring.l + ring.r) / 2;
      const cy = (ring.t + ring.b) / 2;
      const half = (height(index) - 2 * MARGIN) / 2;
      expect(radius).toBeCloseTo(half - (south.height / 2 + labelRoom(size)), 6);
      // The side label starts 0.04 radii past its spoke end.
      const southLeft = south.x + south.width / 2 - textWidth(south) / 2;
      expect(southLeft - (cx + radius)).toBeCloseTo(0.04 * radius, 6);
      // The top label ends 0.04 radii and 0.16 sizes above its spoke end, and the bottom label
      // starts 0.04 radii below its spoke end: `RISE_GAP` and `LABEL_LIFT` cancel there.
      const north = label(labels, 'North');
      expect(cy - radius - (north.y + north.height)).toBeCloseTo(0.04 * radius + 0.16 * size, 6);
      const east = label(labels, 'East');
      expect(east.y - (cy + radius)).toBeCloseTo(0.04 * radius, 6);
      // The plot box (from the left margin) ends 15.66 pt before the line keys, and the ring
      // centres in it.
      const plotRight = 2 * cx - MARGIN;
      expect(lineKey(scene.components).l - plotRight).toBeCloseTo(15.66, 6);
    }
  });
});

describe('value axes and pie labels follow the reference PDFs', () => {
  test('an automatic major unit keeps one label line between ticks, not 1.6 sizes', () => {
    // [slide index, text size]: the 20 pt column chart and the 14 pt and 20 pt radars show
    // 0, 10, ..., 60 in the reference, with 29.4 pt, 21.9 pt, and 29.3 pt between ticks.
    for (const [index, size] of [
      [2, 20],
      [10, 14],
      [11, 20],
    ] as const) {
      const texts = render(index).labels.map((l) => l.paragraphs[0]?.[0]?.text);
      for (const tick of ['0', '10', '20', '30', '40', '50', '60'])
        expect([index, size, tick, texts.includes(tick)]).toEqual([index, size, tick, true]);
    }
    // The 176.6 pt axis of the 20 pt column chart has room for six intervals of one line, but
    // not for six of 1.6 sizes.
    const scale = autoValueScale({
      dataMin: 0,
      dataMax: 55,
      lengthPt: 176.6,
      minIntervalPt: CARLITO_LINE * 20,
      fixedMin: null,
      fixedMax: null,
      fixedMajor: null,
    });
    expect(scale.major).toBe(10);
    expect(176.6 / 6).toBeLessThan(1.6 * 20);
  });

  test('the top value label lowers the plot by half a line and the lift, less 6.5 pt and 0.05 sizes', () => {
    // The 20 pt column chart has no title: the plot starts at the 11 pt chart margin plus the
    // overhang, 6.31 pt for a Carlito line of 1.2207 sizes.
    const { scene, labels } = render(2);
    const line = label(labels, '60').height;
    expect(gridBox(scene.components).t).toBeCloseTo(
      MARGIN + line / 2 + (0.08 - 0.05) * 20 - 6.5,
      6
    );
  });

  test('a best-fit label sits between its slice edges, 4 pt inside the circle', () => {
    const width = 24.2;
    const height = CARLITO_LINE * 14;
    // A wide label in a quarter slice turns towards the level line, and keeps equal room to
    // both edges. The label then moves half a pixel (0.375 pt) left.
    const placed = bestFitInside(width, height, { cx: 0, cy: 0, outer: 80, from: 0, sweep: 90 });
    const x = placed.x + 0.375;
    const up = -placed.y;
    expect(x - width / 2).toBeCloseTo(up - height / 2, 6);
    expect(Math.hypot(x + width / 2, up + height / 2)).toBeCloseTo(76, 6);
    expect((Math.atan2(x, up) * 180) / Math.PI).toBeGreaterThan(46);
    // A slice of half the pie keeps its middle angle.
    const half = bestFitInside(width, height, { cx: 0, cy: 0, outer: 80, from: 270, sweep: 180 });
    expect(half.x + 0.375).toBeCloseTo(0, 6);
    // The labels of the 14 pt pie (`chart-sizes` slide 7): centres measured on the reference
    // PDF, from the pie centre, in points with y down.
    const measured = [
      [46.61, -39.44],
      [42.86, 43.6],
      [-51.64, 34.0],
      [-33.64, -53.12],
    ] as const;
    let from = 0;
    [42, 35, 51, 28].forEach((value, i) => {
      const sweep = (value / 156) * 360;
      const label = bestFitInside(width, height, { cx: 0, cy: 0, outer: 80.1, from, sweep });
      expect([i, Math.abs(label.x - measured[i]![0]) <= 0.5]).toEqual([i, true]);
      expect([i, Math.abs(label.y - measured[i]![1]) <= 0.5]).toEqual([i, true]);
      from += sweep;
    });
  });
});
