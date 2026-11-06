/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { describe, expect, test } from 'bun:test';
import { readChartModel } from '../chart-model.ts';
import { formatChartNumber } from '../chart-number-format.ts';
import { chartInScope } from '../chart-projection.ts';
import { autoValueScale, logValueScale, scaleRatio } from '../chart-scale.ts';
import { chartSpace, fixtureCharts, numericSeries } from './chart-support.ts';

describe('the chart model (ECMA-376 Part 1, 21.2)', () => {
  test('reads the plot type, grouping, and cached values of each fixture chart', () => {
    const models = fixtureCharts().map(({ chart }) => readChartModel(chart.chartSpace)!);
    expect(models.map((m) => m.plots.map((p) => `${p.kind}:${p.grouping}`))).toEqual([
      ['bar:clustered'],
      ['bar:stacked'],
      ['line:standard'],
      ['area:stacked'],
      ['scatter:standard'],
      ['pie:standard'],
      ['doughnut:standard'],
      ['radar:standard'],
    ]);
    const column = models[0]!.plots[0]!;
    expect(column.horizontal).toBe(false);
    expect(models[1]!.plots[0]!.horizontal).toBe(true);
    expect(column.series.map((s) => s.name)).toEqual(['Plan', 'Actual', 'Forecast']);
    expect(column.series[0]!.values).toEqual([42, 35, 51, 28]);
    expect(column.series[0]!.categories).toEqual(['North', 'South', 'East', 'West']);
    expect(models.every(chartInScope)).toBe(true);
  });

  test('refuses a chart past the point bound, so it keeps the placeholder', () => {
    const many = Array.from({ length: 5_000 }, (_, i) => i);
    expect(readChartModel(chartSpace(`<c:barChart>${numericSeries(many)}</c:barChart>`))).toBe(
      null
    );
    expect(
      readChartModel(chartSpace(`<c:barChart>${numericSeries([1, 2, 3])}</c:barChart>`))
    ).not.toBe(null);
  });

  test('reads a point that is not a number as a blank', () => {
    const xml =
      '<c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:numCache>' +
      '<c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>NaN</c:v></c:pt>' +
      '<c:pt idx="2"><c:v>1e999</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:lineChart>';
    const model = readChartModel(chartSpace(xml))!;
    expect(model.plots[0]!.series[0]!.values).toEqual([1, null, null]);
  });

  test('keeps 3-D, stock, bubble, of-pie, and surface charts and trendlines out of scope', () => {
    const series = numericSeries([1, 2, 3]);
    for (const plot of [
      `<c:bar3DChart>${series}</c:bar3DChart>`,
      `<c:stockChart>${series}${numericSeries([2, 3, 4], 1)}</c:stockChart>`,
      `<c:bubbleChart>${series}</c:bubbleChart>`,
      `<c:ofPieChart>${series}</c:ofPieChart>`,
      `<c:surfaceChart>${series}</c:surfaceChart>`,
    ]) {
      const model = readChartModel(chartSpace(plot));
      expect(model === null || !chartInScope(model)).toBe(true);
    }
    const trend = series.replace(
      '</c:ser>',
      '<c:trendline><c:trendlineType val="linear"/></c:trendline></c:ser>'
    );
    expect(chartInScope(readChartModel(chartSpace(`<c:lineChart>${trend}</c:lineChart>`))!)).toBe(
      false
    );
  });
});

describe('chart scales', () => {
  test('pick round ends and steps for the data, as the fixture axes show', () => {
    const scale = autoValueScale({
      dataMin: 0,
      dataMax: 55,
      lengthPt: 300,
      minIntervalPt: 20,
      fixedMin: null,
      fixedMax: null,
      fixedMajor: null,
    });
    expect([scale.min, scale.max, scale.major]).toEqual([0, 60, 10]);
    expect(scale.ticks).toEqual([0, 10, 20, 30, 40, 50, 60]);
    expect(scaleRatio(scale, 30)).toBeCloseTo(0.5);
  });

  test('keep a fixed minimum, maximum, and step', () => {
    const scale = autoValueScale({
      dataMin: 3,
      dataMax: 7,
      lengthPt: 300,
      minIntervalPt: 20,
      fixedMin: 2,
      fixedMax: 8,
      fixedMajor: 2,
    });
    expect(scale.ticks).toEqual([2, 4, 6, 8]);
  });

  test('place values by their logarithm on a logarithmic scale', () => {
    const scale = logValueScale(3, 800, 10, null, null);
    expect(scale.ticks).toEqual([1, 10, 100, 1000]);
    expect(scaleRatio(scale, 100)).toBeCloseTo(2 / 3);
  });
});

describe('chart number formats', () => {
  test('format values with the cached format codes', () => {
    expect(formatChartNumber(0.27, '0%')).toBe('27%');
    expect(formatChartNumber(1234.5, '#,##0.00')).toBe('1,234.50');
    expect(formatChartNumber(42, 'General')).toBe('42');
    expect(formatChartNumber(45_000, 'yyyy-mm-dd')).toBe('2023-03-15');
  });
});
