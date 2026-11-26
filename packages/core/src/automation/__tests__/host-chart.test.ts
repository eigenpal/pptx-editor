// `readChart` and `setChartData` through the automation host, over the editor's chart edits
// (`commands/chart-data.ts`): the chart part (`c:chartSpace`, ECMA-376 Part 1, §21.2.2.29) of a
// chart frame. A shape that is not a chart refuses with `not-chart`; data of the wrong form
// names `chartData`; the engine's reason is the detail of every other refusal.

import { describe, expect, test } from 'bun:test';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { chartDataDeck } from '../../../../../scripts/fixtures/chart-data.ts';
import { PptxDocumentStore, readPptx } from '../../store/index.ts';
import { createStoreAutomationHost, type AutomationOperation } from '../index.ts';

function setup() {
  const read = readPptx(buildPptx(chartDataDeck()));
  if (!read.ok) throw new Error('fixture');
  const store = new PptxDocumentStore(read.pkg);
  return { store, host: createStoreAutomationHost(store) };
}

/** The lookups of shape `id` on the first slide, then `op` on it. */
function onShape(id: string, op: Record<string, unknown>): AutomationOperation[] {
  return [
    { op: 'getSlide', key: { index: 0 } },
    { op: 'getShape', container: { result: 0 }, key: { id } },
    { ...op, shape: { result: 1 } } as AutomationOperation,
  ];
}

describe('automation host: charts', () => {
  test('readChart answers the data of a chart frame', () => {
    const { host } = setup();
    const response = host.execute({ operations: onShape('3', { op: 'readChart' }) });
    expect(response.ok).toBe(true);
    const result = response.results[2];
    expect(result).toMatchObject({
      status: 'ok',
      value: {
        kind: 'chart',
        chart: {
          type: 'columnClustered',
          title: 'Sales by region',
          categories: ['North', 'South', 'East', 'West'],
          workbook: 'embedded',
          readOnlyReason: null,
        },
      },
    });
  });

  test('a shape that is not a chart refuses with not-chart', () => {
    const { host } = setup();
    for (const op of [
      { op: 'readChart' },
      { op: 'setChartData', data: { categories: ['a'], series: [] } },
    ]) {
      const response = host.execute({ operations: onShape('2', op) });
      expect(response.results[2]).toMatchObject({
        status: 'error',
        error: { code: 'unsupported-content', detail: 'not-chart' },
      });
    }
  });

  test('setChartData refuses data of the wrong form, and changes nothing', () => {
    const { store, host } = setup();
    const before = store.revision;
    const cases: [unknown, string, string][] = [
      [null, 'invalid-argument', 'chartData'],
      [
        { categories: ['a'], series: [{ name: 'S', values: [1, 2] }] },
        'invalid-argument',
        'chartData',
      ],
      [
        { categories: ['a'], series: [{ name: 'S', values: ['1'] }] },
        'invalid-argument',
        'errors.chartData.notNumber',
      ],
      [{ categories: [], series: [] }, 'invalid-argument', 'errors.chartData.empty'],
    ];
    for (const [data, code, detail] of cases) {
      const response = host.execute({ operations: onShape('3', { op: 'setChartData', data }) });
      expect(response.results[2]).toMatchObject({ status: 'error', error: { code, detail } });
    }
    expect(store.revision).toBe(before);
  });

  test('setChartData writes the caches, and a read in the same batch sees them', () => {
    const { host } = setup();
    const data = { categories: ['A', 'B'], series: [{ name: 'Only', values: [1, null] }] };
    const response = host.execute({
      operations: [
        ...onShape('3', { op: 'setChartData', data }),
        { op: 'readChart', shape: { result: 1 } },
      ],
    });
    expect(response.ok).toBe(true);
    expect(response.results[3]).toMatchObject({
      status: 'ok',
      value: { kind: 'chart', chart: { categories: ['A', 'B'], series: [data.series[0]] } },
    });
  });
});
