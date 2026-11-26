/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `ShapeCollection.addChart`, an SDK extension: a chart frame (`p:graphicFrame`,
// `CT_GraphicalObjectFrame`, ECMA-376 Part 1, §19.3.1.21) with a new chart part (`c:chartSpace`,
// `CT_ChartSpace`, §21.2.2.29) and an embedded workbook (`c:externalData`, `CT_ExternalData`,
// §21.2.2.63) that holds the same table.

import { describe, expect, test } from 'bun:test';
import { strFromU8 } from 'fflate';
import type { AutomationHost, AutomationOperation } from '@pptx-editor.dev/core/automation';
import { createBrowserAutomationHost, createPptxEditor } from '@pptx-editor.dev/core/editor';
import { binaryPartOf, readZip, type PptxPackage } from '@pptx-editor.dev/core/store';
import { packageIntegrity } from '../../../../scripts/fixtures/package-integrity.ts';
import { PptxEditor as BrowserEditor } from '../browser.ts';
import {
  ChartType,
  isPptxEditorError,
  type ChartData,
  type PptxEditorRuntime,
  type RequestContext,
} from '../index.ts';
import { createRuntime } from '../runtime/runtime.ts';
import { deck, open, refusedAtSync, reopen } from './support.ts';

const DATA: ChartData = {
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [
    { name: 'Revenue', values: [12, 15, null, 18] },
    { name: 'Cost', values: [8, 9, 10, 11] },
  ],
};

const XY: ChartData = {
  categories: ['1', '2.5', '4'],
  series: [{ name: 'Samples', values: [3, 5, 4] }],
};

function whole(bytes: Uint8Array): PptxPackage {
  const { pkg, issues } = packageIntegrity(bytes);
  expect(issues).toEqual([]);
  return pkg!;
}

const shapesOf = (context: RequestContext) => context.presentation.slides.getItemAt(0).shapes;

/** Each chart of the first slide, as its frame box and its chart read. */
async function charts(runtime: PptxEditorRuntime) {
  return runtime.run(async (context) => {
    const shapes = shapesOf(context).load(
      'items/type,items/left,items/top,items/width,items/height'
    );
    await context.sync();
    const frames = shapes.items.filter((s) => s.type === 'Chart');
    const reads = frames.map((s) => s.chart.load());
    await context.sync();
    return frames.map((s, i) => ({
      box: [s.left, s.top, s.width, s.height].map((v) => Math.round(v)),
      chart: reads[i]!.toJSON(),
    }));
  });
}

describe('ShapeCollection.addChart', () => {
  test('adds a chart with the data, an embedded workbook, and the box, through save and reopen', async () => {
    const runtime = await open();
    const id = await runtime.run(async (context) => {
      const shape = shapesOf(context).addChart('ColumnClustered', DATA, {
        left: 60,
        top: 100,
        width: 480,
        height: 300,
      });
      shape.load('id,type');
      await context.sync();
      expect(shape.type).toBe('Chart');
      return shape.id;
    });
    expect(id).toMatch(/^\d+$/);
    const pkg = whole(await runtime.save());
    const book = [...pkg.parts.keys()].find((n) => /^\/ppt\/embeddings\/.*\.xlsx$/.test(n))!;
    const zip = readZip(binaryPartOf(pkg, book)!);
    if (!zip.ok) throw new Error('the workbook does not read');
    const sheet = strFromU8(zip.entries.get('/xl/worksheets/sheet1.xml')!);
    expect(sheet).toContain('<c r="B5"><v>18</v></c>');
    expect(sheet).not.toContain('r="B4"');
    const again = await reopen(runtime);
    expect(await charts(again)).toEqual([
      {
        box: [60, 100, 480, 300],
        chart: {
          chartType: 'ColumnClustered',
          title: null,
          categories: DATA.categories,
          series: DATA.series,
          workbook: 'Embedded',
          readOnlyReason: null,
        },
      },
    ]);
    runtime.dispose();
    again.dispose();
  });

  test('every ChartType adds a chart that reads back as that type', async () => {
    const runtime = await open();
    const types = Object.values(ChartType);
    await runtime.run(async (context) => {
      for (const type of types) shapesOf(context).addChart(type, type.startsWith('XY') ? XY : DATA);
      await context.sync();
    });
    const read = await charts(runtime);
    expect(read.map((c) => c.chart.chartType)).toEqual(types);
    whole(await runtime.save());
    runtime.dispose();
  });

  test('a scatter chart keeps its X values as numbers in the cache and the workbook', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      shapesOf(context).addChart('XYScatter', XY);
      await context.sync();
    });
    const pkg = whole(await runtime.save());
    const chart = [...pkg.parts.keys()].find((n) => /^\/ppt\/charts\/chart\d+\.xml$/.test(n))!;
    const xml = strFromU8(binaryPartOf(pkg, chart)!);
    // `c:xVal` (§21.2.2.234) holds a number cache (`c:numCache`, §21.2.2.120).
    expect(xml).toMatch(/<c:xVal><c:numRef>.*<c:numCache>/s);
    const book = [...pkg.parts.keys()].find((n) => /^\/ppt\/embeddings\/.*\.xlsx$/.test(n))!;
    const zip = readZip(binaryPartOf(pkg, book)!);
    if (!zip.ok) throw new Error('the workbook does not read');
    const sheet = strFromU8(zip.entries.get('/xl/worksheets/sheet1.xml')!);
    expect(sheet).toContain('<c r="A3"><v>2.5</v></c>');
    expect((await charts(runtime))[0]!.chart.categories).toEqual(XY.categories);
    runtime.dispose();
  });

  test('a refused add after other writes of the same sync leaves the slide unchanged', async () => {
    const runtime = await open();
    const error = await runtime.run((context) =>
      refusedAtSync(context, () => {
        shapesOf(context).addChart('Pie', DATA);
        shapesOf(context).addChart('XYScatter', DATA);
      })
    );
    expect(isPptxEditorError(error) && error.debugInfo.innerError?.code).toBe(
      'edit.chart-data-value'
    );
    const count = await runtime.run(async (context) => {
      const n = shapesOf(context).getCount();
      await context.sync();
      return n.value;
    });
    expect(count).toBe(2);
    runtime.dispose();
  });

  test('without a box, the chart is centered below the top fifth of the slide', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      shapesOf(context).addChart(ChartType.pie, {
        categories: ['A', 'B'],
        series: [{ name: 'Share', values: [3, 1] }],
      });
      await context.sync();
    });
    // A 16:9 slide of 960 × 540 points: 576 points wide, 432 high (at most 4:3 of the width).
    const [{ box, chart }] = await charts(runtime);
    expect(box).toEqual([192, Math.round(108 + (432 - 324) / 2), 576, 324]);
    expect(chart).toMatchObject({ chartType: 'Pie', series: [{ name: 'Share', values: [3, 1] }] });
    runtime.dispose();
  });

  test('the new shape takes a load, a chart read, and new data in the same sync', async () => {
    const runtime = await open();
    const series = await runtime.run(async (context) => {
      const shape = shapesOf(context).addChart('Line', DATA);
      shape.chart.setData({ categories: ['A', 'B'], series: [{ name: 'S', values: [1, 2] }] });
      const chart = shape.chart.load('series');
      await context.sync();
      return chart.series;
    });
    expect(series).toEqual([{ name: 'S', values: [1, 2] }]);
    runtime.dispose();
  });

  test('a chart on a layout reads back', async () => {
    const runtime = await open();
    const type = await runtime.run(async (context) => {
      const layout = context.presentation.slideMasters.getItemAt(0).layouts.getItemAt(0);
      const shape = layout.shapes.addChart('BarStacked', DATA);
      const chart = shape.chart.load('chartType');
      await context.sync();
      return chart.chartType;
    });
    expect(type).toBe('BarStacked');
    whole(await runtime.save());
    runtime.dispose();
  });

  test('a type, data, or box that is not accepted refuses with InvalidArgument and adds nothing', async () => {
    const runtime = await open();
    const cases: [() => (context: RequestContext) => void, string | undefined][] = [
      [() => (c) => shapesOf(c).addChart('Bubble' as never, DATA), undefined],
      [() => (c) => shapesOf(c).addChart('Pie', undefined as never), undefined],
      [
        () => (c) => shapesOf(c).addChart('Pie', { categories: ['a'], series: [] }),
        'edit.chart-data-range',
      ],
      [() => (c) => shapesOf(c).addChart('XYScatter', DATA), 'edit.chart-data-value'],
      [() => (c) => shapesOf(c).addChart('Pie', DATA, { width: -1 }), undefined],
      [
        () => (c) =>
          shapesOf(c).addChart('Pie', {
            categories: Array(4001).fill('a'),
            series: [{ name: 's', values: Array(4001).fill(1) }],
          }),
        'edit.chart-data-range',
      ],
      [
        () => (c) =>
          shapesOf(c).addChart('Pie', {
            categories: ['a'],
            series: [{ name: 'x'.repeat(256), values: [1] }],
          }),
        'edit.chart-data-text',
      ],
    ];
    for (const [make, diagnostic] of cases) {
      const call = make();
      const error = await runtime.run((context) => refusedAtSync(context, () => call(context)));
      expect(isPptxEditorError(error) && error.code).toBe('InvalidArgument');
      if (diagnostic)
        expect(isPptxEditorError(error) && error.debugInfo.innerError?.code).toBe(diagnostic);
    }
    const count = await runtime.run(async (context) => {
      const n = shapesOf(context).getCount();
      await context.sync();
      return n.value;
    });
    expect(count).toBe(2);
    runtime.dispose();
  });

  test('data past the chart XML limits refuses as too large, not as malformed data', async () => {
    // The reader reads at most 16 MiB of chart XML a part. 4000 categories by 255 series are far
    // past it, and refuse before the XML is built; by 55 series, the XML is just past it.
    const categories = Array.from({ length: 4000 }, (_, i) => `C${i}`);
    const values = Array.from({ length: 4000 }, (_, i) => i);
    const many = (n: number) => Array.from({ length: n }, (_, i) => ({ name: `S${i}`, values }));
    const runtime = await open();
    const cases = [
      ['Line', 255],
      ['ColumnClustered', 255],
      ['Line', 55],
      ['ColumnClustered', 55],
    ] as const;
    for (const [type, n] of cases) {
      const error = await runtime.run((context) =>
        refusedAtSync(context, () =>
          shapesOf(context).addChart(type, { categories, series: many(n) })
        )
      );
      expect(isPptxEditorError(error) && error.code).toBe('InvalidArgument');
      expect(isPptxEditorError(error) && error.debugInfo.innerError?.code).toBe(
        'edit.chart-data-range'
      );
    }
    // New data for an existing chart is held to the same limit.
    const error = await runtime.run(async (context) => {
      const shape = shapesOf(context).addChart('Line', DATA);
      await context.sync();
      return refusedAtSync(context, () => shape.chart.setData({ categories, series: many(55) }));
    });
    expect(isPptxEditorError(error) && error.debugInfo.innerError?.code).toBe(
      'edit.chart-data-range'
    );
    // Just under the limit, the chart adds and reads back.
    const added = await runtime.run(async (context) => {
      const shape = shapesOf(context).addChart('Line', { categories, series: many(45) });
      const chart = shape.chart.load('series');
      await context.sync();
      return chart.series.length;
    });
    expect(added).toBe(45);
    runtime.dispose();
  }, 60_000);

  test('long X value text that the cache keeps as short numbers adds and takes new data', async () => {
    // The cache holds each X value as its number (`c:xVal`, §21.2.2.234), so the size of the
    // label text does not count against the chart XML limit.
    const categories = Array.from({ length: 4000 }, (_, i) => '0'.repeat(240) + String(i + 1));
    const values = Array.from({ length: 4000 }, (_, i) => i);
    const series = Array.from({ length: 30 }, (_, i) => ({ name: `S${i}`, values }));
    const runtime = await open();
    const read = await runtime.run(async (context) => {
      const shape = shapesOf(context).addChart('XYScatter', XY);
      await context.sync();
      shape.chart.setData({ categories, series });
      const chart = shape.chart.load('categories,series');
      await context.sync();
      return { first: chart.categories[0], count: chart.series.length };
    });
    expect(read).toEqual({ first: '1', count: 30 });
    runtime.dispose();
  }, 60_000);

  test('the browser runtime adds the same chart as the server runtime, as one undo step', async () => {
    const live = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await live.load(deck())).ok).toBe(true);
    const runtime = BrowserEditor.createBrowser(live);
    const server = await open();
    for (const r of [runtime, server])
      await r.run(async (context) => {
        shapesOf(context).addChart('AreaStacked', DATA, {
          left: 10,
          top: 20,
          width: 300,
          height: 200,
        });
        await context.sync();
      });
    expect(await charts(runtime)).toEqual(await charts(server));
    whole(live.save()!);
    const before = live.snapshot().shapes!.length;
    expect(live.exec({ type: 'undo' }).ok).toBe(true);
    expect(live.snapshot().shapes!.length).toBe(before - 1);
    expect(live.snapshot().canUndo).toBe(false);
    server.dispose();
  });

  test('an add sends the lookups and one addShape operation', async () => {
    // The collaboration test (`packages/pro/src/__tests__/yjs-add-chart.test.ts`) sends these
    // operations to two linked editors; this test keeps them equal to what `addChart` sends.
    const live = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await live.load(deck())).ok).toBe(true);
    const inner = createBrowserAutomationHost(live);
    const batches: AutomationOperation[][] = [];
    const host: AutomationHost = {
      ...inner,
      execute: (request) => {
        batches.push([...request.operations]);
        return inner.execute(request);
      },
    };
    await createRuntime({ host, save: false }).run(async (context) => {
      shapesOf(context).addChart('ColumnClustered', DATA, { left: 60, top: 100 });
      await context.sync();
    });
    expect(batches).toEqual([
      [
        { op: 'getPresentation' },
        { op: 'getSlide', key: { index: 0 } },
        {
          op: 'addShape',
          container: { result: 1 },
          shape: { kind: 'chart', chartType: 'columnClustered', data: DATA },
          box: { left: 60, top: 100 },
        },
      ],
    ]);
  });
});
