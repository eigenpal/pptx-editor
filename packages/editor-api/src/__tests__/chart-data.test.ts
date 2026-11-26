/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `Shape.chart` and `Chart.setData`, an SDK extension over the chart part (`c:chartSpace`,
// ECMA-376 Part 1, §21.2.2.29): the series (`c:ser`), their names (`c:tx`, §21.2.2.215), the
// categories (`c:cat`, §21.2.2.24), the values (`c:val`, §21.2.2.224), their caches
// (`c:numCache`, §21.2.2.120; `c:strCache`, §21.2.2.199), and the embedded workbook
// (`c:externalData`, §21.2.2.63).

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import type { AutomationHost, AutomationOperation } from '@pptx-editor.dev/core/automation';
import { createBrowserAutomationHost, createPptxEditor } from '@pptx-editor.dev/core/editor';
import { binaryPartOf, readPptx, readZip, type PptxPackage } from '@pptx-editor.dev/core/store';
import { buildPptx, type FixtureOptions } from '../../../../scripts/fixtures/builder.ts';
import { chartDataDeck } from '../../../../scripts/fixtures/chart-data.ts';
import { packageIntegrity } from '../../../../scripts/fixtures/package-integrity.ts';
import { PptxEditor as BrowserEditor } from '../browser.ts';
import { isPptxEditorError, type PptxEditorRuntime, type RequestContext } from '../index.ts';
import { createRuntime } from '../runtime/runtime.ts';
import { open, refusedAtSync, reopen } from './support.ts';

const FIXTURES = join(import.meta.dir, '../../../../fixtures');

/** The chart data deck, with its chart parts and relationships changed by `edit`. */
function deckWith(edit?: (parts: [string, string, string][]) => void): Uint8Array {
  const options = chartDataDeck();
  const parts = (options.extraParts ?? []).map((p) => [...p] as [string, string, string]);
  edit?.(parts);
  return buildPptx({ ...options, extraParts: parts } as FixtureOptions);
}

/** The column chart of the deck with a line plot beside its bars: a combination chart. */
const combination = (): Uint8Array =>
  deckWith((parts) => {
    const chart = parts.find((p) => p[0] === 'ppt/charts/chart1.xml')!;
    chart[2] = chart[2].replace(
      '</c:barChart>',
      '</c:barChart><c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>' +
        '<c:ser><c:idx val="9"/><c:order val="9"/><c:val><c:numLit><c:formatCode>General</c:formatCode>' +
        '<c:ptCount val="4"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>' +
        '<c:axId val="1001"/><c:axId val="1002"/></c:lineChart>'
    );
  });

/** The column chart with its workbook linked as an external file. */
const linked = (): Uint8Array =>
  deckWith((parts) => {
    const rels = parts.find((p) => p[0] === 'ppt/charts/_rels/chart1.xml.rels')!;
    rels[2] =
      '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="file:///C:/data.xlsx" TargetMode="External"/></Relationships>';
  });

/** The chart frame of slide `index`. */
function chartAt(context: RequestContext, index = 0) {
  return context.presentation.slides.getItemAt(index).shapes.getItem('3');
}

type ChartRead = {
  chartType: string | null;
  title: string | null;
  categories: string[];
  series: { name: string; values: (number | null)[] }[];
  workbook: string;
  readOnlyReason: string | null;
};

async function readChart(runtime: PptxEditorRuntime, index = 0): Promise<ChartRead> {
  return runtime.run(async (context) => {
    const chart = chartAt(context, index).chart.load();
    await context.sync();
    return chart.toJSON() as ChartRead;
  });
}

/** The text of a part of a saved package. */
function partText(pkg: PptxPackage, name: string): string {
  const part = pkg.parts.get(name);
  if (!part || part.kind === 'xml') throw new Error(`no byte part ${name}`);
  return new TextDecoder().decode(part.bytes);
}

/** An entry of the embedded workbook of a saved package. */
function workbookEntry(pkg: PptxPackage, entry: string): string {
  const zip = readZip(binaryPartOf(pkg, '/ppt/embeddings/Workbook1.xlsx')!);
  if (!zip.ok) throw new Error('the workbook does not read');
  return strFromU8(zip.entries.get(entry)!);
}

function whole(bytes: Uint8Array): PptxPackage {
  const { pkg, issues } = packageIntegrity(bytes);
  expect(issues).toEqual([]);
  return pkg!;
}

const NEW_DATA = {
  categories: ['North', 'South', 'East', 'West'],
  series: [
    { name: 'Plan', values: [10, 20, 30, 40] },
    { name: 'Actual', values: [11, null, 31, 41] },
    { name: 'Forecast', values: [12, 22, 32, 42] },
  ],
};

/** The column chart with a hidden series (`c15:filteredBarSeries` in the plot's `c:extLst`). */
const filtered = (): Uint8Array =>
  deckWith((parts) => {
    const chart = parts.find((p) => p[0] === 'ppt/charts/chart1.xml')!;
    chart[2] = chart[2].replace(
      '</c:barChart>',
      '<c:extLst><c:ext uri="{02D57815-91ED-43cb-92C2-25804820EDAC}">' +
        '<c15:filteredBarSeries xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"/>' +
        '</c:ext></c:extLst></c:barChart>'
    );
  });

/** The deck with the first chart part gone: the frame names a part that is missing. */
const missingPart = (): Uint8Array =>
  deckWith((parts) => {
    for (const name of ['ppt/charts/chart1.xml', 'ppt/charts/_rels/chart1.xml.rels'])
      parts.splice(
        parts.findIndex((p) => p[0] === name),
        1
      );
  });

/** A slide whose frame holds a chart of an extension namespace (`cx:chart`). */
const extensionChart = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Chart 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
          '<p:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></p:xfrm>' +
          '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex">' +
          '<cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" r:id="rId99"/>' +
          '</a:graphicData></a:graphic></p:graphicFrame>',
      },
    ],
  });

/** The chart data deck with the column chart also on the blank layout. */
function layoutChart(): Uint8Array {
  const files = unzipSync(deckWith());
  const name = Object.keys(files).find(
    (n) =>
      /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(n) &&
      strFromU8(files[n]!).includes('type="blank"')
  )!;
  const slide = strFromU8(files['ppt/slides/slide1.xml']!);
  const frame = /<p:graphicFrame>.*?<\/p:graphicFrame>/s.exec(slide)![0];
  files[name] = strToU8(strFromU8(files[name]!).replace('</p:spTree>', `${frame}</p:spTree>`));
  const rels = `${name.replace('slideLayouts/', 'slideLayouts/_rels/')}.rels`;
  files[rels] = strToU8(
    strFromU8(files[rels]!).replace(
      '</Relationships>',
      '<Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>'
    )
  );
  return zipSync(files);
}

describe('Shape.chart: read', () => {
  test('a column chart reads its type, title, categories, series, and workbook', async () => {
    const runtime = await open(deckWith());
    expect(await readChart(runtime)).toEqual({
      chartType: 'ColumnClustered',
      title: 'Sales by region',
      categories: ['North', 'South', 'East', 'West'],
      series: [
        { name: 'Plan', values: [42, 35, 51, 28] },
        { name: 'Actual', values: [38, 41, 47, 33] },
        { name: 'Forecast', values: [45, 39, 55, 36] },
      ],
      workbook: 'Embedded',
      readOnlyReason: null,
    });
    runtime.dispose();
  });

  test('a scatter chart reads its X values as categories', async () => {
    const runtime = await open(deckWith());
    const read = await readChart(runtime, 1);
    // `c:scatterStyle val="lineMarker"` with lines between the markers.
    expect(read.chartType).toBe('XYScatterLines');
    expect(read.categories).toEqual(['1', '2', '3', '4', '5']);
    expect(read.series[0]!.values).toEqual([2.5, 3.1, 4.8, 5.2, 6.9]);
    runtime.dispose();
  });

  test('a shape that is not a chart has a null chart, and a load of it passes', async () => {
    const runtime = await open(deckWith());
    const [isNull, other] = await runtime.run(async (context) => {
      const title = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      const chart = title.chart.load('series');
      const frame = chartAt(context).chart;
      await context.sync();
      return [chart.isNullObject, frame.isNullObject];
    });
    expect([isNull, other]).toEqual([true, false]);
    runtime.dispose();
  });

  test('the chart of a shape that a sync found missing is a null object, with no refusal', async () => {
    const runtime = await open(deckWith());
    const isNull = await runtime.run(async (context) => {
      const shape = context.presentation.slides.getItemAt(0).shapes.getItemOrNullObject('999');
      await context.sync();
      const chart = shape.chart.load('series');
      await context.sync();
      return [shape.isNullObject, chart.isNullObject];
    });
    expect(isNull).toEqual([true, true]);
    runtime.dispose();
  });

  test('a load through the shape navigates to the chart', async () => {
    const runtime = await open(deckWith());
    const series = await runtime.run(async (context) => {
      const shape = chartAt(context).load('chart/series');
      await context.sync();
      return shape.chart.series.map((s) => s.name);
    });
    expect(series).toEqual(['Plan', 'Actual', 'Forecast']);
    runtime.dispose();
  });

  test('a combination chart reads the series of every plot, as read-only', async () => {
    const runtime = await open(combination());
    expect(await readChart(runtime)).toEqual({
      chartType: null,
      title: 'Sales by region',
      categories: ['North', 'South', 'East', 'West'],
      series: [
        { name: 'Plan', values: [42, 35, 51, 28] },
        { name: 'Actual', values: [38, 41, 47, 33] },
        { name: 'Forecast', values: [45, 39, 55, 36] },
        { name: '', values: [1, null, null, null] },
      ],
      workbook: 'Embedded',
      readOnlyReason: 'Combination',
    });
    runtime.dispose();
  });

  test('a chart part with many plots reads as read-only with no data', async () => {
    const line =
      '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>' +
      '<c:ser><c:idx val="9"/><c:order val="9"/><c:val><c:numLit><c:formatCode>General</c:formatCode>' +
      '<c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>' +
      '<c:axId val="1001"/><c:axId val="1002"/></c:lineChart>';
    const runtime = await open(
      deckWith((parts) => {
        const chart = parts.find((p) => p[0] === 'ppt/charts/chart1.xml')!;
        chart[2] = chart[2].replace('</c:barChart>', `</c:barChart>${line.repeat(20)}`);
      })
    );
    expect(await readChart(runtime)).toMatchObject({
      categories: [],
      series: [],
      readOnlyReason: 'Combination',
    });
    runtime.dispose();
  });

  test('a chart with a hidden series reads its type and reason, and refuses a write', async () => {
    const runtime = await open(filtered());
    expect(await readChart(runtime)).toMatchObject({
      chartType: 'ColumnClustered',
      title: 'Sales by region',
      readOnlyReason: 'Filtered',
    });
    const error = await runtime.run((context) =>
      refusedAtSync(context, () => chartAt(context).chart.setData(NEW_DATA))
    );
    expect(isPptxEditorError(error) && [error.code, error.debugInfo.innerError?.code]).toEqual([
      'NotSupported',
      'edit.chart-data-unsupported',
    ]);
    runtime.dispose();
  });

  test('a frame whose chart part is missing reads as Unreadable, and refuses a write', async () => {
    const runtime = await open(missingPart());
    expect(await readChart(runtime)).toEqual({
      chartType: null,
      title: null,
      categories: [],
      series: [],
      workbook: 'None',
      readOnlyReason: 'Unreadable',
    });
    const error = await runtime.run((context) =>
      refusedAtSync(context, () => chartAt(context).chart.setData(NEW_DATA))
    );
    expect(isPptxEditorError(error) && [error.code, error.debugInfo.innerError?.code]).toEqual([
      'NotSupported',
      'edit.chart-unreadable',
    ]);
    runtime.dispose();
  });

  test('a chart of an extension namespace reads as Unsupported, with a null chart', async () => {
    const runtime = await open(extensionChart());
    const read = await runtime.run(async (context) => {
      const shape = chartAt(context).load('type');
      const chart = shape.chart;
      await context.sync();
      return [shape.type, chart.isNullObject];
    });
    expect(read).toEqual(['Unsupported', true]);
    runtime.dispose();
  });

  test('a chart on a layout reads and takes new data as a slide chart does', async () => {
    const runtime = await open(layoutChart());
    const series = await runtime.run(async (context) => {
      const layouts = context.presentation.slideMasters.getItemAt(0).layouts.load('items/name');
      await context.sync();
      const lists = layouts.items.map((l) => l.shapes.load('items/type'));
      await context.sync();
      const frame = lists.flatMap((l) => l.items).find((s) => s.type === 'Chart')!;
      frame.chart.setData({ categories: ['A'], series: [{ name: 'S', values: [7] }] });
      const chart = frame.chart.load('series');
      await context.sync();
      return chart.series;
    });
    expect(series).toEqual([{ name: 'S', values: [7] }]);
    whole(await runtime.save());
    runtime.dispose();
  });

  test('a chart with no workbook and a linked workbook read their state', async () => {
    const plain = await open(new Uint8Array(readFileSync(join(FIXTURES, 'charts.pptx'))));
    const workbook = await plain.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes.load('items/type');
      await context.sync();
      const chart = shapes.items.find((s) => s.type === 'Chart')!.chart.load('workbook');
      await context.sync();
      return chart.workbook;
    });
    expect(workbook).toBe('None');
    const external = await open(linked());
    expect((await readChart(external)).workbook).toBe('External');
    plain.dispose();
    external.dispose();
  });
});

describe('Chart.setData', () => {
  test('replaces the values in the caches and the workbook, and survives save and reopen', async () => {
    const runtime = await open(deckWith());
    await runtime.run(async (context) => {
      chartAt(context).chart.setData(NEW_DATA);
      await context.sync();
    });
    const pkg = whole(await runtime.save());
    const chart = partText(pkg, '/ppt/charts/chart1.xml');
    expect(chart).toContain('<c:pt idx="3"><c:v>40</c:v></c:pt>');
    // A blank value has no point, and the count stays.
    expect(chart).toMatch(
      /<c:f>Sheet1!\$C\$2:\$C\$5<\/c:f><c:numCache><c:formatCode>General<\/c:formatCode><c:ptCount val="4"\/><c:pt idx="0"><c:v>11<\/c:v><\/c:pt><c:pt idx="2">/
    );
    // The workbook holds the same values as the caches.
    const sheet = workbookEntry(pkg, '/xl/worksheets/sheet1.xml');
    expect(sheet).toContain('<c r="B5"><v>40</v></c>');
    expect(sheet).toContain('<c r="D2"><v>12</v></c>');
    expect(sheet).not.toContain('r="C3"');
    const again = await reopen(runtime);
    const read = await readChart(again);
    expect(read.series).toEqual(NEW_DATA.series);
    expect(read.workbook).toBe('Embedded');
    runtime.dispose();
    again.dispose();
  });

  test('adds a category and a series, and removes them again, in the workbook too', async () => {
    const runtime = await open(deckWith());
    const grown = {
      categories: [...NEW_DATA.categories, 'Central'],
      series: [
        ...NEW_DATA.series.map((s) => ({ ...s, values: [...s.values, 5] })),
        { name: 'Stretch', values: [1, 2, 3, 4, 5] },
      ],
    };
    await runtime.run(async (context) => {
      chartAt(context).chart.setData(grown);
      await context.sync();
    });
    let pkg = whole(await runtime.save());
    expect(workbookEntry(pkg, '/xl/tables/table1.xml')).toContain('ref="A1:E6"');
    expect(workbookEntry(pkg, '/xl/worksheets/sheet1.xml')).toContain('<c r="E6"><v>5</v></c>');
    expect(await readChart(runtime)).toMatchObject({
      categories: grown.categories,
      series: grown.series,
    });
    const shrunk = {
      categories: ['North', 'South'],
      series: [{ name: 'Plan', values: [1, 2] }],
    };
    await runtime.run(async (context) => {
      chartAt(context).chart.setData(shrunk);
      await context.sync();
    });
    pkg = whole(await runtime.save());
    expect(workbookEntry(pkg, '/xl/tables/table1.xml')).toContain('ref="A1:B3"');
    expect(await readChart(runtime)).toMatchObject(shrunk);
    runtime.dispose();
  });

  test('unknown content of the chart part survives a write, save, and reopen', async () => {
    const keep =
      '<c:extLst><c:ext uri="{00000000-0000-0000-0000-00000000000A}">' +
      '<x:keep xmlns:x="urn:example:keep" note="stays"/></c:ext></c:extLst>';
    const runtime = await open(
      deckWith((parts) => {
        const chart = parts.find((p) => p[0] === 'ppt/charts/chart1.xml')!;
        chart[2] = chart[2].replace('</c:chartSpace>', `${keep}</c:chartSpace>`);
      })
    );
    await runtime.run(async (context) => {
      chartAt(context).chart.setData(NEW_DATA);
      await context.sync();
    });
    const again = await reopen(runtime);
    const pkg = whole(await again.save());
    expect(partText(pkg, '/ppt/charts/chart1.xml')).toContain('note="stays"');
    expect((await readChart(again)).series).toEqual(NEW_DATA.series);
    runtime.dispose();
    again.dispose();
  });

  test('a new table that the embedded workbook cannot take refuses with NotSupported', async () => {
    const runtime = await open(deckWith());
    // A workbook table needs a different name for each column.
    const error = await runtime.run((context) =>
      refusedAtSync(context, () =>
        chartAt(context).chart.setData({
          categories: NEW_DATA.categories,
          series: [...NEW_DATA.series, { name: 'Plan', values: [1, 2, 3, 4] }],
        })
      )
    );
    expect(isPptxEditorError(error) && [error.code, error.debugInfo.innerError?.code]).toEqual([
      'NotSupported',
      'edit.chart-data-layout',
    ]);
    expect((await readChart(runtime)).series).toHaveLength(3);
    runtime.dispose();
  });

  test('a scatter chart takes numeric X values, and refuses text X values', async () => {
    const runtime = await open(deckWith());
    await runtime.run(async (context) => {
      chartAt(context, 1).chart.setData({
        categories: ['0.5', '1', '2'],
        series: [{ name: 'Y', values: [1, 2, 3] }],
      });
      await context.sync();
    });
    expect((await readChart(runtime, 1)).categories).toEqual(['0.5', '1', '2']);
    const error = await runtime.run((context) =>
      refusedAtSync(context, () =>
        chartAt(context, 1).chart.setData({
          categories: ['a'],
          series: [{ name: 'Y', values: [1] }],
        })
      )
    );
    expect(isPptxEditorError(error) && [error.code, error.debugInfo.innerError?.code]).toEqual([
      'InvalidArgument',
      'edit.chart-data-value',
    ]);
    runtime.dispose();
  });

  test('a chart with no workbook and a linked workbook write the caches only', async () => {
    const external = await open(linked());
    await external.run(async (context) => {
      chartAt(context).chart.setData(NEW_DATA);
      await context.sync();
    });
    const pkg = whole(await external.save());
    // The linked file is never opened: its automatic update is turned off.
    expect(partText(pkg, '/ppt/charts/chart1.xml')).toContain('<c:autoUpdate val="0"/>');
    expect((await readChart(external)).series).toEqual(NEW_DATA.series);
    const plain = await open(new Uint8Array(readFileSync(join(FIXTURES, 'charts.pptx'))));
    const before = await plain.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes.load('items/type');
      await context.sync();
      const chart = shapes.items.find((s) => s.type === 'Chart')!.chart.load('categories,series');
      await context.sync();
      chart.setData({
        categories: chart.categories,
        series: chart.series.map((s) => ({ ...s, values: s.values.map((v) => (v ?? 0) + 1) })),
      });
      await context.sync();
      chart.load('series');
      await context.sync();
      return chart.series;
    });
    expect(before.length).toBeGreaterThan(0);
    whole(await plain.save());
    external.dispose();
    plain.dispose();
  });

  test('data of the wrong form refuses with InvalidArgument at the sync, and writes nothing', async () => {
    const runtime = await open(deckWith());
    const cases: unknown[] = [
      null,
      { categories: 'North', series: [] },
      { categories: ['a', 'b'], series: [{ name: 'S', values: [1] }] },
      { categories: ['a'], series: [{ name: 'S', values: [Number.NaN] }] },
      { categories: ['a'], series: [{ name: 3, values: [1] }] },
    ];
    for (const data of cases) {
      const error = await runtime.run((context) =>
        refusedAtSync(context, () => {
          chartAt(context).chart.setData(data as never);
        })
      );
      expect(isPptxEditorError(error) && error.code).toBe('InvalidArgument');
    }
    // The engine checks the counts and the text.
    for (const [data, diagnostic] of [
      [{ categories: [], series: [] }, 'edit.chart-data-range'],
      [{ categories: ['a\u0001'], series: [{ name: 'S', values: [1] }] }, 'edit.chart-data-text'],
      [
        { categories: ['a'], series: [{ name: 'x'.repeat(256), values: [1] }] },
        'edit.chart-data-text',
      ],
    ] as const) {
      const error = await runtime.run((context) =>
        refusedAtSync(context, () => chartAt(context).chart.setData(data as never))
      );
      expect(isPptxEditorError(error) && [error.code, error.debugInfo.innerError?.code]).toEqual([
        'InvalidArgument',
        diagnostic,
      ]);
    }
    expect((await readChart(runtime)).series[0]!.values).toEqual([42, 35, 51, 28]);
    runtime.dispose();
  });

  test('a refused setData leaves the other writes of the sync unapplied', async () => {
    const runtime = await open(deckWith());
    const error = await runtime.run((context) =>
      refusedAtSync(context, () => {
        const shape = chartAt(context);
        shape.name = 'Renamed';
        shape.chart.setData({ categories: [], series: [] });
      })
    );
    expect(isPptxEditorError(error) && error.code).toBe('InvalidArgument');
    const name = await runtime.run(async (context) => {
      const shape = chartAt(context).load('name');
      await context.sync();
      return shape.name;
    });
    expect(name).toBe('Chart 1');
    runtime.dispose();
  });

  test('a sparse array refuses with InvalidArgument', async () => {
    const runtime = await open(deckWith());
    const categories: string[] = ['a', 'b', 'c'];
    delete categories[1];
    const error = await runtime.run((context) =>
      refusedAtSync(context, () =>
        chartAt(context).chart.setData({ categories, series: [{ name: 'S', values: [1, 2, 3] }] })
      )
    );
    expect(isPptxEditorError(error) && error.code).toBe('InvalidArgument');
    runtime.dispose();
  });

  test('a read-only chart refuses with NotSupported, and a shape that is not a chart refuses', async () => {
    const runtime = await open(combination());
    const error = await runtime.run((context) =>
      refusedAtSync(context, () => chartAt(context).chart.setData(NEW_DATA))
    );
    expect(isPptxEditorError(error) && [error.code, error.debugInfo.innerError?.code]).toEqual([
      'NotSupported',
      'edit.chart-data-unsupported',
    ]);
    const other = await runtime.run((context) =>
      refusedAtSync(context, () =>
        context.presentation.slides.getItemAt(0).shapes.getItem('2').chart.setData(NEW_DATA)
      )
    );
    expect(isPptxEditorError(other) && [other.code, other.debugInfo.innerError?.code]).toEqual([
      'NotSupported',
      'edit.not-a-chart',
    ]);
    // After a sync that found the null object, a write refuses at once.
    const known = await runtime.run(async (context) => {
      const chart = context.presentation.slides.getItemAt(0).shapes.getItem('2').chart;
      await context.sync();
      return refusedAtSync(context, () => chart.setData(NEW_DATA)).catch((e: unknown) => e);
    });
    expect(isPptxEditorError(known) && [known.code, known.debugInfo.innerError?.code]).toEqual([
      'InvalidObjectPath',
      'edit.null-object',
    ]);
    runtime.dispose();
  });

  test('the browser runtime reads and writes as the server runtime does, as one undo step', async () => {
    const live = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await live.load(deckWith())).ok).toBe(true);
    const runtime = BrowserEditor.createBrowser(live);
    const server = await open(deckWith());
    expect(await readChart(runtime)).toEqual(await readChart(server));
    const before = live.save()!;
    for (const r of [runtime, server])
      await r.run(async (context) => {
        chartAt(context).chart.setData(NEW_DATA);
        await context.sync();
      });
    expect(await readChart(runtime)).toEqual(await readChart(server));
    const [livePkg, serverPkg] = [whole(live.save()!), whole(await server.save())];
    for (const name of ['/ppt/charts/chart1.xml'])
      expect(partText(livePkg, name)).toBe(partText(serverPkg, name));
    expect(workbookEntry(livePkg, '/xl/worksheets/sheet1.xml')).toBe(
      workbookEntry(serverPkg, '/xl/worksheets/sheet1.xml')
    );
    expect(live.exec({ type: 'undo' }).ok).toBe(true);
    expect(live.snapshot().canUndo).toBe(false);
    const undone = readPptx(live.save()!);
    const original = readPptx(before);
    if (!undone.ok || !original.ok) throw new Error('reopen');
    expect(partText(undone.pkg, '/ppt/charts/chart1.xml')).toBe(
      partText(original.pkg, '/ppt/charts/chart1.xml')
    );
    server.dispose();
  });

  test('a write sends the lookups, the null-object check, and setChartData in one batch', async () => {
    // The collaboration test (`packages/pro/src/__tests__/yjs-chart-data.test.ts`) sends these
    // operations to two linked editors; this test keeps them equal to what `setData` sends.
    const live = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await live.load(deckWith())).ok).toBe(true);
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
      chartAt(context).chart.setData(NEW_DATA);
      await context.sync();
    });
    // The write names the shape, so it waits for no read: one batch.
    expect(batches).toEqual([
      [
        { op: 'getPresentation' },
        { op: 'getSlide', key: { index: 0 } },
        { op: 'getShape', container: { result: 1 }, key: { id: '3' } },
        { op: 'readShape', shape: { result: 2 } },
        { op: 'setChartData', shape: { result: 2 }, data: NEW_DATA },
      ],
    ]);
  });
});
