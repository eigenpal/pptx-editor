// Chart data (`chartData`, `setChartData`): the read of a chart's categories and series, the
// caches, point counts, and references an edit writes, the embedded workbook's cells and table,
// the formatting that follows its series and points, undo, refusals, and save and reopen.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { binaryPartOf, readPptx, readZip, serializeXmlPart } from '../../store/index.ts';
import type { ChartData } from '../../contracts/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import {
  chartDraftCellErrors,
  chartDraftCommand,
  chartDraftFrom,
  setChartDraftCell,
} from '../chart-data-fields.ts';

const FIXTURES = join(import.meta.dir, '../../../../../fixtures');
const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(join(FIXTURES, `${name}.pptx`)));

async function open(bytes: Uint8Array, slide = 0): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  e.goToSlide(slide);
  const chart = e.snapshot().shapes!.find((s) => s.name.startsWith('Chart'))!;
  e.setSelection({ kind: 'shapes', nodeIds: [chart.nodeId] });
  return e;
}

/** The fixture with some of its entries replaced. */
function patched(name: string, edit: (files: Record<string, Uint8Array>) => void): Uint8Array {
  const files = unzipSync(fixture(name));
  edit(files);
  return zipSync(files);
}

const pkgOf = (e: PptxEditorInstance) => e.session!.store.package;

function partText(e: PptxEditorInstance, name: string): string {
  const part = e.session!.store.package.parts.get(name);
  if (!part) throw new Error(`no part ${name}`);
  return part.kind === 'xml' ? serializeXmlPart(part.xml) : new TextDecoder().decode(part.bytes);
}

/** An entry of an embedded workbook, as text. */
function workbookEntry(e: PptxEditorInstance, part: string, entry: string): string {
  const zip = readZip(binaryPartOf(e.session!.store.package, part)!);
  if (!zip.ok) throw new Error('the workbook does not read');
  return strFromU8(zip.entries.get(entry)!);
}

function read(e: PptxEditorInstance): ChartData {
  const r = e.chartData();
  if (!r.ok) throw new Error(r.reason);
  return r.data;
}

describe('chart data: read', () => {
  test('reads the categories, series, and workbook of a column chart', async () => {
    const e = await open(fixture('chart-data'));
    expect(e.snapshot().chart).toEqual({ nodeId: expect.any(String), editable: true });
    const data = read(e);
    expect(data.kind).toBe('category');
    expect(data.workbook).toBe('embedded');
    expect(data.categories.map((c) => c.label)).toEqual(['North', 'South', 'East', 'West']);
    expect(data.series.map((s) => s.name)).toEqual(['Plan', 'Actual', 'Forecast']);
    expect(data.series[1]!.values).toEqual([38, 41, 47, 33]);
  });

  test('reads a scatter chart as X values and Y series', async () => {
    const data = read(await open(fixture('chart-data'), 1));
    expect(data.kind).toBe('xy');
    expect(data.categories.map((c) => c.label)).toEqual(['1', '2', '3', '4', '5']);
    expect(data.series[0]!.values).toEqual([2.5, 3.1, 4.8, 5.2, 6.9]);
  });

  test('reports a chart with no workbook, and no chart when a shape is selected', async () => {
    const e = await open(fixture('charts'));
    expect(read(e).workbook).toBe('none');
    const title = e.snapshot().shapes!.find((s) => s.name === 'Shape 2')!;
    e.setSelection({ kind: 'shapes', nodeIds: [title.nodeId] });
    expect(e.snapshot().chart).toBeNull();
    expect(e.chartData()).toEqual({ ok: false, reason: 'errors.chartData.noChart' });
  });

  test('a bubble chart is not editable, with a reason', async () => {
    const bytes = patched('charts', (f) => {
      const xml = strFromU8(f['ppt/charts/chart1.xml']!);
      f['ppt/charts/chart1.xml'] = strToU8(xml.replaceAll('c:barChart>', 'c:bubbleChart>'));
    });
    const e = await open(bytes);
    expect(e.snapshot().chart).toMatchObject({
      editable: false,
      reason: 'errors.chartData.unsupportedType',
    });
    expect(
      e.exec({ type: 'setChartData', data: { kind: 'category', categories: [], series: [] } })
    ).toMatchObject({ ok: false });
  });
});

describe('chart data: edit', () => {
  test('edits values, renames, and adds and removes a series and a category', async () => {
    const e = await open(fixture('chart-data'));
    const d = read(e);
    const [plan, actual] = d.series;
    const result = e.exec({
      type: 'setChartData',
      data: {
        kind: d.kind,
        // South goes, Central is new.
        categories: [d.categories[0]!, d.categories[2]!, d.categories[3]!, { label: 'Central' }],
        series: [
          { ...plan!, name: 'Budget', values: [1, 3, 4, 9] },
          { ...actual!, values: [38, 47, 33, null] },
          { name: 'Stretch', values: [5, 6, 7, 8] },
        ],
      },
    });
    expect(result.ok).toBe(true);
    const chart = partText(e, '/ppt/charts/chart1.xml');
    expect(chart).toContain('<c:f>Sheet1!$A$2:$A$5</c:f>');
    expect(chart).toContain('<c:v>Budget</c:v>');
    expect(chart).toContain(
      '<c:f>Sheet1!$D$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Stretch</c:v>'
    );
    expect(chart).not.toContain('Forecast');
    expect(chart).not.toContain('South');
    // The West point keeps its own fill and label: it moved from index 3 to 2.
    expect(chart).toContain('<c:dPt><c:idx val="2"/>');
    expect(chart).toContain('<c:dLbl><c:idx val="2"/>');
    // A blank value has no point; the count stays.
    expect(chart).toMatch(
      /<c:f>Sheet1!\$C\$2:\$C\$5<\/c:f><c:numCache><c:formatCode>General<\/c:formatCode><c:ptCount val="4"\/>(<c:pt idx="[012]"><c:v>\d+<\/c:v><\/c:pt>){3}<\/c:numCache>/
    );
    // The new series has its own index and no formatting of its own.
    expect(chart).toContain('<c:ser><c:idx val="3"/><c:order val="2"/><c:tx>');
    const sheet = workbookEntry(e, '/ppt/embeddings/Workbook1.xlsx', '/xl/worksheets/sheet1.xml');
    expect(sheet).toContain('<c r="B1" t="inlineStr"><is><t>Budget</t></is></c>');
    expect(sheet).toContain('<c r="A5" t="inlineStr"><is><t>Central</t></is></c>');
    expect(sheet).toContain('<c r="D5"><v>8</v></c>');
    expect(sheet).not.toContain('r="C5"');
    const table = workbookEntry(e, '/ppt/embeddings/Workbook1.xlsx', '/xl/tables/table1.xml');
    expect(table).toContain('ref="A1:D5"');
    expect(table).toContain('<tableColumn id="2" name="Budget"/>');
    expect(table).toContain('<tableColumn id="4" name="Stretch"/>');
    // Undo brings back the chart part and the workbook as they were.
    const before = await open(fixture('chart-data'));
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(partText(e, '/ppt/charts/chart1.xml')).toBe(partText(before, '/ppt/charts/chart1.xml'));
    expect(binaryPartOf(pkgOf(e), '/ppt/embeddings/Workbook1.xlsx')).toEqual(
      binaryPartOf(pkgOf(before), '/ppt/embeddings/Workbook1.xlsx')!
    );
  });

  test('a removed series in the middle moves the next one up in the workbook table', async () => {
    const e = await open(fixture('chart-data'));
    const d = read(e);
    expect(
      e.exec({
        type: 'setChartData',
        data: { kind: d.kind, categories: d.categories, series: [d.series[0]!, d.series[2]!] },
      }).ok
    ).toBe(true);
    const chart = partText(e, '/ppt/charts/chart1.xml');
    expect(chart).toContain(
      '<c:f>Sheet1!$C$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Forecast'
    );
    const sheet = workbookEntry(e, '/ppt/embeddings/Workbook1.xlsx', '/xl/worksheets/sheet1.xml');
    expect(sheet).toContain('<dimension ref="A1:C5"/>');
    expect(sheet).not.toContain('r="D1"');
    const table = workbookEntry(e, '/ppt/embeddings/Workbook1.xlsx', '/xl/tables/table1.xml');
    expect(table).toContain('ref="A1:C5"');
    expect(table).toContain('<tableColumns count="3">');
  });

  test('edits a scatter chart: X values, a new point, and a new series', async () => {
    const e = await open(fixture('chart-data'), 1);
    const d = read(e);
    expect(
      e.exec({
        type: 'setChartData',
        data: {
          kind: 'xy',
          categories: [...d.categories, { label: '6.5' }],
          series: [
            { ...d.series[0]!, values: [...d.series[0]!.values, 7.5] },
            { ...d.series[1]!, values: [...d.series[1]!.values, null] },
            { name: 'Sample C', values: [1, 2, 3, 4, 5, 6] },
          ],
        },
      }).ok
    ).toBe(true);
    const chart = partText(e, '/ppt/charts/chart2.xml');
    expect(chart).toContain('<c:xVal><c:numRef><c:f>Sheet1!$A$2:$A$7</c:f>');
    expect(chart).toContain('<c:yVal><c:numRef><c:f>Sheet1!$D$2:$D$7</c:f>');
    expect(chart).toContain('<c:pt idx="5"><c:v>6.5</c:v></c:pt>');
    const sheet = workbookEntry(e, '/ppt/embeddings/Workbook2.xlsx', '/xl/worksheets/sheet1.xml');
    expect(sheet).toContain('<c r="A7"><v>6.5</v></c>');
    expect(sheet).toContain('<c r="D1" t="inlineStr"><is><t>Sample C</t></is></c>');
    // A scatter chart's X values are numbers.
    expect(
      e.exec({
        type: 'setChartData',
        data: { ...read(e), categories: read(e).categories.map((c) => ({ ...c, label: 'x' })) },
      })
    ).toMatchObject({ ok: false, reason: 'errors.chartData.notNumber' });
  });

  test('a chart with no workbook gets new references in the same sheet table', async () => {
    const e = await open(fixture('charts'));
    const d = read(e);
    expect(
      e.exec({
        type: 'setChartData',
        data: {
          kind: d.kind,
          categories: d.categories.slice(0, 2),
          series: d.series.slice(0, 1).map((s) => ({ ...s, values: s.values.slice(0, 2) })),
        },
      }).ok
    ).toBe(true);
    const chart = partText(e, '/ppt/charts/chart1.xml');
    expect(chart).toContain('<c:f>Sheet1!$A$2:$A$3</c:f>');
    expect(chart).toContain('<c:ptCount val="2"/>');
  });

  test('literal data stays literal', async () => {
    const e = await open(fixture('slide-structure'));
    const d = read(e);
    expect(d.workbook).toBe('kept');
    expect(
      e.exec({
        type: 'setChartData',
        data: {
          kind: d.kind,
          categories: [...d.categories, { label: 'C' }],
          series: [{ ...d.series[0]!, values: [1, 2, 3] }],
        },
      }).ok
    ).toBe(true);
    const chart = partText(e, '/ppt/charts/chart1.xml');
    expect(chart).toContain('<c:strLit><c:ptCount val="3"/>');
    expect(chart).toContain('<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="3"/>');
  });

  test('a linked workbook turns off its automatic update; a formula cell keeps the workbook', async () => {
    const linked = patched('chart-data', (f) => {
      f['ppt/charts/_rels/chart1.xml.rels'] = strToU8(
        '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="file:///C:/data.xlsx" TargetMode="External"/></Relationships>'
      );
      const xml = strFromU8(f['ppt/charts/chart1.xml']!);
      f['ppt/charts/chart1.xml'] = strToU8(
        xml.replace('<c:autoUpdate val="0"/>', '<c:autoUpdate val="1"/>')
      );
    });
    const e = await open(linked);
    const d = read(e);
    expect(d.workbook).toBe('external');
    expect(
      e.exec({
        type: 'setChartData',
        data: { ...d, series: [{ ...d.series[0]!, name: 'X' }, ...d.series.slice(1)] },
      }).ok
    ).toBe(true);
    expect(partText(e, '/ppt/charts/chart1.xml')).toContain('<c:autoUpdate val="0"/>');

    const formula = patched('chart-data', (f) => {
      const files = unzipSync(f['ppt/embeddings/Workbook1.xlsx']!);
      const sheet = strFromU8(files['xl/worksheets/sheet1.xml']!);
      files['xl/worksheets/sheet1.xml'] = strToU8(
        sheet.replace('<c r="B2"><v>42</v></c>', '<c r="B2"><f>40+2</f><v>42</v></c>')
      );
      f['ppt/embeddings/Workbook1.xlsx'] = zipSync(files);
    });
    const g = await open(formula);
    expect(read(g).workbook).toBe('kept');
    const before = binaryPartOf(g.session!.store.package, '/ppt/embeddings/Workbook1.xlsx');
    const gd = read(g);
    expect(
      g.exec({
        type: 'setChartData',
        data: {
          ...gd,
          series: [{ ...gd.series[0]!, values: [1, 2, 3, 4] }, ...gd.series.slice(1)],
        },
      }).ok
    ).toBe(true);
    expect(binaryPartOf(g.session!.store.package, '/ppt/embeddings/Workbook1.xlsx')).toBe(before);
    expect(partText(g, '/ppt/charts/chart1.xml')).toContain('<c:pt idx="3"><c:v>4</c:v></c:pt>');
  });

  test('refuses data that does not fit the chart', async () => {
    const e = await open(fixture('chart-data'));
    const d = read(e);
    const run = (data: unknown) => e.exec({ type: 'setChartData', data } as never);
    expect(run({ ...d, series: [] })).toMatchObject({
      ok: false,
      reason: 'errors.chartData.empty',
    });
    expect(run({ ...d, kind: 'xy' })).toMatchObject({ ok: false });
    expect(run({ ...d, series: [{ name: 'A', values: [1] }] })).toMatchObject({ ok: false });
    expect(run({ ...d, series: [{ name: 'A\u0001', values: [1, 2, 3, 4] }] })).toMatchObject({
      ok: false,
      reason: 'errors.chartData.invalidText',
    });
    expect(run({ ...d, series: [{ name: 'A', values: [1, 2, 3, Infinity] }] })).toMatchObject({
      ok: false,
      reason: 'errors.chartData.notNumber',
    });
    expect(run({ ...d, series: [d.series[0], d.series[0]] })).toMatchObject({ ok: false });
    expect(e.snapshot().canUndo).toBe(false);
  });

  test('the edit survives save and reopen', async () => {
    const e = await open(fixture('chart-data'));
    const d = read(e);
    const data = {
      kind: d.kind,
      categories: d.categories,
      series: [{ ...d.series[0]!, name: 'Renamed', values: [9, 8, 7, 6] }],
    };
    expect(e.exec({ type: 'setChartData', data }).ok).toBe(true);
    const saved = e.save()!;
    expect(readPptx(saved).ok).toBe(true);
    const again = read(await open(saved));
    expect(again.series).toEqual([{ name: 'Renamed', values: [9, 8, 7, 6], source: 0 }]);
    expect(again.workbook).toBe('embedded');
  });
});

describe('chart data: the workbook and the chart stay in step', () => {
  const SHEET = '/xl/worksheets/sheet1.xml';
  const BOOK = '/ppt/embeddings/Workbook1.xlsx';

  test('a new series never overwrites a cell beside the table', async () => {
    const deck = patched('chart-data', (f) => {
      const files = unzipSync(f['ppt/embeddings/Workbook1.xlsx']!);
      const sheet = strFromU8(files['xl/worksheets/sheet1.xml']!);
      files['xl/worksheets/sheet1.xml'] = strToU8(
        sheet.replace(
          /(<row r="2"[^>]*>.*?)(<\/row>)/,
          '$1<c r="E2" t="inlineStr"><is><t>KEEP</t></is></c>$2'
        )
      );
      f['ppt/embeddings/Workbook1.xlsx'] = zipSync(files);
    });
    const e = await open(deck);
    const d = read(e);
    expect(d.workbook).toBe('embedded');
    const before = binaryPartOf(pkgOf(e), BOOK);
    const result = e.exec({
      type: 'setChartData',
      data: { ...d, series: [...d.series, { name: 'New', values: [1, 1, 1, 1] }] },
    });
    // The workbook took the data before the edit, so the edit is refused, not split.
    expect(result).toMatchObject({ ok: false, reason: 'errors.chartData.workbook' });
    expect(binaryPartOf(pkgOf(e), BOOK)).toBe(before);
  });

  test('series names that a table cannot hold are refused while the workbook is in step', async () => {
    const e = await open(fixture('chart-data'));
    const d = read(e);
    const result = e.exec({
      type: 'setChartData',
      data: { ...d, series: d.series.map((s) => ({ ...s, name: 'Same' })) },
    });
    expect(result).toMatchObject({ ok: false, reason: 'errors.chartData.workbook' });
    expect(workbookEntry(e, BOOK, SHEET)).not.toContain('Same');
  });

  test('a scatter chart with no X values gets numeric X values', async () => {
    const deck = patched('chart-data', (f) => {
      const xml = strFromU8(f['ppt/charts/chart2.xml']!);
      f['ppt/charts/chart2.xml'] = strToU8(xml.replace(/<c:xVal>.*?<\/c:xVal>/gs, ''));
    });
    const e = await open(deck, 1);
    const d = read(e);
    const categories = d.categories.map((c, i) => ({ ...c, label: String((i + 1) * 10) }));
    expect(e.exec({ type: 'setChartData', data: { ...d, categories } }).ok).toBe(true);
    const chart = partText(e, '/ppt/charts/chart2.xml');
    expect(chart).toMatch(/<c:xVal><c:num(Lit|Ref)>/);
    expect(chart).toContain('<c:v>50</c:v>');
  });

  test('the dialog checks X values only when they are numbers', () => {
    const text = chartDraftFrom({
      kind: 'xy',
      categories: [{ label: 'a', source: 0 }],
      series: [{ name: 's', values: [1], source: 0 }],
      workbook: 'none',
    });
    expect([...chartDraftCellErrors(text.grid)]).toEqual([]);
    expect(chartDraftCommand(text)).toMatchObject({
      data: { categories: [{ label: 'a', source: 0 }] },
    });
    const numbers = chartDraftFrom({
      kind: 'xy',
      categories: [{ label: '1', source: 0 }],
      series: [{ name: 's', values: [1], source: 0 }],
      workbook: 'none',
    });
    const typed = setChartDraftCell(numbers.grid, 1, 0, 'x');
    expect(chartDraftCellErrors(typed).get('1,0')).toBe('chartData.errors.xNumber');
  });
});
