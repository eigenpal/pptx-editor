// New charts and chart formatting (`insertChart`, `setChartType`, `setChartFormat`,
// `setChartColor`, `chartFormat`): the parts a new chart adds, its workbook, a data edit on a
// new chart, each type change, the title, legend, labels, gridlines, axes, and colors, undo,
// refusals, and save and reopen.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8 } from 'fflate';
import { binaryPartOf, readPptx, readZip } from '../../store/index.ts';
import type { ChartFormat, ChartType } from '../../contracts/index.ts';
import { CHART_TYPES } from '../../automation/commands/chart-type.ts';
import {
  chartSeriesNote,
  chartTypeChangeNote,
  createPptxEditor,
  listenPaneRequests,
  runToolbarCommand,
  type PptxEditorInstance,
} from '../index.ts';

const FIXTURES = join(import.meta.dir, '../../../../../fixtures');
const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(join(FIXTURES, `${name}.pptx`)));

async function open(bytes: Uint8Array, slide = 0): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  e.goToSlide(slide);
  return e;
}

/** Select the first chart frame of the current slide. */
function selectChart(e: PptxEditorInstance): string {
  const chart = e.snapshot().shapes!.find((s) => s.name.startsWith('Chart'))!;
  e.setSelection({ kind: 'shapes', nodeIds: [chart.nodeId] });
  return chart.nodeId;
}

function format(e: PptxEditorInstance): ChartFormat {
  const r = e.chartFormat();
  if (!r.ok) throw new Error(r.reason);
  return r.format;
}

/** The text of the chart part behind the selected chart. */
function chartXml(e: PptxEditorInstance): string {
  const pkg = e.session!.store.package;
  const names = [...pkg.parts.keys()].filter((n) => /^\/ppt\/charts\/chart\d+\.xml$/.test(n));
  const part = pkg.parts.get(names.at(-1)!)!;
  return part.kind === 'xml' ? '' : new TextDecoder().decode(part.bytes);
}

async function inserted(type: ChartType = 'columnClustered'): Promise<PptxEditorInstance> {
  const e = await open(fixture('charts'));
  e.setSelection({ kind: 'none' });
  const r = e.exec({ type: 'insertChart', chartType: type });
  expect(r).toMatchObject({ ok: true });
  return e;
}

describe('insertChart', () => {
  test('adds a chart part, its workbook, and a selected frame, as one undo step', async () => {
    const e = await open(fixture('charts'));
    const before = e.snapshot().shapes!.length;
    const parts = e.session!.store.package.parts.size;
    expect(e.exec({ type: 'insertChart', chartType: 'columnClustered' }).ok).toBe(true);
    const snap = e.snapshot();
    expect(snap.shapes!.length).toBe(before + 1);
    expect(snap.chart).toMatchObject({ editable: true });
    const data = e.chartData();
    if (!data.ok) throw new Error(data.reason);
    expect(data.data.workbook).toBe('embedded');
    expect(data.data.categories.map((c) => c.label)).toEqual([
      'Category 1',
      'Category 2',
      'Category 3',
      'Category 4',
    ]);
    expect(data.data.series.map((s) => s.name)).toEqual(['Series 1', 'Series 2', 'Series 3']);
    expect(format(e)).toMatchObject({ type: 'columnClustered', title: null, legend: 'bottom' });
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(e.snapshot().shapes!.length).toBe(before);
    expect(e.session!.store.package.parts.size).toBe(parts);
  });

  test('writes a workbook whose table the data edit keeps in agreement', async () => {
    const e = await inserted();
    const data = e.chartData();
    if (!data.ok) throw new Error(data.reason);
    const next = {
      kind: data.data.kind,
      categories: data.data.categories,
      series: data.data.series.map((s, i) => (i === 0 ? { ...s, values: [9, 8, 7, 6] } : s)),
    };
    expect(e.exec({ type: 'setChartData', data: next }).ok).toBe(true);
    const pkg = e.session!.store.package;
    const book = [...pkg.parts.keys()].find((n) => n.endsWith('.xlsx'))!;
    const zip = readZip(binaryPartOf(pkg, book)!);
    if (!zip.ok) throw new Error('workbook');
    const sheet = strFromU8(zip.entries.get('/xl/worksheets/sheet1.xml')!);
    expect(sheet).toContain('<c r="B2"><v>9</v></c>');
    const again = e.chartData();
    expect(again.ok && again.data.workbook).toBe('embedded');
  });

  test('takes caller data, and refuses data that does not fit', async () => {
    const e = await open(fixture('charts'));
    const data = { categories: ['A', 'B'], series: [{ name: 'S', values: [1, null] }] };
    expect(e.exec({ type: 'insertChart', chartType: 'pie', data }).ok).toBe(true);
    const read = e.chartData();
    expect(read.ok && read.data.series[0]!.values).toEqual([1, null]);
    expect(
      e.exec({ type: 'insertChart', chartType: 'line', data: { categories: [], series: [] } })
    ).toMatchObject({ ok: false, reason: 'errors.chartData.empty' });
    expect(
      e.exec({
        type: 'insertChart',
        chartType: 'scatter',
        data: { categories: ['x'], series: [{ name: 'S', values: [1] }] },
      })
    ).toMatchObject({ ok: false, reason: 'errors.chartData.notNumber' });
    expect(e.exec({ type: 'insertChart', chartType: 'cone' as ChartType })).toMatchObject({
      ok: false,
    });
  });

  test('every type inserts, reads back as itself, and survives save and reopen', async () => {
    const e = await open(fixture('charts'));
    for (const type of CHART_TYPES) {
      e.setSelection({ kind: 'none' });
      expect(e.exec({ type: 'insertChart', chartType: type }).ok).toBe(true);
      expect(format(e).type).toBe(type);
    }
    const saved = e.save()!;
    expect(readPptx(saved).ok).toBe(true);
    const again = await open(saved);
    // The shape list is top first: the last chart inserted comes first.
    const charts = again.snapshot().shapes!.filter((s) => s.name.startsWith('Chart'));
    const types = charts
      .slice(0, CHART_TYPES.length)
      .reverse()
      .map((c) => {
        again.setSelection({ kind: 'shapes', nodeIds: [c.nodeId] });
        return format(again).type;
      });
    expect(types).toEqual([...CHART_TYPES]);
  });
});

describe('setChartType', () => {
  test('changes the plot and keeps the data, names, and series colors', async () => {
    const e = await open(fixture('chart-data'));
    selectChart(e);
    const before = e.chartData();
    expect(e.exec({ type: 'setChartType', chartType: 'lineMarkers' }).ok).toBe(true);
    expect(format(e).type).toBe('lineMarkers');
    const after = e.chartData();
    if (!before.ok || !after.ok) throw new Error('read');
    expect(after.data.series.map((s) => [s.name, s.values])).toEqual(
      before.data.series.map((s) => [s.name, s.values])
    );
    expect(format(e).colors.map((c) => c.color)).toEqual(['#4472C4', '#ED7D31', '#A5A5A5']);
    // The workbook still takes edits after the change.
    expect(after.data.workbook).toBe('embedded');
  });

  test('goes through every type and back, with axes only where the type has them', async () => {
    const e = await open(fixture('chart-data'));
    selectChart(e);
    for (const type of CHART_TYPES) {
      if (type === 'columnClustered') continue;
      expect(e.exec({ type: 'setChartType', chartType: type }).ok).toBe(true);
      const f = format(e);
      expect(f.type).toBe(type);
      expect(f.axes === null).toBe(type === 'pie' || type === 'doughnut');
      expect(f.colorsBy).toBe(type === 'pie' || type === 'doughnut' ? 'point' : 'series');
      const xml = chartXml(e);
      expect(xml).not.toContain('<c:view3D');
    }
    expect(e.exec({ type: 'setChartType', chartType: 'columnClustered' }).ok).toBe(true);
    expect(readPptx(e.save()!).ok).toBe(true);
  });

  test('refuses the type the chart has, and turns columns into bars with turned axes', async () => {
    const e = await open(fixture('chart-data'));
    selectChart(e);
    expect(e.exec({ type: 'setChartType', chartType: 'columnClustered' })).toMatchObject({
      ok: false,
      reason: 'errors.chartFormat.sameType',
    });
    expect(e.exec({ type: 'setChartType', chartType: 'barStacked' }).ok).toBe(true);
    const xml = chartXml(e);
    expect(xml).toContain('<c:barDir val="bar"/>');
    expect(xml).toMatch(/<c:catAx><c:axId val="1001"\/>.*?<c:axPos val="l"\/>/);
  });
});

describe('setChartFormat', () => {
  test('sets and removes the title, keeping its run formatting', async () => {
    const e = await open(fixture('chart-data'));
    selectChart(e);
    expect(format(e).title).toBe('Sales by region');
    expect(e.exec({ type: 'setChartFormat', format: { title: 'Revenue' } }).ok).toBe(true);
    expect(format(e).title).toBe('Revenue');
    expect(chartXml(e)).toContain('<a:rPr lang="en-US" sz="1862" b="0">');
    expect(e.exec({ type: 'setChartFormat', format: { title: null } }).ok).toBe(true);
    expect(format(e).title).toBeNull();
    expect(chartXml(e)).toContain('<c:autoTitleDeleted val="1"/>');
    expect(e.exec({ type: 'setChartFormat', format: { title: 'a & <b>' } }).ok).toBe(true);
    expect(format(e).title).toBe('a & <b>');
  });

  test('moves and removes the legend, and turns labels, gridlines, and axes on and off', async () => {
    const e = await open(fixture('chart-data'));
    selectChart(e);
    const set = (f: object) => e.exec({ type: 'setChartFormat', format: f });
    expect(set({ legend: 'right', dataLabels: true, gridlines: false, axes: false }).ok).toBe(true);
    expect(format(e)).toMatchObject({
      legend: 'right',
      dataLabels: true,
      gridlines: false,
      axes: false,
    });
    expect(set({ legend: 'none', dataLabels: false, gridlines: true, axes: true }).ok).toBe(true);
    expect(format(e)).toMatchObject({
      legend: 'none',
      dataLabels: false,
      gridlines: true,
      axes: true,
    });
    // The second series' own label is gone with the plot's labels.
    expect(chartXml(e)).not.toContain('<c:dLbl>');
  });

  test('sets axis titles, and refuses axis edits on a pie chart', async () => {
    const e = await inserted('lineMarkers');
    expect(
      e.exec({
        type: 'setChartFormat',
        format: { categoryAxisTitle: 'Quarter', valueAxisTitle: 'Sales' },
      }).ok
    ).toBe(true);
    expect(format(e).axisTitles).toEqual({ category: 'Quarter', value: 'Sales' });
    expect(chartXml(e)).toContain('<a:bodyPr rot="-5400000" vert="horz"/>');
    expect(e.exec({ type: 'setChartFormat', format: { valueAxisTitle: null } }).ok).toBe(true);
    expect(format(e).axisTitles).toEqual({ category: 'Quarter', value: null });
    const pie = await inserted('pie');
    expect(pie.exec({ type: 'setChartFormat', format: { gridlines: true } })).toMatchObject({
      ok: false,
      reason: 'errors.chartFormat.noAxes',
    });
    expect(pie.exec({ type: 'setChartFormat', format: { bogus: 1 } as never }).ok).toBe(false);
  });
});

describe('setChartColor', () => {
  test('colors a series, and gives back its automatic color', async () => {
    const e = await inserted('lineMarkers');
    expect(format(e).colors[1]).toEqual({
      name: 'Series 2',
      color: null,
      automatic: { theme: 'accent2' },
    });
    expect(e.exec({ type: 'setChartColor', index: 1, color: '#C00000' }).ok).toBe(true);
    expect(format(e).colors[1]!.color).toBe('#C00000');
    expect(
      e.exec({ type: 'setChartColor', index: 1, color: { theme: 'accent6', shift: -25 } }).ok
    ).toBe(true);
    expect(format(e).colors[1]!.color).toEqual({ theme: 'accent6', shift: -25 });
    expect(e.exec({ type: 'setChartColor', index: 1, color: null }).ok).toBe(true);
    expect(format(e).colors[1]!.color).toBeNull();
    expect(e.exec({ type: 'setChartColor', index: 9, color: '#000000' }).ok).toBe(false);
  });

  test('colors the points of a pie chart, in index order', async () => {
    const e = await inserted('pie');
    expect(format(e).colorsBy).toBe('point');
    expect(e.exec({ type: 'setChartColor', index: 2, color: '#00B050' }).ok).toBe(true);
    expect(e.exec({ type: 'setChartColor', index: 0, color: '#FFC000' }).ok).toBe(true);
    expect(format(e).colors.map((c) => c.color)).toEqual(['#FFC000', null, '#00B050', null]);
    expect(chartXml(e).indexOf('<c:idx val="0"/><c:bubble3D')).toBeLessThan(
      chartXml(e).indexOf('<c:idx val="2"/><c:bubble3D')
    );
    expect(e.exec({ type: 'setChartColor', index: 2, color: null }).ok).toBe(true);
    expect(chartXml(e)).not.toContain('<c:idx val="2"/><c:bubble3D');
    // Pie point colors stay through a change to a doughnut.
    expect(e.exec({ type: 'setChartType', chartType: 'doughnut' }).ok).toBe(true);
    expect(format(e).colors[0]!.color).toBe('#FFC000');
  });

  test('a column chart turned into a pie colors one entry per point, each its own color', async () => {
    const e = await inserted('columnClustered');
    expect(e.exec({ type: 'setChartType', chartType: 'pie' }).ok).toBe(true);
    const f = format(e);
    expect(f.colorsBy).toBe('point');
    const data = e.chartData();
    if (!data.ok) throw new Error(data.reason);
    expect(f.colors.map((c) => c.name)).toEqual(data.data.categories.map((c) => c.label));
    const automatic = f.colors.map((c) => JSON.stringify(c.automatic));
    expect(new Set(automatic).size).toBe(automatic.length);
  });

  test('a scatter chart of markers only colors its markers and keeps its lines hidden', async () => {
    const e = await inserted('scatter');
    expect(e.exec({ type: 'setChartColor', index: 0, color: '#7030A0' }).ok).toBe(true);
    expect(format(e)).toMatchObject({ type: 'scatter' });
    expect(format(e).colors[0]!.color).toBe('#7030A0');
  });
});

describe('the Chart tab after an insert or a type change', () => {
  test('charts inserted one after another cascade, each with its own corner', async () => {
    const e = await inserted('columnClustered');
    for (const type of ['line', 'pie'] as const) {
      e.setSelection({ kind: 'none' });
      expect(e.exec({ type: 'insertChart', chartType: type }).ok).toBe(true);
    }
    const zip = readZip(e.save()!);
    if (!zip.ok) throw new Error('package');
    const xml = strFromU8(zip.entries.get('/ppt/slides/slide1.xml')!);
    const box = /<p:xfrm><a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/g;
    const boxes = [...xml.matchAll(box)].map((m) => m.slice(1, 5).map(Number)).slice(-3);
    expect(boxes).toHaveLength(3);
    // Each later chart starts one step (5 % of the slide) past the one before it.
    const { slideWidth, slideHeight } = e.snapshot();
    for (const [i, b] of boxes.entries()) {
      if (i === 0) continue;
      const a = boxes[i - 1]!;
      expect(b[0]! - a[0]!).toBe(Math.round(slideWidth * 0.05));
      expect(b[1]! - a[1]!).toBe(Math.round(slideHeight * 0.05));
    }
  });

  test('Insert > Chart opens the Chart tab and keeps the focus on the slide', async () => {
    const e = await open(fixture('charts'));
    e.setSelection({ kind: 'none' });
    const asked: [string, boolean][] = [];
    const stop = listenPaneRequests(e, (pane, request) => asked.push([pane, request.focus]));
    expect(runToolbarCommand(e, 'insert.chart', 'pie').ok).toBe(true);
    stop();
    expect(asked).toEqual([['chart', false]]);
  });

  test('a type change says what it drops, and a pie chart says it shows one series', async () => {
    const e = await inserted('columnClustered');
    expect(e.exec({ type: 'setChartColor', index: 0, color: '#00B050' }).ok).toBe(true);
    const column = format(e);
    expect(e.exec({ type: 'setChartType', chartType: 'line' }).ok).toBe(true);
    expect(chartTypeChangeNote(column, format(e))).toBeNull();
    const line = format(e);
    expect(e.exec({ type: 'setChartType', chartType: 'pie' }).ok).toBe(true);
    expect(chartTypeChangeNote(line, format(e))).toBe('chartFormat.typeChangeRemoved');
    const data = e.chartData();
    if (!data.ok) throw new Error(data.reason);
    expect(chartSeriesNote(format(e), data.data.series.length)).toBe('chartFormat.pieFirstSeries');
    expect(chartSeriesNote(format(e), 1)).toBeNull();
    expect(chartSeriesNote(line, data.data.series.length)).toBeNull();
  });
});
