// The chart data edits (`setChartData`): the save evidence of the chart row's own edit. Each
// recipe edits the data of a chart in the `chart-data` fixture through the public editor API,
// saves, and opens the saved copy again. Its check reads the chart part's caches, point counts,
// and cell references, and the cells and table of the embedded workbook, so the chart and its
// workbook must agree after the reopen. `FEATURE_EDITS['objects.charts']` makes every edit at
// once; the recipes here make one edit each.
//
// `chart-data-edits.test.ts` runs every recipe. `bun scripts/fidelity/chart-data-edits.ts --out
// <dir>` writes each edited copy as `<id>.pptx`, to check for a repair prompt
// (CONTRIBUTING.md, "Fidelity harness").

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { parseArgs } from 'node:util';
import { strFromU8 } from 'fflate';
import type { ChartData } from '../../packages/core/src/contracts/index.ts';
import { binaryPartOf, readZip, type PptxPackage } from '../../packages/core/src/store/index.ts';
import {
  all,
  has,
  lacks,
  runFeatureEdit,
  type EditTools,
  type FeatureEdit,
  type PartRule,
  type SavedDeck,
} from './feature-edit-kit.ts';

/** The parts a chart data edit may change: the chart parts and their workbooks. */
export const CHART_DATA_PARTS: readonly PartRule[] = [
  /^\/ppt\/charts\/chart\d+\.xml$/,
  /^\/ppt\/embeddings\/Workbook\d+\.xlsx$/,
];

/** The text of a part that the reader keeps as bytes (a chart part). */
export function bytesText(pkg: PptxPackage, part: string): string {
  const bytes = binaryPartOf(pkg, part);
  if (!bytes) throw new Error(`no part ${part}`);
  return strFromU8(bytes);
}

/** An entry of the embedded workbook `book`, as text. */
export function workbookText(pkg: PptxPackage, book: string, entry: string): string {
  const zip = readZip(binaryPartOf(pkg, `/ppt/embeddings/${book}`)!);
  if (!zip.ok) throw new Error(`${book} does not read: ${zip.reason}`);
  const bytes = zip.entries.get(entry);
  if (!bytes) throw new Error(`${book} has no ${entry}`);
  return strFromU8(bytes);
}

const chart = (d: SavedDeck, n: number): string => bytesText(d.pkg, `/ppt/charts/chart${n}.xml`);
const sheet = (d: SavedDeck, n: number): string =>
  workbookText(d.pkg, `Workbook${n}.xlsx`, '/xl/worksheets/sheet1.xml');
const table = (d: SavedDeck): string =>
  workbookText(d.pkg, 'Workbook1.xlsx', '/xl/tables/table1.xml');

/** Select the chart of the current slide and replace its data with `edit` of its data. */
export function editChart(t: EditTools, edit: (data: ChartData) => Omit<ChartData, 'workbook'>) {
  t.select('Chart 1');
  const read = t.editor.chartData();
  if (!read.ok) throw new Error(`the chart data does not read: ${read.reason}`);
  t.run({ type: 'setChartData', data: edit(read.data) });
}

/** A string cache of one cell, as the chart part writes it. */
const nameCache = (ref: string, name: string): string =>
  `<c:f>${ref}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${name}</c:v></c:pt></c:strCache>`;
const inline = (ref: string, text: string): string =>
  `<c r="${ref}" t="inlineStr"><is><t>${text}</t></is></c>`;
const num = (ref: string, v: number): string => `<c r="${ref}"><v>${v}</v></c>`;

const recipe = (r: Omit<FeatureEdit, 'fixture' | 'parts'>): FeatureEdit => ({
  fixture: 'chart-data',
  parts: CHART_DATA_PARTS,
  ...r,
});

/** The recipes, by `objects.charts: <edit>`. */
export const CHART_DATA_EDITS: Readonly<Record<string, FeatureEdit>> = {
  'objects.charts: edit values': recipe({
    edit: (t) =>
      editChart(t, (d) => ({
        ...d,
        series: d.series.map((s, j) => (j === 1 ? { ...s, values: [12, 24.5, null, 48] } : s)),
      })),
    check: (d) =>
      all(
        has(
          chart(d, 1),
          '<c:f>Sheet1!$C$2:$C$5</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="4"/><c:pt idx="0"><c:v>12</c:v></c:pt><c:pt idx="1"><c:v>24.5</c:v></c:pt><c:pt idx="3"><c:v>48</c:v></c:pt></c:numCache>'
        ),
        has(sheet(d, 1), num('C2', 12), num('C3', 24.5), num('C5', 48)),
        lacks(sheet(d, 1), 'r="C4"')
      ),
  }),
  'objects.charts: rename a series': recipe({
    edit: (t) =>
      editChart(t, (d) => ({
        ...d,
        series: d.series.map((s, j) => (j === 0 ? { ...s, name: 'Budget & plan' } : s)),
      })),
    check: (d) =>
      all(
        has(chart(d, 1), nameCache('Sheet1!$B$1', 'Budget &amp; plan')),
        has(sheet(d, 1), inline('B1', 'Budget &amp; plan')),
        has(table(d), '<tableColumn id="2" name="Budget &amp; plan"/>')
      ),
  }),
  'objects.charts: add a series': recipe({
    edit: (t) =>
      editChart(t, (d) => ({
        ...d,
        series: [...d.series, { name: 'Stretch', values: [50, 45, 60, 40] }],
      })),
    check: (d) =>
      all(
        has(
          chart(d, 1),
          `<c:ser><c:idx val="3"/><c:order val="3"/><c:tx><c:strRef>${nameCache('Sheet1!$E$1', 'Stretch')}`,
          '<c:f>Sheet1!$E$2:$E$5</c:f>'
        ),
        has(
          sheet(d, 1),
          inline('E1', 'Stretch'),
          num('E2', 50),
          num('E5', 40),
          '<dimension ref="A1:E5"/>'
        ),
        has(
          table(d),
          'ref="A1:E5"',
          '<tableColumns count="5">',
          '<tableColumn id="5" name="Stretch"/>'
        )
      ),
  }),
  'objects.charts: remove a series': recipe({
    edit: (t) => editChart(t, (d) => ({ ...d, series: [d.series[0]!, d.series[2]!] })),
    check: (d) =>
      all(
        has(chart(d, 1), nameCache('Sheet1!$C$1', 'Forecast'), '<c:f>Sheet1!$C$2:$C$5</c:f>'),
        lacks(chart(d, 1), 'Actual', 'C00000'),
        has(sheet(d, 1), inline('C1', 'Forecast'), num('C2', 45), '<dimension ref="A1:C5"/>'),
        lacks(sheet(d, 1), 'r="D1"'),
        has(table(d), 'ref="A1:C5"', '<tableColumns count="3">')
      ),
  }),
  'objects.charts: add a category': recipe({
    edit: (t) =>
      editChart(t, (d) => ({
        ...d,
        categories: [...d.categories, { label: 'Central' }],
        series: d.series.map((s, j) => ({ ...s, values: [...s.values, 30 + j] })),
      })),
    check: (d) =>
      all(
        has(
          chart(d, 1),
          '<c:f>Sheet1!$A$2:$A$6</c:f><c:strCache><c:ptCount val="5"/>',
          '<c:pt idx="4"><c:v>Central</c:v></c:pt>',
          '<c:f>Sheet1!$D$2:$D$6</c:f>'
        ),
        has(sheet(d, 1), inline('A6', 'Central'), num('B6', 30), num('D6', 32)),
        has(table(d), 'ref="A1:D6"')
      ),
  }),
  'objects.charts: remove a category': recipe({
    edit: (t) =>
      editChart(t, (d) => ({
        ...d,
        categories: d.categories.filter((_, i) => i !== 1),
        series: d.series.map((s) => ({ ...s, values: s.values.filter((_, i) => i !== 1) })),
      })),
    check: (d) =>
      all(
        has(
          chart(d, 1),
          '<c:f>Sheet1!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/>',
          // The West point keeps its fill and label: index 3 is now 2.
          '<c:dPt><c:idx val="2"/>',
          '<c:dLbl><c:idx val="2"/>'
        ),
        lacks(chart(d, 1), 'South', '<c:idx val="3"/><c:invertIfNegative'),
        has(sheet(d, 1), inline('A3', 'East'), '<dimension ref="A1:D4"/>'),
        lacks(sheet(d, 1), 'r="A5"'),
        has(table(d), 'ref="A1:D4"')
      ),
  }),
  'objects.charts: edit a scatter chart': recipe({
    slide: 1,
    edit: (t) =>
      editChart(t, (d) => ({
        ...d,
        categories: [...d.categories, { label: '6.5' }],
        series: [
          ...d.series.map((s) => ({ ...s, values: [...s.values, 8] })),
          { name: 'Sample C', values: [1, 2, 3, 4, 5, 6] },
        ],
      })),
    check: (d) =>
      all(
        has(
          chart(d, 2),
          '<c:xVal><c:numRef><c:f>Sheet1!$A$2:$A$7</c:f>',
          '<c:pt idx="5"><c:v>6.5</c:v></c:pt>',
          '<c:yVal><c:numRef><c:f>Sheet1!$D$2:$D$7</c:f>'
        ),
        has(sheet(d, 2), num('A7', 6.5), num('B7', 8), inline('D1', 'Sample C'), num('D7', 6))
      ),
  }),
};

if (import.meta.main) {
  const { values } = parseArgs({ options: { out: { type: 'string' } } });
  if (values.out) mkdirSync(values.out, { recursive: true });
  let failed = 0;
  for (const [id, r] of Object.entries(CHART_DATA_EDITS)) {
    const result = await runFeatureEdit(id, r);
    if (!result.ok) failed++;
    console.log(`${id}: ${result.ok ? `ok ${result.fingerprint}` : `FAIL ${result.failure}`}`);
    const file = id.replace(/[^a-z0-9.-]+/gi, '-');
    if (values.out && result.bytes) writeFileSync(join(values.out, `${file}.pptx`), result.bytes);
  }
  process.exitCode = failed ? 1 : 0;
}
