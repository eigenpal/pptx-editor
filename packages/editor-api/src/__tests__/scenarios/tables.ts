/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Scenarios for tables: add, cells, merge and split, borders, margins, style options, rows and
// columns, text runs, and clear.

import type { RequestContext, Table } from '../../runtime/public.ts';
import { check, pt, same, type Json, type Scenario } from './types.ts';

/** The table on the first slide of `tables`. */
const fixtureTable = (context: RequestContext, slide = 0): Table =>
  context.presentation.slides.getItemAt(slide).shapes.getItem('2').getTable();

async function gridOf(table: Table): Promise<Json> {
  table.load('rowCount,columnCount,values');
  await table.context.sync();
  return { rows: table.rowCount, columns: table.columnCount, values: table.values };
}

/** The table a scenario added: the last shape of the slide. */
async function addedTable(context: RequestContext, slide: number): Promise<Table> {
  const shapes = context.presentation.slides.getItemAt(slide).shapes.load('items/type');
  await context.sync();
  const shape = shapes.items[shapes.items.length - 1]!;
  check(shape.type === 'Table', 'the last shape is a table');
  return shape.getTable();
}

async function cellFormat(table: Table, row: number, column: number): Promise<Json> {
  const cell = table.getCellOrNullObject(row, column);
  cell.load('text,horizontalAlignment,verticalAlignment');
  const font = cell.font.load('bold,color,size');
  const fill = cell.fill.load('type,foregroundColor');
  await table.context.sync();
  return {
    text: cell.text,
    align: [cell.horizontalAlignment, cell.verticalAlignment],
    font: { bold: font.bold, color: font.color, size: font.size },
    fill: fill.type === 'Solid' ? fill.foregroundColor : fill.type,
  };
}

async function edgesOf(table: Table): Promise<Json> {
  const cell = table.getCellOrNullObject(1, 1);
  const bottom = cell.borders.bottom.load('color,weight,dashStyle,transparency');
  const top = table.getCellOrNullObject(2, 1).borders.top.load('color,weight');
  const margins = cell.margins.load('left,right,top,bottom');
  await table.context.sync();
  return {
    bottom: [bottom.color ?? null, bottom.weight ?? null, bottom.dashStyle ?? null],
    belowTop: [top.color ?? null, top.weight ?? null],
    margins: [margins.left, margins.right, margins.top, margins.bottom].map((m) =>
      m === undefined ? null : pt(m)
    ),
  };
}

async function settingsOf(table: Table): Promise<Json> {
  const s = table.styleSettings.load(
    'areRowsBanded,areColumnsBanded,isFirstRowHighlighted,isLastRowHighlighted,isFirstColumnHighlighted,isLastColumnHighlighted'
  );
  await table.context.sync();
  return {
    rowsBanded: s.areRowsBanded,
    columnsBanded: s.areColumnsBanded,
    firstRow: s.isFirstRowHighlighted,
    lastRow: s.isLastRowHighlighted,
    firstColumn: s.isFirstColumnHighlighted,
    lastColumn: s.isLastColumnHighlighted,
  };
}

async function mergesOf(table: Table): Promise<Json> {
  const areas = table.getMergedAreas().load('items/rowIndex,items/columnIndex,items/rowCount');
  await table.context.sync();
  const spans = areas.items.map((a) => a.load('columnCount'));
  await table.context.sync();
  return areas.items.map(
    (a, i) => `${a.rowIndex},${a.columnIndex} ${a.rowCount}x${spans[i]!.columnCount}`
  );
}

export const TABLE_SCENARIOS: readonly Scenario[] = [
  {
    name: 'tables.add-with-options',
    area: 'tables',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const shapes = context.presentation.slides.getItemAt(4).shapes;
      const shape = shapes.addTable(3, 3, {
        left: 60,
        top: 150,
        width: 480,
        values: [
          ['Region', 'Q2', 'Q3'],
          ['North', '120', '135'],
          ['South', '98', '110'],
        ],
        uniformCellProperties: { font: { size: 14 } },
        specificCellProperties: [
          [{ font: { bold: true } }, { font: { bold: true } }, { font: { bold: true } }],
          [{}, {}, {}],
          [{}, {}, { fill: { color: '#E2EFDA' } }],
        ],
        columns: [{ columnWidth: 200 }, {}, {}],
      });
      shape.name = 'Results';
      await context.sync();
      const table = shape.getTable();
      const column = table.columns.getItemAt(0).load('width');
      await context.sync();
      return {
        grid: await gridOf(table),
        firstColumn: pt(column.width),
        header: await cellFormat(table, 0, 1),
        highlighted: await cellFormat(table, 2, 2),
      };
    },
    async readBack(context) {
      const table = await addedTable(context, 4);
      return {
        grid: await gridOf(table),
        header: await cellFormat(table, 0, 1),
        highlighted: await cellFormat(table, 2, 2),
      };
    },
  },
  {
    name: 'tables.cell-format',
    area: 'tables',
    deck: 'tables',
    writes: true,
    async script(context) {
      const table = fixtureTable(context);
      const before = await cellFormat(table, 1, 1);
      const cell = table.getCellOrNullObject(1, 1);
      cell.text = 'Updated';
      cell.font.bold = true;
      cell.font.color = '#C00000';
      cell.font.size = 16;
      cell.fill.setSolidColor('#FFF2CC');
      cell.horizontalAlignment = 'Center';
      cell.verticalAlignment = 'Bottom';
      await context.sync();
      return { before, after: await cellFormat(table, 1, 1) };
    },
    readBack: (context) => cellFormat(fixtureTable(context), 1, 1),
  },
  {
    name: 'tables.merge-and-split',
    area: 'tables',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const shape = context.presentation.slides
        .getItemAt(4)
        .shapes.addTable(4, 4, { left: 40, top: 140, width: 400 });
      await context.sync();
      const table = shape.getTable();
      table.mergeCells(0, 0, 1, 4);
      table.mergeCells(1, 0, 2, 2);
      await context.sync();
      const merged = await mergesOf(table);
      table.getCellOrNullObject(1, 0).split(2, 1);
      await context.sync();
      const afterSplit = await mergesOf(table);
      table.getCellOrNullObject(0, 0).resize(1, 3);
      await context.sync();
      return { merged, afterSplit, resized: await mergesOf(table) };
    },
    async readBack(context) {
      return mergesOf(await addedTable(context, 4));
    },
  },
  {
    name: 'tables.borders-and-margins',
    area: 'tables',
    deck: 'tables',
    writes: true,
    async script(context) {
      const table = fixtureTable(context);
      const cell = table.getCellOrNullObject(1, 1);
      const bottom = cell.borders.bottom;
      bottom.color = '#1F4E79';
      bottom.weight = 3;
      bottom.dashStyle = 'Dash';
      cell.margins.left = 10;
      cell.margins.top = 4;
      await context.sync();
      return edgesOf(table);
    },
    readBack: (context) => edgesOf(fixtureTable(context)),
  },
  {
    name: 'tables.style-options',
    area: 'tables',
    deck: 'tables',
    writes: true,
    async script(context) {
      const table = fixtureTable(context);
      const before = await settingsOf(table);
      const settings = table.styleSettings;
      settings.areRowsBanded = true;
      settings.areColumnsBanded = true;
      settings.isFirstRowHighlighted = false;
      settings.isLastRowHighlighted = true;
      settings.isFirstColumnHighlighted = true;
      settings.isLastColumnHighlighted = true;
      await context.sync();
      return { before, after: await settingsOf(table) };
    },
    readBack: (context) => settingsOf(fixtureTable(context)),
  },
  {
    name: 'tables.rows-and-columns',
    area: 'tables',
    deck: 'tables',
    writes: true,
    async script(context) {
      const table = fixtureTable(context);
      const before = await gridOf(table);
      table.rows.add(null, 2);
      table.columns.add(0, 1);
      await context.sync();
      const grown = await gridOf(table);
      const rows = table.rows.load('items');
      await context.sync();
      table.rows.deleteRows([rows.items[rows.items.length - 1]!]);
      table.rows.getItemAt(0).height = 40;
      table.columns.getItemAt(0).width = 90;
      await context.sync();
      const row = table.rows.getItemAt(0).load('height');
      const column = table.columns.getItemAt(0).load('width');
      await context.sync();
      return {
        before,
        grown,
        after: await gridOf(table),
        size: [pt(row.height), pt(column.width)],
      };
    },
    readBack: (context) => gridOf(fixtureTable(context)),
  },
  {
    name: 'tables.text-runs',
    area: 'tables',
    deck: 'tables',
    writes: true,
    async script(context) {
      const cell = fixtureTable(context).getCellOrNullObject(0, 0);
      cell.textRuns = [
        { text: 'Total ', font: { bold: true } },
        { text: 'revenue', font: { italic: true, color: '#2F5597' } },
      ];
      await context.sync();
      return this.readBack!(context);
    },
    async readBack(context) {
      const cell = fixtureTable(context).getCellOrNullObject(0, 0).load('text,textRuns');
      await context.sync();
      return {
        text: cell.text,
        runs: cell.textRuns.map((r) => ({
          text: r.text,
          bold: r.font?.bold ?? null,
          italic: r.font?.italic ?? null,
          color: r.font?.color ?? null,
        })),
      };
    },
  },
  {
    name: 'tables.clear',
    area: 'tables',
    deck: 'tables',
    writes: true,
    async script(context) {
      const first = fixtureTable(context, 0);
      first.getCellOrNullObject(0, 0).font.bold = true;
      await context.sync();
      first.clear({ text: true });
      fixtureTable(context, 1).clear({ format: true });
      await context.sync();
      const bold = first.getCellOrNullObject(0, 0).font.load('bold');
      await context.sync();
      same(bold.bold, true, 'clearing text keeps the format');
      return this.readBack!(context);
    },
    async readBack(context) {
      const first = await gridOf(fixtureTable(context, 0));
      const second = await gridOf(fixtureTable(context, 1));
      return { first, second };
    },
  },
];
