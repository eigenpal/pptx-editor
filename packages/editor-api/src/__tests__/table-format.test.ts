/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Table format and shape adjustments end to end: cell borders and margins, table style
// settings, merged areas, the options of addTable, and preset adjust values.

import { describe, expect, test } from 'bun:test';
import { buildPptx, cell, shape, table } from '../../../../scripts/fixtures/builder.ts';
import { isPptxEditorError, type RequestContext, type Table } from '../index.ts';
import { IN, open, refusedAtSync, rejection, reopen, thrown } from './support.ts';

const code = (error: unknown): string | undefined =>
  isPptxEditorError(error) ? error.code : String(error);
const inner = (error: unknown): string | undefined =>
  isPptxEditorError(error) ? error.debugInfo.innerError?.code : undefined;

/** A 3 by 3 table on the second slide. */
async function addTable(context: RequestContext): Promise<Table> {
  const table = context.presentation.slides
    .getItemAt(1)
    .shapes.addTable(3, 3, { left: 20, top: 20 })
    .getTable();
  await context.sync();
  return table;
}

function savedTable(context: RequestContext): Table {
  return context.presentation.slides.getItemAt(1).shapes.getItemAt(0).getTable();
}

const LIGHT_STYLE =
  '<a:tblStyle styleId="{9D7B26C5-4107-4FEC-AEDC-1716B250EE53}" styleName="Light Style 1">' +
  '<a:wholeTbl><a:tcStyle><a:tcBdr/></a:tcStyle></a:wholeTbl></a:tblStyle>';

describe('table cell borders and margins', () => {
  test('borders and margins write, agree with the neighbor, and survive a save', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const cell = (await addTable(context)).getCellOrNullObject(1, 1);
      cell.borders.left.color = '#FF0000';
      cell.borders.left.weight = 3;
      cell.borders.diagonalUp.dashStyle = 'Dash';
      cell.borders.diagonalUp.transparency = 0.25;
      cell.margins.left = 12;
      cell.margins.top = 0;
      await context.sync();
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const table = savedTable(context);
      const cell = table.getCellOrNullObject(1, 1);
      const left = cell.borders.left.load();
      const diagonal = cell.borders.diagonalUp.load('dashStyle,transparency,color,weight');
      const across = table.getCellOrNullObject(1, 0).borders.right.load('color,weight');
      const margins = cell.margins.load();
      await context.sync();
      expect([left.color, left.weight, left.dashStyle, left.transparency]).toEqual([
        '#FF0000',
        3,
        'Solid',
        0,
      ]);
      expect([diagonal.dashStyle, diagonal.transparency, diagonal.color, diagonal.weight]).toEqual([
        'Dash',
        0.25,
        '#000000',
        1,
      ]);
      expect([across.color, across.weight]).toEqual(['#FF0000', 3]);
      expect([margins.left, margins.right, margins.top, margins.bottom]).toEqual([12, 7.2, 0, 3.6]);
      cell.margins.left = undefined;
      await context.sync();
      const reset = cell.margins.load('left');
      await context.sync();
      expect(reset.left).toBe(7.2);
    });
  });

  test('a cell with no line reads undefined, and load names the borders', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const cell = (await addTable(context)).getCellOrNullObject(0, 0);
      cell.load('borders,margins');
      await context.sync();
      expect(cell.borders.diagonalDown.color).toBeUndefined();
      expect(cell.borders.diagonalDown.weight).toBeUndefined();
      expect(cell.margins.bottom).toBe(3.6);
    });
  });

  test('bad values refuse at the sync, and nothing applies', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const cell = (await addTable(context)).getCellOrNullObject(0, 0);
      expect(code(await refusedAtSync(context, () => (cell.borders.top.weight = -1)))).toBe(
        'InvalidArgument'
      );
      expect(code(await refusedAtSync(context, () => (cell.borders.top.transparency = 2)))).toBe(
        'InvalidArgument'
      );
      expect(
        code(await refusedAtSync(context, () => (cell.borders.top.dashStyle = 'Wavy' as never)))
      ).toBe('InvalidArgument');
      expect(code(await refusedAtSync(context, () => (cell.borders.top.color = undefined)))).toBe(
        'InvalidArgument'
      );
      expect(code(await refusedAtSync(context, () => (cell.margins.top = Number.NaN)))).toBe(
        'InvalidArgument'
      );
      expect(code(await refusedAtSync(context, () => (cell.margins.top = -2)))).toBe(
        'InvalidArgument'
      );
      await context.sync();
    });
    const error = await rejection(
      runtime.run(async (context) => {
        const cell = (await addTable(context)).getCellOrNullObject(0, 0);
        cell.borders.top.color = 'not a color';
        await context.sync();
      })
    );
    expect(code(error)).toBe('InvalidArgument');
  });
});

describe('table style settings', () => {
  test('the style options and a style the deck defines write and read back', async () => {
    const runtime = await open(
      buildPptx({
        slides: [
          { layout: 'blank', shapes: '' },
          { layout: 'blank', shapes: '' },
        ],
        tableStyles: LIGHT_STYLE,
      })
    );
    await runtime.run(async (context) => {
      const settings = (await addTable(context)).styleSettings.load();
      await context.sync();
      expect(settings.style).toBe('MediumStyle2Accent1');
      expect([settings.isFirstRowHighlighted, settings.areRowsBanded]).toEqual([true, true]);
      settings.isFirstRowHighlighted = false;
      settings.isLastRowHighlighted = true;
      settings.isFirstColumnHighlighted = true;
      settings.isLastColumnHighlighted = true;
      settings.areColumnsBanded = true;
      settings.areRowsBanded = false;
      settings.style = 'LightStyle1';
      await context.sync();
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const table = savedTable(context);
      table.load('styleSettings');
      await context.sync();
      const s = table.styleSettings;
      expect([
        s.isFirstRowHighlighted,
        s.isLastRowHighlighted,
        s.isFirstColumnHighlighted,
        s.isLastColumnHighlighted,
        s.areColumnsBanded,
        s.areRowsBanded,
        s.style,
      ]).toEqual([false, true, true, true, true, false, 'LightStyle1']);
      s.style = 'NoStyleNoGrid';
      await context.sync();
      s.load('style');
      await context.sync();
      expect(s.style).toBe('NoStyleNoGrid');
    });
  });

  test('a style the deck does not define, or with no TableStyle value, refuses', async () => {
    const custom =
      '<a:tblStyle styleId="{11111111-2222-3333-4444-555555555555}" styleName="House style">' +
      '<a:wholeTbl><a:tcStyle><a:tcBdr/></a:tcStyle></a:wholeTbl></a:tblStyle>';
    const runtime = await open(
      buildPptx({
        slides: [
          { layout: 'blank', shapes: '' },
          { layout: 'blank', shapes: '' },
        ],
        tableStyles: custom,
      })
    );
    const error = await rejection(
      runtime.run(async (context) => {
        (await addTable(context)).styleSettings.style = 'DarkStyle2';
        await context.sync();
      })
    );
    expect([code(error), inner(error)]).toEqual(['NotSupported', 'edit.table-style-undefined']);
    await runtime.run(async (context) => {
      const table = await addTable(context);
      expect(
        code(await refusedAtSync(context, () => (table.styleSettings.style = 'Fancy' as never)))
      ).toBe('InvalidArgument');
      expect(
        code(await refusedAtSync(context, () => (table.styleSettings.areRowsBanded = 1 as never)))
      ).toBe('InvalidArgument');
    });
    // A table that uses the deck's own style, which has no TableStyle value.
    const house = await open(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: table({
              id: 2,
              x: 0,
              y: 0,
              cx: IN,
              cy: IN,
              columns: [IN],
              tblPr:
                '<a:tblPr bandRow="1"><a:tableStyleId>{11111111-2222-3333-4444-555555555555}' +
                '</a:tableStyleId></a:tblPr>',
              rows: `<a:tr h="370840">${cell('x')}</a:tr>`,
            }),
          },
        ],
        tableStyles: custom,
      })
    );
    await house.run(async (context) => {
      const settings = context.presentation.slides
        .getItemAt(0)
        .shapes.getItemAt(0)
        .getTable()
        .styleSettings.load();
      await context.sync();
      expect(settings.areRowsBanded).toBe(true);
      const read = thrown(() => settings.style);
      expect([code(read), inner(read)]).toEqual(['NotImplemented', 'read.table-style-unknown']);
    });
  });
});

describe('merged areas', () => {
  test('getMergedAreas lists each merged area by its first cell', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const table = await addTable(context);
      table.mergeCells(0, 0, 2, 2);
      table.mergeCells(2, 1, 1, 2);
      await context.sync();
      const areas = table.getMergedAreas();
      areas.load('items');
      const count = areas.getCount();
      const covered = areas.getItemAtOrNullObject(1, 1).load('rowIndex,columnIndex,rowCount');
      const none = areas.getItemAtOrNullObject(2, 0);
      await context.sync();
      expect(count.value).toBe(2);
      expect(
        areas.items.map((c) => [c.rowIndex, c.columnIndex, c.rowCount, c.columnCount])
      ).toEqual([
        [0, 0, 2, 2],
        [2, 1, 1, 2],
      ]);
      expect([covered.rowIndex, covered.columnIndex, covered.rowCount]).toEqual([0, 0, 2]);
      expect(none.isNullObject).toBe(true);
      areas.items[1]!.text = 'merged';
      await context.sync();
      table.load('values');
      await context.sync();
      expect(table.values[2]![1]).toBe('merged');
    });
  });
});

describe('addTable options', () => {
  test('sizes, merges, uniform and specific cell properties apply in one batch', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(1).shapes.addTable(2, 3, {
        values: [
          ['a', 'b', 'c'],
          ['d', 'e', 'f'],
        ],
        columns: [{ columnWidth: 50 }, {}, { columnWidth: 70 }],
        rows: [{ rowHeight: 30 }, { rowHeight: 40 }],
        mergedAreas: [{ rowIndex: 1, columnIndex: 1, rowCount: 1, columnCount: 2 }],
        uniformCellProperties: {
          margins: { left: 4 },
          borders: { bottom: { color: '#00B050', weight: 2 } },
          verticalAlignment: 'Middle',
        },
        specificCellProperties: [
          [{ fill: { color: '#FFF2CC' }, font: { italic: true } }, {}, {}],
          [{ text: 'D', horizontalAlignment: 'Right' }, {}, {}],
        ],
      });
      await context.sync();
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const table = savedTable(context);
      table.load('values');
      const widths = table.columns.load('items');
      const heights = table.rows.load('items');
      const first = table.getCellOrNullObject(0, 0).load('verticalAlignment');
      const fill = table.getCellOrNullObject(0, 0).fill.load('foregroundColor');
      const italic = table.getCellOrNullObject(0, 0).font.load('italic');
      const d = table.getCellOrNullObject(1, 0).load('text,horizontalAlignment');
      const margin = table.getCellOrNullObject(0, 2).margins.load('left');
      const bottom = table.getCellOrNullObject(0, 2).borders.bottom.load('color,weight');
      const merged = table.getCellOrNullObject(1, 1).load('columnCount');
      await context.sync();
      expect(widths.items.map((c) => c.width)).toEqual([50, expect.any(Number), 70]);
      expect(heights.items[0]!.height).toBeGreaterThanOrEqual(30);
      expect(heights.items[1]!.height).toBeGreaterThanOrEqual(40);
      expect(first.verticalAlignment).toBe('Middle');
      expect(fill.foregroundColor).toBe('#FFF2CC');
      expect(italic.italic).toBe(true);
      expect([d.text, d.horizontalAlignment]).toEqual(['D', 'Right']);
      expect(margin.left).toBe(4);
      expect([bottom.color, bottom.weight]).toEqual(['#00B050', 2]);
      expect(merged.columnCount).toBe(2);
    });
  });

  test('a bad option refuses at the sync and adds no table', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(1).shapes;
      const bad = [
        { columns: [{ columnWidth: 50 }] },
        { rows: [{ rowHeight: -1 }, {}] },
        { mergedAreas: [{ rowIndex: 0, columnIndex: 1, rowCount: 1, columnCount: 2 }] },
        { uniformCellProperties: { text: 'x', textRuns: [{ text: 'y' }] } },
        { uniformCellProperties: { borders: { inside: {} } } },
        { specificCellProperties: [[{}, {}]] },
        { style: 'Fancy' },
        { uniformCellProperties: { font: { size: 0 } } },
        { somethingElse: true },
      ];
      for (const options of bad)
        expect(
          code(await refusedAtSync(context, () => shapes.addTable(2, 2, options as never)))
        ).toBe('InvalidArgument');
      shapes.load('items');
      await context.sync();
      expect(shapes.items.length).toBe(0);
    });
  });
});

describe('shape adjustments', () => {
  const rounded = (): Uint8Array =>
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes:
            shape({ id: 2, name: 'Round', x: IN, y: IN, cx: 2 * IN, cy: IN, prst: 'roundRect' }) +
            shape({ id: 3, name: 'Arc', x: IN, y: IN, cx: IN, cy: IN, prst: 'arc' }),
        },
      ],
    });

  test('count, get, and set follow the preset, and survive a save', async () => {
    const runtime = await open(rounded());
    await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      const round = shapes.getItemAt(0).adjustments.load('count');
      const first = round.get(0);
      const arc = shapes.getItemAt(1).adjustments;
      const angle = arc.get(0);
      await context.sync();
      expect(round.count).toBe(1);
      expect(first.value).toBe(0.16667);
      expect(angle.value).toBe(270);
      round.set(0, 0.25);
      arc.set(1, 45);
      await context.sync();
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      const round = shapes.getItemAt(0).adjustments.get(0);
      const arc = shapes.getItemAt(1).adjustments.get(1);
      await context.sync();
      expect([round.value, arc.value]).toEqual([0.25, 45]);
      // The handle of roundRect stops at 50 000 (half the short side).
      shapes.getItemAt(0).adjustments.set(0, 3);
      const bounded = shapes.getItemAt(0).adjustments.get(0);
      await context.sync();
      expect(bounded.value).toBe(0.5);
    });
  });

  test('a table, a bad index, and a bad value refuse', async () => {
    const runtime = await open(rounded());
    const table = await rejection(
      runtime.run(async (context) => {
        const t = context.presentation.slides.getItemAt(0).shapes.addTable(1, 1);
        t.adjustments.load('count');
        await context.sync();
      })
    );
    expect([code(table), inner(table)]).toEqual(['NotSupported', 'edit.no-preset-geometry']);
    const index = await rejection(
      runtime.run(async (context) => {
        context.presentation.slides.getItemAt(0).shapes.getItemAt(0).adjustments.set(3, 0.1);
        await context.sync();
      })
    );
    expect(code(index)).toBe('InvalidArgument');
    const missing = await rejection(
      runtime.run(async (context) => {
        context.presentation.slides.getItemAt(0).shapes.getItemAt(0).adjustments.get(2);
        await context.sync();
      })
    );
    expect(code(missing)).toBe('ItemNotFound');
    await runtime.run(async (context) => {
      const adjustments = context.presentation.slides.getItemAt(0).shapes.getItemAt(0).adjustments;
      expect(code(await refusedAtSync(context, () => adjustments.set(-1, 0)))).toBe(
        'InvalidArgument'
      );
      expect(code(await refusedAtSync(context, () => adjustments.set(0, Number.NaN)))).toBe(
        'InvalidArgument'
      );
      expect(code(await refusedAtSync(context, () => adjustments.get(0.5)))).toBe(
        'InvalidArgument'
      );
    });
  });
});
