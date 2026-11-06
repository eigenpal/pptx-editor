/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Table cells end to end: text runs, font, fill, alignment, indent, resize, split, and clear.

import { describe, expect, test } from 'bun:test';
import { isPptxEditorError, type RequestContext, type Table } from '../index.ts';
import { open, refusedAtSync, rejection, reopen, thrown } from './support.ts';

const code = (error: unknown): string | undefined =>
  isPptxEditorError(error) ? error.code : String(error);

/** A 3 by 3 table on the second slide, synced so its cells can be looked up. */
async function addTable(context: RequestContext): Promise<Table> {
  const table = context.presentation.slides
    .getItemAt(1)
    .shapes.addTable(3, 3, { left: 20, top: 20 })
    .getTable();
  await context.sync();
  return table;
}

/** The table on the second slide of a reopened presentation. */
function savedTable(context: RequestContext): Table {
  return context.presentation.slides.getItemAt(1).shapes.getItemAt(0).getTable();
}

describe('table cells', () => {
  test('text runs, font, fill, alignment, and indent survive a save', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const cell = (await addTable(context)).getCellOrNullObject(1, 1);
      cell.textRuns = [
        { text: 'Net ', font: { italic: true } },
        { text: '42', font: { bold: true, color: '#C00000', underline: 'Double' } },
      ];
      cell.horizontalAlignment = 'Right';
      cell.indentLevel = 2;
      cell.verticalAlignment = 'MiddleCentered';
      cell.fill.setSolidColor('#FFF2CC');
      await context.sync();
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const table = savedTable(context);
      const cell = table.getCellOrNullObject(1, 1);
      cell.load('text,horizontalAlignment,indentLevel,verticalAlignment');
      const whole = cell.font.load('bold');
      const fill = cell.fill.load('type,foregroundColor');
      const plain = table.getCellOrNullObject(0, 0).load('verticalAlignment,horizontalAlignment');
      await context.sync();
      expect(cell.text).toBe('Net 42');
      expect(cell.horizontalAlignment).toBe('Right');
      expect(cell.indentLevel).toBe(2);
      expect(cell.verticalAlignment).toBe('MiddleCentered');
      expect(whole.bold).toBeNull();
      expect([fill.type, fill.foregroundColor]).toEqual(['Solid', '#FFF2CC']);
      expect(plain.verticalAlignment).toBe('Top');
      const runs = table.getCellOrNullObject(1, 1).load('textRuns');
      await context.sync();
      expect(runs.textRuns.map((run) => run.text)).toEqual(['Net ', '42']);
      expect(runs.textRuns[0]!.font).toMatchObject({ italic: true, bold: false });
      expect(runs.textRuns[1]!.font).toMatchObject({
        bold: true,
        italic: false,
        color: '#C00000',
        underline: 'Double',
      });
      expect(runs.textRuns[0]!.font?.underline).toBe('None');
      cell.font.size = 9;
      await context.sync();
      const sized = cell.font.load('size');
      await context.sync();
      expect(sized.size).toBe(9);
    });
  });

  test('textRuns counts a CR LF pair as one paragraph break', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const table = await addTable(context);
      // Row 1: the first row of the default table style is bold.
      const cell = table.getCellOrNullObject(1, 0);
      cell.textRuns = [{ text: 'a\r\nb' }, { text: 'C', font: { bold: true } }];
      await context.sync();
      const plain = cell.load('text');
      const whole = cell.font.load('bold');
      await context.sync();
      expect(plain.text).toBe('a\rbC');
      // Only the last run is bold, so the cell's text reads as mixed.
      expect(whole.bold).toBeNull();
    });
  });

  test('textRuns read back writes the same text and formats', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const table = await addTable(context);
      const source = table.getCellOrNullObject(1, 0);
      source.textRuns = [{ text: 'a\rb' }, { text: 'C', font: { bold: true, size: 20 } }];
      const empty = table.getCellOrNullObject(2, 2).load('textRuns');
      await context.sync();
      source.load('textRuns');
      await context.sync();
      expect(source.textRuns.map((run) => run.text)).toEqual(['a\rb', 'C']);
      expect(empty.textRuns).toEqual([]);
      const copy = table.getCellOrNullObject(2, 1);
      copy.textRuns = source.textRuns;
      await context.sync();
      copy.load('textRuns, text');
      await context.sync();
      expect(copy.text).toBe('a\rbC');
      expect(copy.textRuns).toEqual(source.textRuns);
    });
  });

  test('textRuns refuses a bad run', async () => {
    const runtime = await open();
    const bad = await rejection(
      runtime.run(async (context) => {
        (await addTable(context)).getCellOrNullObject(0, 0).textRuns = [
          { text: 'x', font: { weight: 700 } as never },
        ];
        await context.sync();
      })
    );
    expect(code(bad)).toBe('InvalidArgument');
    await runtime.run(async (context) => {
      const cell = (await addTable(context)).getCellOrNullObject(0, 0);
      expect(code(thrown(() => cell.textRuns))).toBe('PropertyNotLoaded');
      await context.sync();
    });
  });

  test('resize grows and shrinks a span; split divides a cell', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const table = await addTable(context);
      table.getCellOrNullObject(0, 0).resize(2, 2);
      await context.sync();
      const grown = table
        .getCellOrNullObject(1, 1)
        .load('rowIndex,columnIndex,rowCount,columnCount');
      await context.sync();
      expect([grown.rowIndex, grown.columnIndex, grown.rowCount, grown.columnCount]).toEqual([
        0, 0, 2, 2,
      ]);
      table.getCellOrNullObject(0, 0).resize(1, 2);
      table.getCellOrNullObject(2, 2).split(1, 2);
      table.load('rowCount,columnCount');
      await context.sync();
      expect([table.rowCount, table.columnCount]).toEqual([3, 4]);
      const freed = table.getCellOrNullObject(1, 0).load('rowIndex,rowCount');
      await context.sync();
      expect([freed.rowIndex, freed.rowCount]).toEqual([1, 1]);
    });
    const error = await rejection(
      runtime.run(async (context) => {
        (await addTable(context)).getCellOrNullObject(0, 0).resize(0, 1);
        await context.sync();
      })
    );
    expect(code(error)).toBe('InvalidArgument');
    const overlap = await rejection(
      runtime.run(async (context) => {
        const table = await addTable(context);
        table.mergeCells(1, 1, 2, 2);
        table.getCellOrNullObject(0, 0).resize(2, 2);
        await context.sync();
      })
    );
    expect(code(overlap)).toBe('NotSupported');
    expect(isPptxEditorError(overlap) && overlap.debugInfo.innerError?.code).toBe(
      'edit.merge-overlap'
    );
  });

  test('clear removes text, format, or both', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const shape = context.presentation.slides.getItemAt(1).shapes.addTable(2, 2, {
        values: [
          ['a', 'b'],
          ['c', 'd'],
        ],
      });
      const table = shape.getTable();
      await context.sync();
      // The second row: the table style makes the first row bold.
      const first = table.getCellOrNullObject(1, 0);
      first.font.bold = true;
      first.fill.setSolidColor('#00B050');
      await context.sync();
      table.clear({ format: true });
      await context.sync();
      const font = first.font.load('bold');
      const fill = first.fill.load('foregroundColor');
      table.load('values');
      await context.sync();
      expect(table.values).toEqual([
        ['a', 'b'],
        ['c', 'd'],
      ]);
      expect(font.bold).toBe(false);
      expect(fill.foregroundColor).not.toBe('#00B050');
      table.clear();
      table.load('values');
      await context.sync();
      expect(table.values).toEqual([
        ['', ''],
        ['', ''],
      ]);
      expect(code(await refusedAtSync(context, () => table.clear({ bogus: true } as never)))).toBe(
        'InvalidArgument'
      );
    });
  });
});
