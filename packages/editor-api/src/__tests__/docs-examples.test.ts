/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The examples of docs/site/content/editor-api, run as written against a synthetic deck.

import { expect, test } from 'bun:test';
import { PptxEditor, ShapeZOrder } from '../index.ts';
import { deck, reopen } from './support.ts';

test('the overview example edits and saves', async () => {
  const runtime = await PptxEditor.createServer(deck(), { substitutes: false });
  try {
    await runtime.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(0);
      const box = slide.shapes.addTextBox('Quarterly results', { left: 40, top: 40, width: 400 });
      box.textFrame.textRange.font.bold = true;
      const table = slide.shapes
        .addTable(2, 2, {
          values: [
            ['Region', 'Total'],
            ['North', '42'],
          ],
        })
        .getTable();
      await context.sync();

      const cell = table.getCellOrNullObject(1, 1);
      cell.horizontalAlignment = 'Right';
      cell.fill.setSolidColor('#FFF2CC');
      await context.sync();
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes.load('items/name,items/type');
      await context.sync();
      expect(shapes.items.map((s) => s.type)).toContain('Table');
    });
  } finally {
    runtime.dispose();
  }
});

test('the compatibility example renames shapes, and enums take both spellings', async () => {
  const runtime = await PptxEditor.createServer(deck(), { substitutes: false });
  await runtime.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    shapes.load('items/name');
    await context.sync();
    for (const shape of shapes.items) shape.name = shape.name.toUpperCase();
    shapes.items[0]!.setZOrder(ShapeZOrder.bringToFront);
    shapes.items[1]!.setZOrder('BringToFront');
    await context.sync();
    shapes.load('items/name');
    await context.sync();
    expect(shapes.items.map((s) => s.name)).toEqual(['BOX', 'OTHER']);
  });
  runtime.dispose();
});

test('the common mistakes examples behave as the docs say', async () => {
  const runtime = await PptxEditor.createServer(deck(), { substitutes: false });
  try {
    await runtime.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(0);
      // A read with no load.
      const box = slide.shapes.getItem('2');
      expect(() => box.name).toThrow();
      box.load('name');
      await context.sync();
      expect(box.name).toBe('Box');
      // A sync inside a loop: one sync for all the writes.
      const shapes = slide.shapes.load('items/name');
      await context.sync();
      for (const shape of shapes.items) shape.name = shape.name.trim();
      await context.sync();
      // A retry after a refusal: the same proxy works.
      let refused = false;
      try {
        box.fill.setSolidColor('notacolor');
        await context.sync();
      } catch {
        refused = true;
        box.fill.setSolidColor('#1F4E79');
        await context.sync();
      }
      expect(refused).toBe(true);
      const fill = box.fill.load('foregroundColor');
      await context.sync();
      expect(fill.foregroundColor).toBe('#1F4E79');
      // Bindings are in memory only.
      context.presentation.bindings.add(box, 'Shape', 'kpi');
      await context.sync();
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const all = context.presentation.bindings.load('items');
      await context.sync();
      expect(all.items).toEqual([]);
    });
  } finally {
    runtime.dispose();
  }
});
