/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A runtime that borrows a live editor: selection members and one undo step per sync.

import { describe, expect, test } from 'bun:test';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { isPptxEditorError, PptxEditor } from '../browser.ts';
import { deck, rejection } from './support.ts';

async function editor(): Promise<PptxEditorInstance> {
  const instance = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await instance.load(deck())).ok).toBe(true);
  return instance;
}

describe('the browser runtime', () => {
  test('has a selection and no save', async () => {
    const runtime = PptxEditor.createBrowser(await editor());
    expect(runtime.capabilities).toMatchObject({ selection: true, save: false });
    expect('save' in runtime).toBe(false);
  });

  test('formats the selected text in one sync, as one undo step', async () => {
    const live = await editor();
    const runtime = PptxEditor.createBrowser(live);
    await runtime.run(async (context) => {
      const box = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      box.textFrame.textRange.getSubstring(6, 5).setSelected();
      await context.sync();
    });
    expect(live.snapshot().selection).toMatchObject({ kind: 'text' });
    await runtime.run(async (context) => {
      const range = context.presentation.getSelectedTextRange();
      range.font.color = '#C00000';
      range.load('text,start,length');
      await context.sync();
      expect([range.text, range.start, range.length]).toEqual(['World', 6, 5]);
      const font = range.font.load('color');
      await context.sync();
      expect(font.color).toBe('#C00000');
    });
    expect(live.exec({ type: 'undo' }).ok).toBe(true);
    await runtime.run(async (context) => {
      const font = context.presentation.slides
        .getItemAt(0)
        .shapes.getItem('2')
        .textFrame.textRange.getSubstring(6, 5)
        .font.load('color');
      await context.sync();
      expect(font.color).not.toBe('#C00000');
    });
  });

  test('a sync that splits before a read is still one undo step', async () => {
    const live = await editor();
    const runtime = PptxEditor.createBrowser(live);
    await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      shapes.addTextBox('Added', { left: 10, top: 10, width: 100, height: 20 });
      // A write to a lookup that follows the add: the sync splits before the lookup.
      const found = shapes.getItemOrNullObject('2');
      found.name = 'Renamed';
      await context.sync();
    });
    const names = async (): Promise<string[]> =>
      runtime.run(async (context) => {
        const shapes = context.presentation.slides.getItemAt(0).shapes.load('items/name');
        await context.sync();
        return shapes.items.map((s) => s.name);
      });
    expect(await names()).toContain('Renamed');
    expect(live.exec({ type: 'undo' }).ok).toBe(true);
    expect(live.snapshot().canUndo).toBe(false);
    expect(await names()).not.toContain('Renamed');
    expect((await names()).length).toBe(2);
  });

  test('a null selected text range when no text is selected', async () => {
    const runtime = PptxEditor.createBrowser(await editor());
    await runtime.run(async (context) => {
      const range = context.presentation.getSelectedTextRangeOrNullObject();
      range.load('text');
      await context.sync();
      expect(range.isNullObject).toBe(true);
    });
    const error = await rejection(
      runtime.run(async (context) => {
        context.presentation.getSelectedTextRange().load('text');
        await context.sync();
      })
    );
    expect(isPptxEditorError(error) && error.code).toBe('ItemNotFound');
  });

  test('selects and reads slides and shapes', async () => {
    const live = await editor();
    const runtime = PptxEditor.createBrowser(live);
    await runtime.run(async (context) => {
      const slides = context.presentation.slides.load('items/id');
      await context.sync();
      slides.items[0]!.setSelectedShapes(['2', '3']);
      await context.sync();
      const shapes = context.presentation.getSelectedShapes().load('items/name');
      await context.sync();
      expect(shapes.items.map((s) => s.name)).toEqual(['Box', 'Other']);
      const group = context.presentation.getSelectedShapes();
      group.load('items');
      await context.sync();
      const made = group.group().load('type');
      await context.sync();
      expect(made.type).toBe('Group');
      context.presentation.setSelectedSlides([slides.items[1]!.id]);
      await context.sync();
      const selected = context.presentation.getSelectedSlides().load('items/index');
      const first = context.presentation.getSelectedSlides().getItemAt(0).load('id');
      await context.sync();
      expect(selected.items.map((s) => s.index)).toEqual([1]);
      expect(first.id).toBe(slides.items[1]!.id);
      const exported = context.presentation.getSelectedSlides().exportAsBase64Presentation();
      await context.sync();
      expect(exported.value.length).toBeGreaterThan(100);
    });
    expect(live.snapshot().currentSlide).toBe(1);
  });

  test('groups the selected shapes without a load of the selection', async () => {
    const live = await editor();
    const runtime = PptxEditor.createBrowser(live);
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(0).setSelectedShapes(['2', '3']);
      await context.sync();
      const group = context.presentation.getSelectedShapes().group();
      group.name = 'Grouped';
      await context.sync();
      group.load('type,name');
      const members = group.group.shapes.load('items/name');
      await context.sync();
      expect([group.type, group.name]).toEqual(['Group', 'Grouped']);
      expect(members.items.map((s) => s.name)).toEqual(['Box', 'Other']);
    });
    // The read of the selection and the group are one undo step, as one sync.
    expect(live.exec({ type: 'undo' }).ok).toBe(true);
    await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes.load('items/name');
      await context.sync();
      expect(shapes.items.map((s) => s.name)).toEqual(['Box', 'Other']);
    });
  });

  test('selection reads see the selection writes queued before them in the same sync', async () => {
    const live = await editor();
    const runtime = PptxEditor.createBrowser(live);
    await runtime.run(async (context) => {
      const p = context.presentation;
      p.slides.getItemAt(0).setSelectedShapes(['3']);
      const shapes = p.getSelectedShapes().load('items/name');
      const box = p.slides.getItemAt(0).shapes.getItem('2').textFrame.textRange;
      box.getSubstring(6, 5).setSelected();
      const range = p.getSelectedTextRange();
      range.load('text');
      range.font.italic = true;
      p.setSelectedSlides(['257']);
      const slides = p.getSelectedSlides().load('items/id');
      await context.sync();
      expect(shapes.items.map((s) => s.name)).toEqual(['Other']);
      expect(range.text).toBe('World');
      expect(slides.items.map((s) => s.id)).toEqual(['257']);
      const italic = box.getSubstring(6, 5).font.load('italic');
      await context.sync();
      expect(italic.italic).toBe(true);
    });
  });

  test('refuses writes while the editor is open for viewing', async () => {
    const live = await editor();
    const runtime = PptxEditor.createBrowser(live);
    live.setMode('view');
    const error = await rejection(
      runtime.run(async (context) => {
        context.presentation.slides.getItemAt(0).shapes.getItem('2').name = 'No';
        await context.sync();
      })
    );
    expect(isPptxEditorError(error) && error.debugInfo.innerError?.code).toBe('edit.read-only');
  });
});
