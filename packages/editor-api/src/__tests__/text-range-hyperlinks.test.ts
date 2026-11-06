/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `TextRange.hyperlinks`: the text links that overlap a range, as a `HyperlinkScopedCollection`.

import { describe, expect, test } from 'bun:test';
import { isPptxEditorError, type PptxEditorServerRuntime, type RequestContext } from '../index.ts';
import { open, rejection, reopen } from './support.ts';

const box = (context: RequestContext) =>
  context.presentation.slides.getItemAt(0).shapes.getItem('2');

/** The test deck with "Hello" and "World" (the text "Hello\rWorld") linked apart. */
async function linked(): Promise<PptxEditorServerRuntime> {
  const runtime = await open();
  await runtime.run(async (context) => {
    const range = box(context).textFrame.textRange;
    range.getSubstring(0, 5).setHyperlink({ address: 'https://a.example/', screenTip: 'A' });
    range.getSubstring(6, 5).setHyperlink({ address: 'https://b.example/' });
    box(context).setHyperlink({ address: 'https://shape.example/' });
    await context.sync();
  });
  return reopen(runtime);
}

describe('TextRange.hyperlinks', () => {
  test('reads the text links of the whole text, without the shape link', async () => {
    const runtime = await linked();
    await runtime.run(async (context) => {
      const links = box(context).textFrame.textRange.hyperlinks.load('items');
      const count = box(context).textFrame.textRange.hyperlinks.getCount();
      const second = box(context).textFrame.textRange.hyperlinks.getItemAt(1).load('address');
      await context.sync();
      expect(links.items.map((l) => [l.type, l.address, l.screenTip])).toEqual([
        ['TextRange', 'https://a.example/', 'A'],
        ['TextRange', 'https://b.example/', ''],
      ]);
      expect(count.value).toBe(2);
      expect(second.address).toBe('https://b.example/');
    });
  });

  test('loads through a path from the range', async () => {
    const runtime = await linked();
    await runtime.run(async (context) => {
      const range = box(context).textFrame.textRange.load('hyperlinks/items/address');
      await context.sync();
      expect(range.hyperlinks.items.map((l) => l.address)).toEqual([
        'https://a.example/',
        'https://b.example/',
      ]);
    });
  });

  test('a sub-range lists only the links it overlaps, each with its whole text', async () => {
    const runtime = await linked();
    await runtime.run(async (context) => {
      const range = box(context).textFrame.textRange;
      const inside = range.getSubstring(2, 2).hyperlinks.load('items');
      const between = range.getSubstring(5, 1).hyperlinks.getCount();
      const across = range.getSubstring(3, 5).hyperlinks.getCount();
      await context.sync();
      expect(inside.items.map((l) => l.address)).toEqual(['https://a.example/']);
      expect([between.value, across.value]).toEqual([0, 2]);
      const text = inside.items[0]!.getLinkedTextRangeOrNullObject().load('text');
      const shape = inside.items[0]!.getLinkedShapeOrNullObject();
      await context.sync();
      expect(text.text).toBe('Hello');
      expect(shape.isNullObject).toBe(true);
    });
  });

  test('a link deleted through a sub-range goes, and the other link stays', async () => {
    const runtime = await linked();
    await runtime.run(async (context) => {
      box(context).textFrame.textRange.getSubstring(1, 1).hyperlinks.getItemAt(0).delete();
      await context.sync();
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const text = box(context).textFrame.textRange.hyperlinks.load('items/address');
      const slide = context.presentation.slides.getItemAt(0).hyperlinks.load('items/type');
      const words = box(context).textFrame.textRange.load('text');
      await context.sync();
      expect(text.items.map((l) => l.address)).toEqual(['https://b.example/']);
      // A slide lists the shape link of a shape before its text links.
      expect(slide.items.map((l) => l.type)).toEqual(['Shape', 'TextRange']);
      expect(words.text).toBe('Hello\rWorld');
    });
  });

  test('sees a link written earlier in the same sync', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const range = box(context).textFrame.textRange.getSubstring(0, 5);
      range.setHyperlink({ address: 'https://a.example/' });
      const count = range.hyperlinks.getCount();
      await context.sync();
      expect(count.value).toBe(1);
    });
  });

  test('getItemAt past the end refuses with ItemNotFound', async () => {
    const runtime = await linked();
    const error = await rejection(
      runtime.run(async (context) => {
        box(context).textFrame.textRange.getSubstring(0, 2).hyperlinks.getItemAt(1).load();
        await context.sync();
      })
    );
    expect(isPptxEditorError(error) && error.code).toBe('ItemNotFound');
  });
});
