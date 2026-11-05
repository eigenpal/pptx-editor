/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `Presentation.bindings`: ids that name shapes for the lifetime of a runtime, never saved.

import { describe, expect, test } from 'bun:test';
import { createPptxEditor } from '@pptx-editor.dev/core/editor';
import { PptxEditor as BrowserEditor } from '../browser.ts';
import { BindingType, isPptxEditorError, type RequestContext } from '../index.ts';
import { deck, open, rejection, reopen } from './support.ts';

const shape = (context: RequestContext, id: string) =>
  context.presentation.slides.getItemAt(0).shapes.getItem(id);

const failure = (error: unknown) =>
  isPptxEditorError(error)
    ? [error.code, error.debugInfo.innerError?.code]
    : ['not a PptxEditorError', String(error)];

describe('Presentation.bindings', () => {
  test('adds bindings, reads them back by id and position, and finds their shapes', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const bindings = context.presentation.bindings;
      const added = bindings.add(shape(context, '2'), BindingType.shape, 'title');
      bindings.add(shape(context, '3'), 'Shape', 'note');
      const sameBatch = added.getShape().load('name');
      await context.sync();
      expect([added.id, added.type, sameBatch.name]).toEqual(['title', 'Shape', 'Box']);
      const count = bindings.getCount();
      const byId = bindings.getItem('note').getShape().load('name');
      const byIndex = bindings.getItemAt(0).load('id,type');
      const none = bindings.getItemOrNullObject('missing');
      const all = bindings.load('items');
      await context.sync();
      expect(count.value).toBe(2);
      expect(byId.name).toBe('Other');
      expect([byIndex.id, byIndex.type]).toEqual(['title', 'Shape']);
      expect(none.isNullObject).toBe(true);
      expect(all.items.map((b) => b.id)).toEqual(['title', 'note']);
    });
  });

  test('keep their bindings across runs of one runtime', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      context.presentation.bindings.add(shape(context, '2'), 'Shape', 'title');
      await context.sync();
    });
    await runtime.run(async (context) => {
      const found = context.presentation.bindings.getItem('title').getShape().load('id');
      await context.sync();
      expect(found.id).toBe('2');
    });
  });

  test('add with an id that is in use moves the binding to the new shape', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const bindings = context.presentation.bindings;
      bindings.add(shape(context, '2'), 'Shape', 'title');
      bindings.add(shape(context, '3'), 'Shape', 'title');
      // A lookup that a navigation waits for goes before the writes: sync between them.
      await context.sync();
      const count = bindings.getCount();
      const found = bindings.getItem('title').getShape().load('id');
      await context.sync();
      expect([count.value, found.id]).toEqual([1, '3']);
    });
  });

  test('delete removes a binding found by position, and leaves its shape', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const bindings = context.presentation.bindings;
      bindings.add(shape(context, '2'), 'Shape', 'title');
      bindings.add(shape(context, '3'), 'Shape', 'note');
      await context.sync();
      bindings.getItemAt(0).delete();
      await context.sync();
      const left = bindings.load('items/id');
      const shapes = context.presentation.slides.getItemAt(0).shapes.getCount();
      await context.sync();
      expect(left.items.map((b) => b.id)).toEqual(['note']);
      expect(shapes.value).toBe(2);
    });
  });

  test('a binding whose shape is deleted refuses getShape with ItemNotFound', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      context.presentation.bindings.add(shape(context, '2'), 'Shape', 'title');
      shape(context, '2').delete();
      await context.sync();
    });
    const error = await rejection(
      runtime.run(async (context) => {
        context.presentation.bindings.getItem('title').getShape().load('name');
        await context.sync();
      })
    );
    expect(failure(error)).toEqual(['ItemNotFound', 'binding.shape-deleted']);
    await runtime.run(async (context) => {
      const count = context.presentation.bindings.getCount();
      await context.sync();
      expect(count.value).toBe(1);
    });
  });

  test('are never saved with the presentation', async () => {
    const runtime = await open();
    const before = await runtime.save();
    await runtime.run(async (context) => {
      context.presentation.bindings.add(shape(context, '2'), 'Shape', 'binding-id-not-in-file');
      await context.sync();
    });
    const after = await runtime.save();
    expect(after).toEqual(before);
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const count = context.presentation.bindings.getCount();
      await context.sync();
      expect(count.value).toBe(0);
    });
  });

  test('refuse unknown ids, a bad type, and a selection on a server', async () => {
    const runtime = await open();
    const missing = await rejection(
      runtime.run(async (context) => {
        context.presentation.bindings.getItem('missing').load();
        await context.sync();
      })
    );
    expect(failure(missing)[0]).toBe('ItemNotFound');
    const badType = await rejection(
      runtime.run(async (context) => {
        context.presentation.bindings.add(shape(context, '2'), 'Table' as 'Shape', 'x');
        await context.sync();
      })
    );
    expect(failure(badType)[0]).toBe('InvalidArgument');
    const server = await rejection(
      runtime.run(async (context) => {
        context.presentation.bindings.addFromSelection('Shape', 'x');
        await context.sync();
      })
    );
    expect(failure(server)).toEqual(['NotSupported', 'host.no-selection']);
  });

  test('addFromSelection binds the selected shape in a browser runtime', async () => {
    const live = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await live.load(deck())).ok).toBe(true);
    const runtime = BrowserEditor.createBrowser(live);
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(0).setSelectedShapes(['3']);
      await context.sync();
      const bound = context.presentation.bindings.addFromSelection(BindingType.shape, 'picked');
      const found = bound.getShape().load('name');
      await context.sync();
      expect(found.name).toBe('Other');
    });
    const error = await rejection(
      runtime.run(async (context) => {
        context.presentation.slides.getItemAt(0).setSelectedShapes(['2', '3']);
        await context.sync();
        context.presentation.bindings.addFromSelection('Shape', 'two');
        await context.sync();
      })
    );
    expect(failure(error)).toEqual(['InvalidArgument', 'binding.selection']);
    runtime.dispose();
  });
});
