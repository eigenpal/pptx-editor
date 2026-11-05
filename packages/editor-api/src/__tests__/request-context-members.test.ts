/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The Office.js `RequestContext` members besides `load` and `presentation`: `sync` with a
// pass-through value, `application`, `debugInfo`, `trace`, and `loadRecursive`.

import { describe, expect, test } from 'bun:test';
import * as api from '../index.ts';
import { Application, isPptxEditorError, type Shape, type Slide } from '../index.ts';
import { TYPE_NAMES } from '../model/type-names.ts';
import { open, rejection, thrown } from './support.ts';

const failure = (error: unknown) =>
  isPptxEditorError(error) ? [error.code, error.debugInfo.innerError?.code] : [String(error)];

describe('RequestContext', () => {
  test('sync answers its pass-through value after the batch applies', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const count = context.presentation.slides.getCount();
      const passed = await context.sync({ tag: 'kept' });
      expect(passed).toEqual({ tag: 'kept' });
      expect(count.value).toBe(2);
      expect(await context.sync()).toBeUndefined();
    });
  });

  test('application is one object with nothing to load', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const application = context.application;
      expect(application).toBeInstanceOf(Application);
      expect(context.application).toBe(application);
      application.load();
      await context.sync();
      expect(application.toJSON()).toEqual({});
      expect(failure(thrown(() => application.load('name')))[0]).toBe('InvalidArgument');
    });
  });

  test('debugInfo lists the queued calls and traces, and empties at the sync', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      expect(context.debugInfo.pendingStatements).toEqual([]);
      context.trace('before');
      context.presentation.slides.getCount();
      context.trace('after');
      const pending = context.debugInfo.pendingStatements;
      expect(pending[0]).toBe('before');
      expect(pending.at(-1)).toBe('after');
      expect(pending).toContain('presentation.slides.getCount');
      await context.sync();
      expect(context.debugInfo.pendingStatements).toEqual([]);
    });
  });

  test('a failed sync carries the trace messages queued before the failed call', async () => {
    const runtime = await open();
    const error = await rejection(
      runtime.run(async (context) => {
        context.trace('read the count');
        context.presentation.slides.getCount();
        context.trace('look up slide 9');
        context.presentation.slides.getItemAt(9).load('id');
        context.trace('never reached');
        context.presentation.slides.getCount();
        await context.sync();
      })
    );
    expect(isPptxEditorError(error) && error.code).toBe('ItemNotFound');
    expect(isPptxEditorError(error) && error.traceMessages).toEqual([
      'read the count',
      'look up slide 9',
    ]);
    const plain = await rejection(
      runtime.run(async (context) => {
        context.presentation.slides.getItemAt(9).load('id');
        await context.sync();
      })
    );
    expect(isPptxEditorError(plain) && plain.traceMessages).toEqual([]);
    expect(failure(await rejection(runtime.run(async (c) => c.trace(1 as never))))[0]).toBe(
      'InvalidArgument'
    );
  });

  test('loadRecursive loads each type it reaches, in one sync', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const slides = context.presentation.slides;
      context.loadRecursive(slides, {
        SlideCollection: 'items',
        Slide: 'id,shapes',
        ShapeCollection: 'items',
        Shape: 'name,textFrame',
        TextFrame: 'hasText',
      });
      await context.sync();
      const names = slides.items.map((slide: Slide) =>
        slide.shapes.items.map((shape: Shape) => [shape.name, shape.textFrame.hasText])
      );
      expect(names).toEqual([
        [
          ['Box', true],
          ['Other', false],
        ],
        [],
      ]);
      expect(slides.items.every((slide) => typeof slide.id === 'string')).toBe(true);
    });
  });

  test('loadRecursive stops at maxDepth and at a type with no option', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const slides = context.presentation.slides;
      const options = {
        SlideCollection: 'items',
        Slide: 'shapes',
        ShapeCollection: 'items',
        Shape: 'textFrame',
        TextFrame: 'hasText',
      };
      context.loadRecursive(slides, options, 2);
      await context.sync();
      // Depth 2 reaches the shape collections; the text frames are two levels further down.
      const box = slides.items[0]!.shapes.items[0]!;
      expect(box.name).toBe('Box');
      expect(failure(thrown(() => box.textFrame.hasText))[0]).toBe('PropertyNotLoaded');
      context.loadRecursive(slides, options, 4);
      await context.sync();
      expect(slides.items[0]!.shapes.items[0]!.textFrame.hasText).toBe(true);
    });
    await runtime.run(async (context) => {
      const slides = context.presentation.slides;
      context.loadRecursive(slides, { SlideCollection: 'items', ShapeCollection: 'items' });
      await context.sync();
      // `items` loads the scalars of each slide; with no Slide option, the walk stops there.
      expect(typeof slides.items[0]!.id).toBe('string');
      expect(failure(thrown(() => slides.items[0]!.shapes.items))[0]).toBe('PropertyNotLoaded');
    });
  });

  test('loadRecursive knows the type name of every exported class of the model', () => {
    const classes = Object.entries(api).filter(
      ([name, value]) =>
        typeof value === 'function' &&
        name !== 'ClientObject' &&
        value.prototype instanceof api.ClientObject
    );
    expect(classes.length).toBeGreaterThan(50);
    expect(classes.map(([name]) => name).filter((name) => !TYPE_NAMES.includes(name))).toEqual([]);
  });

  test('loadRecursive refuses a bad type name, option, depth, or object', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const slides = context.presentation.slides;
      const typo = thrown(() => context.loadRecursive(slides, { Slides: 'id' }));
      expect(failure(typo)[0]).toBe('InvalidArgument');
      expect(isPptxEditorError(typo) && typo.message).toContain("'Slide'");
      expect(failure(thrown(() => context.loadRecursive(slides, { Slide: 7 as never })))[0]).toBe(
        'InvalidArgument'
      );
      expect(failure(thrown(() => context.loadRecursive(slides, {}, 33)))[0]).toBe(
        'InvalidArgument'
      );
      expect(failure(thrown(() => context.loadRecursive({} as never, {})))[0]).toBe(
        'InvalidArgument'
      );
    });
    const other = await runtime.run(async (context) => {
      const slides = context.presentation.slides;
      context.trackedObjects.add(slides);
      return slides;
    });
    await runtime.run(async (context) => {
      expect(failure(thrown(() => context.loadRecursive(other, {})))[0]).toBe('InvalidObjectPath');
    });
  });
});
