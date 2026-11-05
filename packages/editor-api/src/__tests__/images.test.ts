/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Slide.getImageAsBase64 and Shape.getImageAsBase64: a server runtime refuses, and a runtime
// over a host that draws fills the result in the same sync. The host here draws fixed bytes; the
// drawing itself is the editor's slide picture paint, which needs a browser.

import { describe, expect, test } from 'bun:test';
import {
  createAutomationHost,
  storeAutomationPort,
  type AutomationBatchResponse,
  type AutomationHost,
} from '@pptx-editor.dev/core/automation';
import { PptxDocumentStore, readPptx } from '@pptx-editor.dev/core/store';
import { isPptxEditorError, ShapeGetImageFormatType } from '../index.ts';
import { createRuntime } from '../runtime/runtime.ts';
import { deck, open, refusedAtSync, rejection } from './support.ts';

const PNG = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
const BASE64 = btoa(String.fromCharCode(...PNG));

const code = (error: unknown): string | undefined =>
  isPptxEditorError(error) ? error.code : String(error);
const inner = (error: unknown): string | undefined =>
  isPptxEditorError(error) ? error.debugInfo.innerError?.code : undefined;

interface Target {
  readonly slideIndex: number;
  readonly shapeNodeId?: string;
  readonly width?: number;
  readonly height?: number;
}

/** A runtime over a host that draws: each request is recorded, and `draw` answers it. */
function drawingRuntime(
  draw: () => Promise<Uint8Array> = async () => PNG,
  answer: (response: AutomationBatchResponse) => AutomationBatchResponse = (r) => r
) {
  const read = readPptx(deck());
  if (!read.ok) throw new Error('fixture');
  const store = new PptxDocumentStore(read.pkg);
  const calls: Target[] = [];
  const host: AutomationHost = createAutomationHost({
    port: {
      ...storeAutomationPort(store),
      renderImage: (target) => {
        calls.push(target);
        return draw();
      },
    },
    capabilities: {
      document: true,
      save: true,
      events: true,
      selection: false,
      autofit: false,
      images: true,
    },
  });
  const answering: AutomationHost = {
    ...host,
    execute: (request) => answer(host.execute(request)),
  };
  return { runtime: createRuntime({ host: answering, save: true }), calls, store };
}

describe('getImageAsBase64', () => {
  test('a server runtime refuses with NotSupported and says why', async () => {
    const runtime = await open();
    expect(runtime.capabilities.images).toBe(false);
    const error = await rejection(
      runtime.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        slides.items[0]!.getImageAsBase64();
        await context.sync();
      })
    );
    expect(code(error)).toBe('NotSupported');
    expect(inner(error)).toBe('host.no-images');
  });

  test('a slide picture fills in the same sync', async () => {
    const { runtime, calls } = drawingRuntime();
    expect(runtime.capabilities.images).toBe(true);
    const value = await runtime.run(async (context) => {
      const image = context.presentation.slides.getItemAt(1).getImageAsBase64({ width: 320 });
      await context.sync();
      return image.value;
    });
    expect(value).toBe(BASE64);
    expect(calls).toEqual([{ slideIndex: 1, width: 320 }]);
  });

  test('a shape picture takes the PNG format and a box', async () => {
    const { runtime, calls } = drawingRuntime();
    const value = await runtime.run(async (context) => {
      const shape = context.presentation.slides.getItemAt(0).shapes.getItem('3');
      const a = shape.getImageAsBase64({ format: 'Png', width: 10, height: 20 });
      const b = shape.getImageAsBase64({ format: ShapeGetImageFormatType.png });
      await context.sync();
      return [a.value, b.value];
    });
    expect(value).toEqual([BASE64, BASE64]);
    expect(calls).toHaveLength(2);
    expect(calls[0]).toMatchObject({ slideIndex: 0, width: 10, height: 20 });
    expect(typeof calls[0]!.shapeNodeId).toBe('string');
  });

  test('wrong options refuse at the sync', async () => {
    const { runtime, calls } = drawingRuntime();
    await runtime.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(0);
      const shape = slide.shapes.getItem('2');
      for (const options of [{ width: 0 }, { height: -1 }, { width: Number.NaN }, 'big'])
        expect(
          code(await refusedAtSync(context, () => slide.getImageAsBase64(options as never)))
        ).toBe('InvalidArgument');
      expect(
        code(
          await refusedAtSync(context, () => shape.getImageAsBase64({ format: 'Jpeg' as never }))
        )
      ).toBe('InvalidArgument');
      await context.sync();
    });
    expect(calls).toEqual([]);
  });

  test('a renderer failure refuses the sync with a typed reason', async () => {
    const { runtime } = drawingRuntime(() => Promise.reject(new Error('canvas')));
    const error = await rejection(
      runtime.run(async (context) => {
        context.presentation.slides.getItemAt(0).getImageAsBase64();
        await context.sync();
      })
    );
    expect(code(error)).toBe('GeneralException');
    expect(isPptxEditorError(error) && error.target).toBe(
      'presentation.slides.getItemAt(0).getImageAsBase64'
    );
  });

  test('after writes in the same sync, a failed picture says the writes stay', async () => {
    const { runtime, store } = drawingRuntime(() => Promise.reject(new Error('canvas')));
    const before = store.revision;
    const error = await rejection(
      runtime.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(0);
        slide.shapes.getItem('2').name = 'Renamed';
        slide.getImageAsBase64();
        await context.sync();
      })
    );
    expect(store.revision).toBeGreaterThan(before);
    expect(inner(error)).toBe('edit.read-after-commit');
  });

  test('a shape of a layout has no picture', async () => {
    const { runtime } = drawingRuntime();
    const error = await rejection(
      runtime.run(async (context) => {
        const shapes = context.presentation.slides.getItemAt(1).layout.shapes;
        shapes.load('items');
        await context.sync();
        shapes.items[0]!.getImageAsBase64();
        await context.sync();
      })
    );
    expect(code(error)).toBe('NotSupported');
    expect(inner(error)).toBe('image.not-on-slide');
  });

  test('a picture that fails after a later answer failed is not an unhandled rejection', async () => {
    const unhandled: unknown[] = [];
    const listen = (reason: unknown): void => void unhandled.push(reason);
    process.on('unhandledRejection', listen);
    try {
      // The picture's drawing fails later; the count after it answers with the wrong kind, so
      // the sync fails before it awaits the picture.
      const { runtime } = drawingRuntime(
        () => new Promise((_, reject) => setTimeout(() => reject(new Error('canvas')), 5)),
        (response) => ({
          ...response,
          results: response.results.map((r, i) =>
            i === response.results.length - 1 && r.status === 'ok'
              ? { status: 'ok' as const, value: { kind: 'applied' as const } }
              : r
          ),
        })
      );
      const error = await rejection(
        runtime.run(async (context) => {
          context.presentation.slides.getItemAt(0).getImageAsBase64();
          context.presentation.slides.getCount();
          await context.sync();
        })
      );
      expect(code(error)).toBe('GeneralException');
      await new Promise((resolve) => setTimeout(resolve, 30));
      expect(unhandled).toEqual([]);
    } finally {
      process.off('unhandledRejection', listen);
    }
  });
});
