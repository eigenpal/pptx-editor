/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// getImageAsBase64 on a server runtime with a rasterizer: the engine paints the picture with the
// editor's own renderer into an SVG document, and the rasterizer draws it. The stub here records
// each request and answers a PNG header of the asked size.

import { describe, expect, test } from 'bun:test';
import type { ImageRasterRequest } from '@pptx-editor.dev/core/editor';
import { isPptxEditorError, PptxEditor } from '../index.ts';
import { deck, rejection } from './support.ts';

/** The PNG signature and an `IHDR` chunk: enough for the size check. */
function pngOf(width: number, height: number): Uint8Array {
  const b = new Uint8Array(33);
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52]);
  const view = new DataView(b.buffer);
  view.setUint32(16, width);
  view.setUint32(20, height);
  b.set([8, 6, 0, 0, 0], 24);
  return b;
}

const base64 = (bytes: Uint8Array): string => btoa(String.fromCharCode(...bytes));

const failure = (error: unknown) =>
  isPptxEditorError(error) ? [error.code, error.debugInfo.innerError?.code] : [String(error)];

describe('getImageAsBase64 on a server runtime with a rasterizer', () => {
  test('draws a slide and a shape through the rasterizer', async () => {
    const requests: ImageRasterRequest[] = [];
    const runtime = await PptxEditor.createServer(deck(), {
      substitutes: false,
      rasterize: async (request) => {
        requests.push(request);
        return pngOf(request.width, request.height);
      },
    });
    expect(runtime.capabilities.images).toBe(true);
    const [slide, shape] = await runtime.run(async (context) => {
      const first = context.presentation.slides.getItemAt(0);
      const slideImage = first.getImageAsBase64({ width: 480 });
      const shapeImage = first.shapes.getItem('2').getImageAsBase64({ format: 'Png' });
      await context.sync();
      return [slideImage.value, shapeImage.value];
    });
    expect(slide).toBe(base64(pngOf(480, 270)));
    // The box is three inches by one inch at 96 pixels per inch.
    expect(shape).toBe(base64(pngOf(288, 96)));
    expect(requests.map((r) => [r.width, r.height])).toEqual([
      [480, 270],
      [288, 96],
    ]);
    expect(requests[0]!.svg).toContain('Hello');
    expect(requests[1]!.svg).not.toContain('data-kind="background"');
    runtime.dispose();
  });

  test('a picture shows the writes of the syncs before it', async () => {
    const svgs: string[] = [];
    const runtime = await PptxEditor.createServer(deck(), {
      substitutes: false,
      rasterize: (request) => {
        svgs.push(request.svg);
        return pngOf(request.width, request.height);
      },
    });
    await runtime.run(async (context) => {
      const box = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      box.textFrame.textRange.text = 'Changed';
      await context.sync();
      box.getImageAsBase64();
      await context.sync();
    });
    expect(svgs[0]).toContain('Changed');
    runtime.dispose();
  });

  test('a rasterizer that fails or answers the wrong size refuses with a typed reason', async () => {
    for (const rasterize of [
      () => pngOf(1, 1),
      () => {
        throw new Error('no rasterizer');
      },
    ]) {
      const runtime = await PptxEditor.createServer(deck(), { substitutes: false, rasterize });
      const error = await rejection(
        runtime.run(async (context) => {
          context.presentation.slides.getItemAt(0).getImageAsBase64();
          await context.sync();
        })
      );
      expect(failure(error)).toEqual(['GeneralException', 'image.failed']);
      runtime.dispose();
    }
  });

  test('a rasterize option that is not a function refuses', async () => {
    const error = await rejection(
      PptxEditor.createServer(deck(), {
        substitutes: false,
        rasterize: 'png' as unknown as () => Uint8Array,
      })
    );
    expect(failure(error)[0]).toBe('InvalidArgument');
  });
});
