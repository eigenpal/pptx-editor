/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Stub drawing for the picture path in bun: a DOM with a stub canvas for the live editor, and a
// stub rasterizer for a server runtime.
//
// The editor paints the slide as SVG with the real paint, then draws it on a canvas. happy-dom
// has no canvas and decodes no images, so the stub decodes every image at once and answers a
// PNG header of the canvas size. The paint, the fit, and the crop to a shape stay real.

import { GlobalRegistrator } from '@happy-dom/global-registrator';

/** A PNG of `width` by `height`: the signature and the IHDR chunk, which is all `pngSize` reads. */
export function pngHeader(width: number, height: number): Uint8Array<ArrayBuffer> {
  const png = new Uint8Array(33);
  png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13], 0);
  png.set([0x49, 0x48, 0x44, 0x52], 12);
  const view = new DataView(png.buffer);
  view.setUint32(16, width);
  view.setUint32(20, height);
  png.set([8, 6, 0, 0, 0], 24);
  return png;
}

/**
 * A stub rasterizer for a server runtime (`createServer({ rasterize })`): the engine paints the
 * SVG document for real, and the stub answers a PNG header of the asked size.
 */
export function stubRasterize(request: { width: number; height: number }): Uint8Array {
  return pngHeader(request.width, request.height);
}

/** Run `body` with a global DOM whose canvas and image decode are stubs. */
export async function withStubDrawing<T>(body: () => Promise<T>): Promise<T> {
  const registered = GlobalRegistrator.isRegistered;
  if (!registered) GlobalRegistrator.register();
  const g = globalThis as unknown as Record<string, unknown>;
  const image = g.Image;
  const canvas = (g.HTMLCanvasElement as { prototype: Record<string, unknown> }).prototype;
  const { getContext, toBlob } = canvas;
  g.Image = class {
    onload: (() => void) | null = null;
    onerror: (() => void) | null = null;
    set src(_url: string) {
      queueMicrotask(() => this.onload?.());
    }
    decode(): Promise<void> {
      return Promise.resolve();
    }
  };
  canvas.getContext = () => ({ drawImage() {} });
  canvas.toBlob = function (this: { width: number; height: number }, done: (b: Blob) => void) {
    done(new Blob([pngHeader(this.width, this.height)], { type: 'image/png' }));
  };
  try {
    return await body();
  } finally {
    g.Image = image;
    canvas.getContext = getContext;
    canvas.toBlob = toBlob;
    if (!registered) await GlobalRegistrator.unregister();
  }
}
