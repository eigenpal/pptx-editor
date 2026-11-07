// EMF and WMF signatures and picture frames, and WebP sizes, from the header alone.

import { describe, expect, test } from 'bun:test';
import {
  emfRectangle,
  webpHeader,
  wmfRectangle,
} from '../../../../../scripts/fixtures/picture-bytes.ts';
import { imageSize, metafileFormat, metafileSizeEmu } from '../index.ts';

describe('metafile headers', () => {
  test('the signature gives the format', () => {
    expect(metafileFormat(emfRectangle())).toBe('emf');
    expect(metafileFormat(wmfRectangle())).toBe('wmf');
    // A WMF with no placeable header starts with the standard header.
    expect(metafileFormat(wmfRectangle().subarray(22))).toBe('wmf');
  });

  test('other bytes, and a header with a wrong field, are not a metafile', () => {
    const emf = emfRectangle();
    expect(metafileFormat(emf.subarray(0, 40))).toBeUndefined();
    const unsigned = emf.slice();
    unsigned[40] = 0;
    expect(metafileFormat(unsigned)).toBeUndefined();
    const oddSize = emf.slice();
    oddSize[4] = 90;
    expect(metafileFormat(oddSize)).toBeUndefined();
    const wmf = wmfRectangle().slice(22);
    wmf[2] = 10;
    expect(metafileFormat(wmf)).toBeUndefined();
    expect(metafileFormat(new Uint8Array([0x89, 0x50, 0x4e, 0x47]))).toBeUndefined();
    expect(metafileFormat(new Uint8Array())).toBeUndefined();
  });

  test('the picture frame is the natural size', () => {
    // 200 × 100 pixels at 96 DPI: 5292 × 2646 hundredths of a millimetre.
    expect(metafileSizeEmu(emfRectangle())).toEqual({ cx: 5292 * 360, cy: 2646 * 360 });
    // 1440 × 720 units at 1440 units per inch.
    expect(metafileSizeEmu(wmfRectangle())).toEqual({ cx: 914_400, cy: 457_200 });
  });

  test('an empty, huge, or absent frame gives no size', () => {
    expect(metafileSizeEmu(emfRectangle(200, 100, 0, { right: 0, bottom: 0 }))).toBeUndefined();
    const huge = { right: 0x7fffffff, bottom: 0x7fffffff };
    expect(metafileSizeEmu(emfRectangle(200, 100, 0, huge))).toBeUndefined();
    expect(metafileSizeEmu(wmfRectangle(1440, 720, 0))).toBeUndefined();
    expect(metafileSizeEmu(wmfRectangle().subarray(22))).toBeUndefined();
  });
});

describe('WebP sizes', () => {
  test('each chunk kind gives the size', () => {
    expect(imageSize(webpHeader('VP8 ', 640, 480))).toEqual({ width: 640, height: 480 });
    expect(imageSize(webpHeader('VP8L', 1000, 3000))).toEqual({ width: 1000, height: 3000 });
    expect(imageSize(webpHeader('VP8X', 20_000, 9))).toEqual({ width: 20_000, height: 9 });
  });

  test('a damaged or oversized header gives no size', () => {
    const lossy = webpHeader('VP8 ', 64, 64);
    lossy[23] = 0;
    expect(imageSize(lossy)).toBeNull();
    expect(imageSize(webpHeader('VP8X', 70_000, 10))).toBeNull();
    expect(imageSize(webpHeader('VP8L', 64, 64).subarray(0, 24))).toBeNull();
  });
});
