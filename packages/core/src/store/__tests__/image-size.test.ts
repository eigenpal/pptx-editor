import { describe, expect, test } from 'bun:test';
import { imageSize, imageSizeEmu } from '../index.ts';

/** A PNG signature, IHDR, an optional pHYs, and IEND. CRCs are zero: the reader never checks them. */
function png(w: number, h: number, ppm?: number): Uint8Array {
  const chunk = (type: string, data: number[]): number[] => {
    const len = [24, 16, 8, 0].map((s) => (data.length >>> s) & 255);
    return [...len, ...[...type].map((c) => c.charCodeAt(0)), ...data, 0, 0, 0, 0];
  };
  const u32 = (v: number): number[] => [24, 16, 8, 0].map((s) => (v >>> s) & 255);
  return new Uint8Array([
    0x89,
    0x50,
    0x4e,
    0x47,
    0x0d,
    0x0a,
    0x1a,
    0x0a,
    ...chunk('IHDR', [...u32(w), ...u32(h), 8, 6, 0, 0, 0]),
    ...(ppm === undefined ? [] : chunk('pHYs', [...u32(ppm), ...u32(ppm), 1])),
    ...chunk('IEND', []),
  ]);
}

describe('image header sizes', () => {
  test('PNG with and without a resolution', () => {
    expect(imageSize(png(128, 96))).toEqual({ width: 128, height: 96 });
    const hi = imageSize(png(300, 150, 5906))!; // 150 DPI
    expect(hi.width).toBe(300);
    expect(hi.dpiX).toBe(150);
    // 144 DPI is 5669.29 pixels per metre: the whole count reads back as whole DPI.
    expect(imageSize(png(128, 96, 5669))!.dpiY).toBe(144);
    // A count beyond that rounding keeps its exact DPI: 5680 per metre is 144.272 DPI.
    expect(imageSize(png(128, 96, 5680))!.dpiX).toBeCloseTo(144.272, 6);
    // A tiny count stays rejected: 20 and 39 per metre are under 1 DPI.
    expect(imageSize(png(128, 96, 20))).toEqual({ width: 128, height: 96 });
    expect(imageSize(png(128, 96, 39))).toEqual({ width: 128, height: 96 });
    expect(imageSizeEmu(imageSize(png(96, 48))!)).toEqual({ cx: 914_400, cy: 457_200 });
  });

  test('JPEG frame header and JFIF density', () => {
    const jfif = [0xff, 0xe0, 0, 16, 0x4a, 0x46, 0x49, 0x46, 0, 1, 1, 1, 0, 72, 0, 72, 0, 0];
    const sof = [0xff, 0xc0, 0, 11, 8, 0, 50, 0, 80, 1, 1, 0x11, 0];
    const size = imageSize(new Uint8Array([0xff, 0xd8, ...jfif, ...sof]));
    expect(size).toEqual({ width: 80, height: 50, dpiX: 72, dpiY: 72 });
  });

  test('GIF and BMP', () => {
    expect(imageSize(new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 10, 0, 20, 0]))).toEqual({
      width: 10,
      height: 20,
    });
    const bmp = new Uint8Array(54);
    bmp.set([0x42, 0x4d]);
    const v = new DataView(bmp.buffer);
    v.setInt32(18, 40, true);
    v.setInt32(22, -30, true);
    v.setInt32(38, 3780, true);
    v.setInt32(42, 3780, true);
    const s = imageSize(bmp)!;
    expect([s.width, s.height, s.dpiX]).toEqual([40, 30, 96]);
  });

  test('truncated, oversized, and unknown input', () => {
    expect(imageSize(png(10, 10).subarray(0, 20))).toBeNull();
    expect(imageSize(new Uint8Array([0xff, 0xd8, 0xff, 0xc0, 0, 1]))).toBeNull();
    expect(imageSize(new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0, 0, 0, 0]))).toBeNull();
    expect(imageSize(new Uint8Array([1, 2, 3, 4]))).toBeNull();
  });
});
