import { describe, expect, test } from 'bun:test';
import { PNG } from 'pngjs';
import { comparePixels } from './compare.ts';
import { downsample, downsampledSize } from './supersample.ts';

const image = (w: number, h: number, px: (x: number, y: number) => number[]): PNG => {
  const p = new PNG({ width: w, height: h });
  for (let y = 0; y < h; y += 1)
    for (let x = 0; x < w; x += 1) p.data.set(px(x, y), (y * w + x) * 4);
  return p;
};
const pixel = (p: PNG, x: number, y: number): number[] => [
  ...p.data.subarray((y * p.width + x) * 4, (y * p.width + x) * 4 + 4),
];

describe('downsample', () => {
  test('each output pixel is the mean of its block', () => {
    // A black row at the top of each 4 × 4 block covers a quarter of it.
    const out = downsample(
      image(8, 4, (_, y) => (y === 0 ? [0, 0, 0, 255] : [255, 255, 255, 255])),
      4
    );
    expect([out.width, out.height]).toEqual([2, 1]);
    expect(pixel(out, 0, 0)).toEqual([191, 191, 191, 255]);
    expect(pixel(out, 1, 0)).toEqual([191, 191, 191, 255]);
  });

  test('a glyph edge that snaps by one device pixel moves the filtered edge by 1/scale', () => {
    // An edge at device row 4 against row 5 at scale 4: 1/4 of a comparison pixel apart.
    const edge = (row: number): PNG =>
      downsample(
        image(4, 8, (_, y) => (y < row ? [0, 0, 0, 255] : [255, 255, 255, 255])),
        4
      );
    expect(pixel(edge(4), 0, 1)[0]).toBe(255);
    expect(pixel(edge(5), 0, 1)[0]).toBe(191);
  });

  test('a block that the image does not fill averages the pixels it has', () => {
    const out = downsample(
      image(5, 5, (x) => (x === 4 ? [0, 0, 0, 255] : [255, 255, 255, 255])),
      4
    );
    expect(downsampledSize({ width: 5, height: 5 }, 4)).toEqual({ width: 2, height: 2 });
    expect([out.width, out.height]).toEqual([2, 2]);
    expect(pixel(out, 1, 1)).toEqual([0, 0, 0, 255]);
    expect(pixel(out, 0, 0)).toEqual([255, 255, 255, 255]);
  });

  test('color is weighted by alpha', () => {
    const out = downsample(
      image(2, 1, (x) => (x === 0 ? [255, 0, 0, 255] : [0, 0, 255, 0])),
      2
    );
    expect(pixel(out, 0, 0)).toEqual([255, 0, 0, 128]);
  });

  test('scale 1 returns the image; other scales must be whole numbers', () => {
    const p = image(3, 3, () => [1, 2, 3, 4]);
    expect(downsample(p, 1)).toBe(p);
    expect(() => downsample(p, 1.5)).toThrow(RangeError);
    expect(() => downsample(p, 0)).toThrow(RangeError);
  });

  test('the comparison size of a short raster stays short by the same whole pixels', () => {
    // A page side of 799.5 px at scale 1 rasterizes to 800 px; at scale 4, to 3198 px.
    expect(downsampledSize({ width: Math.ceil(799.5 * 4), height: 4320 }, 4)).toEqual({
      width: 800,
      height: 1080,
    });
  });
});

describe('anti-aliasing judged at the device scale', () => {
  // A horizontal stroke 12 device pixels (3 comparison pixels, 1.5 pt) thick on white, as a
  // rasterizer draws it at device scale 4, with its top edge at device row `top`. With `ramp`,
  // each edge has a 1 device pixel anti-aliased row, as a glyph edge has.
  const stroke = (top: number, ramp = false): PNG =>
    image(64, 64, (_, y) => {
      if (ramp && (y === top - 1 || y === top + 12)) return [128, 128, 128, 255];
      return y >= top && y < top + 12 ? [0, 0, 0, 255] : [255, 255, 255, 255];
    });
  const filtered = (p: PNG): Buffer => PNG.sync.write(downsample(p, 4));
  const score = (render: PNG, reference: PNG): number =>
    comparePixels(filtered(render), filtered(reference), {
      device: { render: PNG.sync.write(render), reference: PNG.sync.write(reference), scale: 4 },
    }).mismatch;

  test('a stroke moved by one comparison pixel counts, though the filter makes it look like anti-aliasing', () => {
    // Judged on the filtered images alone, pixelmatch calls every changed pixel anti-aliasing.
    expect(comparePixels(filtered(stroke(22)), filtered(stroke(26))).mismatch).toBe(0);
    expect(score(stroke(22), stroke(26))).toBeGreaterThan(0);
  });

  test('an anti-aliased edge that snaps by one device pixel does not count', () => {
    expect(score(stroke(22, true), stroke(23, true))).toBe(0);
    expect(score(stroke(22), stroke(22))).toBe(0);
  });
});
