import { describe, expect, test } from 'bun:test';
import pixelmatch from 'pixelmatch';
import { PNG } from 'pngjs';
import { comparePixels } from './compare.ts';
import { deviceAntialiased, downsample, downsampledSize } from './supersample.ts';
import { THRESHOLDS } from './thresholds.ts';

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

describe('device anti-aliasing by tiles', () => {
  // The whole-image method that `deviceAntialiased` replaces: one pixelmatch call over every
  // device pixel, with the blocks under red pixels unmasked.
  const wholeImage = (diff: PNG, a: PNG, b: PNG, scale: number): { n: number; diff: Buffer } => {
    const out = new Uint8Array(a.width * a.height * 4);
    const ignore = new Uint8Array(a.width * a.height).fill(1);
    const red: number[] = [];
    for (let y = 0; y < diff.height; y += 1)
      for (let x = 0; x < diff.width; x += 1) {
        const i = (y * diff.width + x) * 4;
        if (diff.data[i] === 255 && diff.data[i + 1] === 0 && diff.data[i + 2] === 0) {
          red.push(x, y);
          for (let yy = y * scale; yy < (y + 1) * scale; yy += 1)
            ignore.fill(0, yy * a.width + x * scale, yy * a.width + (x + 1) * scale);
        }
      }
    pixelmatch(a.data, b.data, out, a.width, a.height, {
      threshold: THRESHOLDS.pixelThreshold,
      includeAA: false,
      alpha: 0,
      diffColor: [255, 0, 0],
      diffColorAlt: [255, 0, 0],
      aaColor: [255, 255, 0],
      ignoreMask: ignore,
    });
    const copy = PNG.sync.write(diff);
    const marked = PNG.sync.read(copy);
    let n = 0;
    for (let k = 0; k < red.length; k += 2) {
      let real = 0;
      let aa = 0;
      for (let yy = red[k + 1] * scale; yy < (red[k + 1] + 1) * scale; yy += 1)
        for (let xx = red[k] * scale; xx < (red[k] + 1) * scale; xx += 1) {
          const j = (yy * a.width + xx) * 4;
          if (out[j] === 255 && out[j + 1] === 0 && out[j + 2] === 0) real += 1;
          else if (out[j] === 255 && out[j + 1] === 255 && out[j + 2] === 0) aa += 1;
        }
      if (real > aa) n += 1;
      else marked.data.set([255, 255, 0], (red[k + 1] * diff.width + red[k]) * 4);
    }
    return { n, diff: PNG.sync.write(marked) };
  };

  /** The top-left `w` × `h` px of an image. */
  const crop = (p: PNG, w: number, h: number): PNG =>
    image(w, h, (x, y) => [...p.data.subarray((y * p.width + x) * 4, (y * p.width + x) * 4 + 4)]);

  // Seeded noise: the same images on every run.
  let seed = 7;
  const random = (): number => (seed = (seed * 1103515245 + 12345) % 2 ** 31) / 2 ** 31;

  test('every block gets the same class as in a whole-image comparison, across tile edges', () => {
    const scale = 4;
    // 300 × 260 device px: tiles of 128 device px, with partial tiles at the right and bottom.
    const w = 300;
    const h = 260;
    let counted = 0;
    let excused = 0;
    for (let round = 0; round < 6; round += 1) {
      // Anti-aliased strokes, some of them across tile edges, moved by 0 to 3 device pixels.
      const strokes = Array.from({ length: 12 }, () => ({
        x: Math.floor(random() * w),
        y: Math.floor(random() * h),
        len: 10 + Math.floor(random() * 120),
        across: random() < 0.5,
        shift: Math.floor(random() * 4),
      }));
      // In odd rounds the render is larger than the reference by a few device pixels, as a
      // size within `maxSizeDeltaPx` may be: only the overlap is compared.
      const [dw, dh] = round % 2 ? [6, 5] : [0, 0];
      const draw = (moved: boolean): PNG =>
        image(w + (moved ? dw : 0), h + (moved ? dh : 0), (x, y) => {
          for (const s of strokes) {
            const top = s.y + (moved ? s.shift : 0);
            const [u, v, start] = s.across ? [x, y, s.x] : [y, x, s.x];
            if (u >= start && u < start + s.len) {
              if (v >= top && v < top + 6) return [0, 0, 0, 255];
              if (v === top - 1 || v === top + 6) return [128, 128, 128, 255];
            }
          }
          return [255, 255, 255, 255];
        });
      const a = draw(false);
      const b = draw(true);
      const fa = downsample(a, scale);
      const fb = crop(downsample(b, scale), fa.width, fa.height);
      const diff = new PNG({ width: fa.width, height: fa.height });
      pixelmatch(fa.data, fb.data, diff.data, fa.width, fa.height, {
        threshold: THRESHOLDS.pixelThreshold,
        includeAA: true,
      });
      const want = wholeImage(diff, a, crop(b, w, h), scale);
      const red = (p: PNG): number => {
        let r = 0;
        for (let i = 0; i < p.data.length; i += 4)
          if (p.data[i] === 255 && p.data[i + 1] === 0 && p.data[i + 2] === 0) r += 1;
        return r;
      };
      counted += want.n;
      excused += red(diff) - want.n;
      const tiled = PNG.sync.read(PNG.sync.write(diff));
      expect(deviceAntialiased(tiled, a, b, scale)).toBe(want.n);
      expect(PNG.sync.write(tiled).equals(want.diff)).toBe(true);
    }
    // The images exercise both classes: blocks that count and blocks that are anti-aliasing.
    expect(counted).toBeGreaterThan(0);
    expect(excused).toBeGreaterThan(0);
  });
});
