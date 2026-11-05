// Supersampled comparison rasters. Chromium and Poppler snap each glyph to the device pixel grid
// in different ways: Chromium rounds a glyph's y to a whole pixel, and Poppler's Splash
// rasterizer floors it. At the comparison size, that puts an exact text line 1 px lower in our
// image than in the reference image whenever the PDF baseline has a pixel fraction of 0.5 or
// more. So both images are rasterized at `THRESHOLDS.rasterScale` device pixels per comparison
// pixel, and each is box-filtered down to the comparison size with `downsample`. The snapping
// then moves a glyph by at most 1/scale of a comparison pixel, and both images go through the
// same filter (`CONTRIBUTING.md`, "Fidelity harness").
//
// The filter turns every edge into a ramp, and pixelmatch's anti-aliasing test calls a ramp
// anti-aliasing wherever it lies. Judged on the filtered images, a stroke moved by 1 px would
// not count. So anti-aliasing is judged on the device-scale images (`deviceAntialiased`), where
// the two rasterizers' ramps are 1 device pixel wide again.

import pixelmatch from 'pixelmatch';
import { PNG } from 'pngjs';
import { THRESHOLDS } from './thresholds.ts';
import type { SizePx } from './references.ts';

/**
 * The comparison size of a raster that is `scale` times larger: each side `ceil(side / scale)`.
 * A raster that is short where its page ends (`expectedRasterSizePx`) stays short by the same
 * whole pixels, because `ceil(ceil(x) / n) = ceil(x / n)` for a whole number `n`.
 */
export const downsampledSize = (size: SizePx, scale: number): SizePx => ({
  width: Math.ceil(size.width / scale),
  height: Math.ceil(size.height / scale),
});

/** RGBA pixels: a decoded PNG, or a raster that was never a PNG (`references.ts`). */
export interface Raster {
  readonly width: number;
  readonly height: number;
  readonly data: Uint8Array;
}

/** The pixels of a PNG file's bytes, or the raster itself. */
export const pixels = (image: Buffer | Raster): Raster =>
  Buffer.isBuffer(image) ? PNG.sync.read(image) : image;

/**
 * Box-filter an RGBA image down by a whole `scale`: each output pixel is the mean of its
 * `scale × scale` block, with color weighted by alpha. A block at the right or bottom edge that
 * the image does not fill averages the pixels it has. The mean is taken on the stored (sRGB)
 * values, as both rasterizers blend their own anti-aliased edges. Scale 1 returns the image.
 */
export function downsample(image: Raster, scale: number): PNG {
  if (!Number.isInteger(scale) || scale < 1)
    throw new RangeError(`the raster scale must be a whole number of 1 or more, not ${scale}`);
  if (scale === 1) {
    if (image instanceof PNG) return image;
    const copy = new PNG({ width: image.width, height: image.height });
    copy.data.set(image.data);
    return copy;
  }
  const { width: w, height: h } = downsampledSize(image, scale);
  const out = new PNG({ width: w, height: h });
  const src = image.data;
  const dst = out.data;
  const width = image.width;
  // One row of block sums at a time: premultiplied red, green, blue, and alpha. Every sum is a
  // whole number below 2^53, so the order of the additions does not change it.
  const sums = new Float64Array(w * 4);
  for (let by = 0; by < h; by += 1) {
    sums.fill(0);
    const y0 = by * scale;
    const y1 = Math.min(image.height, y0 + scale);
    for (let y = y0; y < y1; y += 1) {
      const row = y * width * 4;
      for (let bx = 0, x = 0; bx < w; bx += 1) {
        const x1 = Math.min(width, x + scale);
        let r = 0;
        let g = 0;
        let b = 0;
        let a = 0;
        for (let i = row + x * 4; x < x1; x += 1, i += 4) {
          const al = src[i + 3];
          r += src[i] * al;
          g += src[i + 1] * al;
          b += src[i + 2] * al;
          a += al;
        }
        const s = bx * 4;
        sums[s] += r;
        sums[s + 1] += g;
        sums[s + 2] += b;
        sums[s + 3] += a;
      }
    }
    const rows = y1 - y0;
    for (let bx = 0; bx < w; bx += 1) {
      const s = bx * 4;
      const o = (by * w + bx) * 4;
      const count = rows * (Math.min(width, (bx + 1) * scale) - bx * scale);
      const alpha = sums[s + 3];
      if (alpha > 0) {
        dst[o] = Math.round(sums[s] / alpha);
        dst[o + 1] = Math.round(sums[s + 1] / alpha);
        dst[o + 2] = Math.round(sums[s + 2] / alpha);
      }
      dst[o + 3] = Math.round(alpha / count);
    }
  }
  return out;
}

/** `downsample` on a PNG file's bytes, returning PNG bytes. */
export function downsamplePng(png: Buffer, scale: number): Buffer {
  if (scale === 1) return png;
  return PNG.sync.write(downsample(PNG.sync.read(png), scale));
}

/** Pixelmatch's output colors, set apart so the classes can be read back. */
const REAL: readonly [number, number, number] = [255, 0, 0];
const AA: readonly [number, number, number] = [255, 255, 0];
const is = (d: Uint8Array, i: number, c: readonly number[]): boolean =>
  d[i] === c[0] && d[i + 1] === c[1] && d[i + 2] === c[2];

/** The overlap of an image with `w` × `h` px, as RGBA bytes. */
function cropped(p: Raster, w: number, h: number): Uint8Array {
  if (p.width === w && p.height === h) return p.data;
  const out = new Uint8Array(w * h * 4);
  for (let y = 0; y < h; y += 1)
    out.set(p.data.subarray(y * p.width * 4, y * p.width * 4 + w * 4), y * w * 4);
  return out;
}

/**
 * The differing comparison pixels that count, judged at the device scale. `diff` is pixelmatch's
 * output on the filtered images with every difference marked pure red (`includeAA`). The two
 * device images are compared with pixelmatch's own anti-aliasing test. A red pixel counts when
 * its `scale × scale` block of device pixels has more real differences than anti-aliased ones;
 * otherwise it is repainted as anti-aliasing (yellow), so the overlay and its box show only what
 * counts. A block with no device difference at all is the sum of differences below the threshold,
 * and does not count. Returns the number of pixels that count.
 *
 * Only the blocks under red pixels are read, so pixelmatch skips every other device pixel
 * (`ignoreMask`). Its anti-aliasing test reads the neighbors from the images, not the mask, so
 * each block gets the same classes as from a full comparison.
 */
export function deviceAntialiased(
  diff: PNG,
  render: Buffer | Raster,
  reference: Buffer | Raster,
  scale: number
): number {
  const a = pixels(render);
  const b = pixels(reference);
  const w = Math.min(a.width, b.width);
  const h = Math.min(a.height, b.height);
  const red: number[] = [];
  for (let y = 0, i = 0; y < diff.height; y += 1)
    for (let x = 0; x < diff.width; x += 1, i += 4) if (is(diff.data, i, REAL)) red.push(x, y);
  const ignore = new Uint8Array(w * h).fill(1);
  for (let k = 0; k < red.length; k += 2) {
    const x0 = red[k] * scale;
    const x1 = Math.min(w, x0 + scale);
    if (x0 >= x1) continue;
    for (let yy = red[k + 1] * scale; yy < Math.min(h, (red[k + 1] + 1) * scale); yy += 1)
      ignore.fill(0, yy * w + x0, yy * w + x1);
  }
  const out = new Uint8Array(w * h * 4);
  pixelmatch(cropped(a, w, h), cropped(b, w, h), out, w, h, {
    threshold: THRESHOLDS.pixelThreshold,
    includeAA: false,
    alpha: 0,
    diffColor: [...REAL] as [number, number, number],
    diffColorAlt: [...REAL] as [number, number, number],
    aaColor: [...AA] as [number, number, number],
    ignoreMask: ignore,
  });
  let n = 0;
  for (let k = 0; k < red.length; k += 2) {
    const x = red[k];
    const y = red[k + 1];
    let real = 0;
    let aa = 0;
    for (let yy = y * scale; yy < Math.min(h, (y + 1) * scale); yy += 1)
      for (let xx = x * scale; xx < Math.min(w, (x + 1) * scale); xx += 1) {
        const j = (yy * w + xx) * 4;
        if (is(out, j, REAL)) real += 1;
        else if (is(out, j, AA)) aa += 1;
      }
    if (real > aa) n += 1;
    else diff.data.set(AA, (y * diff.width + x) * 4);
  }
  return n;
}
