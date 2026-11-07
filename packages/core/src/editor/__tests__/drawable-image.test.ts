// Picture bytes reach the browser only when their signature and header give a picture within
// the limits; anything else paints the placeholder.

import { expect, test } from 'bun:test';
import { drawableImageMime } from '../drawable-image.ts';

/** A PNG header with the given size: the signature and an IHDR chunk. */
function png(width: number, height: number): Uint8Array {
  const b = new Uint8Array(33);
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52]);
  new DataView(b.buffer).setUint32(16, width);
  new DataView(b.buffer).setUint32(20, height);
  b[24] = 8;
  b[25] = 6;
  return b;
}

const text = (s: string) => new TextEncoder().encode(s);

test('the signature decides the type, not the part name', () => {
  expect(drawableImageMime(png(10, 10), 'image/png')).toBe('image/png');
  expect(drawableImageMime(png(10, 10), undefined)).toBe('image/png');
  expect(drawableImageMime(text('RIFF\0\0\0\0WEBPVP8 '), 'image/png')).toBe('image/webp');
});

test('empty, unknown, and markup bytes in a raster part draw nothing', () => {
  expect(drawableImageMime(new Uint8Array(), 'image/png')).toBeUndefined();
  expect(drawableImageMime(text('\x01\0\0\0 EMF'), 'image/png')).toBeUndefined();
  expect(
    drawableImageMime(text('<svg xmlns="http://www.w3.org/2000/svg"/>'), 'image/png')
  ).toBeUndefined();
});

test('an SVG part must start as markup', () => {
  const bom = String.fromCharCode(0xfeff);
  expect(drawableImageMime(text(`${bom}  <svg/>`), 'image/svg+xml')).toBe('image/svg+xml');
  expect(drawableImageMime(png(10, 10), 'image/svg+xml')).toBeUndefined();
});

test('a header beyond the size limits draws nothing', () => {
  expect(drawableImageMime(png(0, 10), 'image/png')).toBeUndefined();
  expect(drawableImageMime(png(60_000, 60_000), 'image/png')).toBeUndefined();
  expect(drawableImageMime(png(20_000, 20_000), 'image/png')).toBeUndefined();
});
