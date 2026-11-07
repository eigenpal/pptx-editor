// Read an image's pixel size and resolution from its header: PNG (`IHDR`, `pHYs`), JPEG (`SOFn`,
// JFIF density), GIF, BMP, and WebP (`VP8 `, `VP8L`, `VP8X`). Every read is bounds-checked and the scan is capped, because the
// bytes come from the file (AGENTS.md "Security"). Nothing is decoded.

/** Pixel size and resolution. `dpiX`/`dpiY` are undefined when the file states none. */
export interface ImageSize {
  readonly width: number;
  readonly height: number;
  readonly dpiX?: number;
  readonly dpiY?: number;
}

/** The largest dimension accepted, in pixels. */
const MAX_PIXELS = 1 << 16;
/** How far into a JPEG the scan for a frame header goes. */
const MAX_JPEG_SCAN = 4 * 1024 * 1024;

const u16be = (b: Uint8Array, i: number): number => (b[i] << 8) | b[i + 1];
const u16le = (b: Uint8Array, i: number): number => b[i] | (b[i + 1] << 8);
const u32be = (b: Uint8Array, i: number): number =>
  ((b[i] << 24) >>> 0) + (b[i + 1] << 16) + (b[i + 2] << 8) + b[i + 3];
const i32le = (b: Uint8Array, i: number): number =>
  b[i] | (b[i + 1] << 8) | (b[i + 2] << 16) | (b[i + 3] << 24);

const valid = (w: number, h: number): boolean =>
  Number.isInteger(w) &&
  Number.isInteger(h) &&
  w > 0 &&
  h > 0 &&
  w <= MAX_PIXELS &&
  h <= MAX_PIXELS;

/** A plausible resolution, or undefined. */
const dpi = (v: number): number | undefined =>
  Number.isFinite(v) && v >= 1 && v <= 10_000 ? v : undefined;

function png(b: Uint8Array): ImageSize | null {
  if (b.length < 33 || u32be(b, 12) !== 0x49484452) return null; // IHDR
  const width = u32be(b, 16);
  const height = u32be(b, 20);
  if (!valid(width, height)) return null;
  // Chunks: length, type, data, CRC. pHYs must come before IDAT.
  let i = 8;
  for (let n = 0; n < 64 && i + 8 <= b.length; n += 1) {
    const length = u32be(b, i);
    const type = u32be(b, i + 4);
    if (type === 0x49444154) break; // IDAT
    if (type === 0x70485973 && length === 9 && i + 17 <= b.length && b[i + 16] === 1) {
      // pHYs in pixels per metre.
      const x = dpi(u32be(b, i + 8) * 0.0254);
      const y = dpi(u32be(b, i + 12) * 0.0254);
      return { width, height, ...(x ? { dpiX: x } : {}), ...(y ? { dpiY: y } : {}) };
    }
    if (length > b.length) break;
    i += 12 + length;
  }
  return { width, height };
}

function jpeg(b: Uint8Array): ImageSize | null {
  let i = 2;
  let density: { x?: number; y?: number } = {};
  const end = Math.min(b.length, MAX_JPEG_SCAN);
  while (i + 4 <= end) {
    if (b[i] !== 0xff) return null;
    const marker = b[i + 1];
    if (marker === 0xff) {
      i += 1;
      continue;
    }
    const length = u16be(b, i + 2);
    if (length < 2) return null;
    if (marker === 0xe0 && i + 18 <= end && b[i + 4] === 0x4a && b[i + 5] === 0x46) {
      // JFIF: units 1 = dots per inch, 2 = dots per centimetre.
      const units = b[i + 11];
      const f = units === 1 ? 1 : units === 2 ? 2.54 : 0;
      if (f) density = { x: dpi(u16be(b, i + 12) * f), y: dpi(u16be(b, i + 14) * f) };
    }
    const isFrame =
      marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
    if (isFrame) {
      if (i + 9 > end) return null;
      const height = u16be(b, i + 5);
      const width = u16be(b, i + 7);
      if (!valid(width, height)) return null;
      return {
        width,
        height,
        ...(density.x ? { dpiX: density.x } : {}),
        ...(density.y ? { dpiY: density.y } : {}),
      };
    }
    i += 2 + length;
  }
  return null;
}

function gif(b: Uint8Array): ImageSize | null {
  if (b.length < 10) return null;
  const width = u16le(b, 6);
  const height = u16le(b, 8);
  return valid(width, height) ? { width, height } : null;
}

function bmp(b: Uint8Array): ImageSize | null {
  if (b.length < 46) return null;
  const width = i32le(b, 18);
  const height = Math.abs(i32le(b, 22));
  if (!valid(width, height)) return null;
  const x = dpi(i32le(b, 38) * 0.0254);
  const y = dpi(i32le(b, 42) * 0.0254);
  return { width, height, ...(x ? { dpiX: x } : {}), ...(y ? { dpiY: y } : {}) };
}

const u24le = (b: Uint8Array, i: number): number => b[i] | (b[i + 1] << 8) | (b[i + 2] << 16);
const ascii = (b: Uint8Array, i: number, text: string): boolean =>
  [...text].every((c, k) => b[i + k] === c.charCodeAt(0));

/** A WebP: the first chunk after the RIFF header gives the size. WebP states no resolution. */
function webp(b: Uint8Array): ImageSize | null {
  if (b.length < 30) return null;
  let width = 0;
  let height = 0;
  if (ascii(b, 12, 'VP8 ')) {
    // A key frame: a 3-byte frame tag, the start code 9D 01 2A, then 14-bit sizes.
    if (b[23] !== 0x9d || b[24] !== 0x01 || b[25] !== 0x2a) return null;
    width = u16le(b, 26) & 0x3fff;
    height = u16le(b, 28) & 0x3fff;
  } else if (ascii(b, 12, 'VP8L')) {
    // The signature 2F, then the width and height less one in 14 bits each.
    if (b[20] !== 0x2f) return null;
    width = 1 + (b[21] | ((b[22] & 0x3f) << 8));
    height = 1 + ((b[22] >> 6) | (b[23] << 2) | ((b[24] & 0x0f) << 10));
  } else if (ascii(b, 12, 'VP8X')) {
    // The canvas width and height less one in 24 bits each.
    width = 1 + u24le(b, 24);
    height = 1 + u24le(b, 27);
  } else return null;
  return valid(width, height) ? { width, height } : null;
}

/** The size of a PNG, JPEG, GIF, BMP, or WebP image, or null for anything else. */
export function imageSize(bytes: Uint8Array): ImageSize | null {
  const b = bytes;
  if (b.length >= 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47)
    return png(b);
  if (b.length >= 4 && b[0] === 0xff && b[1] === 0xd8) return jpeg(b);
  if (b.length >= 6 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return gif(b);
  if (b.length >= 2 && b[0] === 0x42 && b[1] === 0x4d) return bmp(b);
  if (b.length >= 12 && ascii(b, 0, 'RIFF') && ascii(b, 8, 'WEBP')) return webp(b);
  return null;
}

/** The size in EMU at the image's resolution (96 DPI when it states none). */
export function imageSizeEmu(size: ImageSize): { cx: number; cy: number } {
  return {
    cx: Math.round((size.width * 914_400) / (size.dpiX ?? 96)),
    cy: Math.round((size.height * 914_400) / (size.dpiY ?? 96)),
  };
}
