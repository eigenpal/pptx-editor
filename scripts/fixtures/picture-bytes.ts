// Small picture files for tests and fixtures: an EMF and a WMF that draw a filled rectangle, and
// WebP headers. The bytes follow the EMF and WMF record layouts and the WebP container layout.

/** An EMF record: type, size in bytes, then 32-bit words. */
const record = (type: number, words: readonly number[]): number[] => [
  type,
  8 + words.length * 4,
  ...words,
];

/**
 * An EMF of `width` × `height` pixels at 96 DPI that fills a rectangle inset by a tenth of the
 * size in `color` (`0xRRGGBB`). `frame` overrides the picture frame in 0.01 mm.
 */
export function emfRectangle(
  width = 200,
  height = 100,
  color = 0xff0000,
  frame?: { readonly right: number; readonly bottom: number }
): Uint8Array {
  // A COLORREF is 0x00BBGGRR.
  const colorRef = ((color & 0xff) << 16) | (color & 0xff00) | ((color >> 16) & 0xff);
  const dx = Math.round(width / 10);
  const dy = Math.round(height / 10);
  const records = [
    record(39, [1, 0, colorRef, 0]), // EMR_CREATEBRUSHINDIRECT: handle 1, solid
    record(37, [1]), // EMR_SELECTOBJECT
    record(43, [dx, dy, width - dx, height - dy]), // EMR_RECTANGLE
    record(14, [0, 16, 20]), // EMR_EOF
  ];
  const body = records.flat();
  const right = frame?.right ?? Math.round((width * 2540) / 96);
  const bottom = frame?.bottom ?? Math.round((height * 2540) / 96);
  // EMR_HEADER (88 bytes): bounds, frame, ` EMF`, version, size, records, handles, reserved,
  // description, palette, device size in pixels and in millimetres.
  const header = [1, 88, 0, 0, width - 1, height - 1, 0, 0, right, bottom, 0x464d4520, 0x10000];
  header.push((22 + body.length) * 4, records.length + 1, 2, 0, 0, 0, 1920, 1080, 508, 286);
  return new Uint8Array(new Uint32Array([...header, ...body]).buffer);
}

/** A placeable WMF of `width` × `height` units at `inch` units per inch that fills a rectangle. */
export function wmfRectangle(width = 1440, height = 720, inch = 1440): Uint8Array {
  const words: number[] = [];
  const u32 = (v: number) => words.push(v & 0xffff, (v >>> 16) & 0xffff);
  const wmfRecord = (fn: number, params: readonly number[]) => {
    u32(3 + params.length);
    words.push(fn, ...params);
  };
  // The placeable header: key, handle, bounding box, units per inch, reserved, checksum.
  u32(0x9ac6cdd7);
  words.push(0, 0, 0, width, height, inch, 0, 0);
  let sum = 0;
  for (const w of words.slice(0, 10)) sum ^= w;
  words.push(sum);
  const start = words.length;
  // The standard header: disk metafile, 9 words, version 3, size, 1 object, max record, 0.
  words.push(2, 9, 0x300, 0, 0, 1, 0, 0, 0);
  wmfRecord(0x020b, [0, 0]); // META_SETWINDOWORG
  wmfRecord(0x020c, [height, width]); // META_SETWINDOWEXT
  wmfRecord(0x02fc, [0, 0xff, 0x00, 0]); // META_CREATEBRUSHINDIRECT: solid, red
  wmfRecord(0x012d, [0]); // META_SELECTOBJECT
  wmfRecord(0x041b, [height - 72, width - 72, 72, 72]); // META_RECTANGLE: bottom, right, top, left
  wmfRecord(0x0000, []); // META_EOF
  const size = words.length - start;
  words[start + 3] = size & 0xffff;
  words[start + 4] = size >>> 16;
  words[start + 6] = 7;
  return new Uint8Array(new Uint16Array(words).buffer);
}

/** A WebP header of the given chunk kind and size: enough for a size read, not a decode. */
export function webpHeader(kind: 'VP8 ' | 'VP8L' | 'VP8X', width: number, height: number): Uint8Array {
  const b = new Uint8Array(32);
  b.set([...'RIFF'].map((c) => c.charCodeAt(0)), 0);
  new DataView(b.buffer).setUint32(4, 24, true);
  b.set([...`WEBP${kind}`].map((c) => c.charCodeAt(0)), 8);
  const view = new DataView(b.buffer);
  if (kind === 'VP8 ') {
    b.set([0x9d, 0x01, 0x2a], 23);
    view.setUint16(26, width & 0x3fff, true);
    view.setUint16(28, height & 0x3fff, true);
  } else if (kind === 'VP8L') {
    const w = width - 1;
    const h = height - 1;
    b[20] = 0x2f;
    b[21] = w & 0xff;
    b[22] = ((w >> 8) & 0x3f) | ((h & 0x03) << 6);
    b[23] = (h >> 2) & 0xff;
    b[24] = (h >> 10) & 0x0f;
  } else {
    for (const [at, v] of [
      [24, width - 1],
      [27, height - 1],
    ] as const)
      b.set([v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff], at);
  }
  return b;
}
