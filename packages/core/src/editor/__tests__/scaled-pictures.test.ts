// The raster pictures of one part draw within a decoded pixel budget on screen: over it, each
// one draws scaled down off the main thread, and the placeholder paints until it is ready.

import { afterEach, expect, test } from 'bun:test';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '../../store/index.ts';
import { createPptxEditor } from '../index.ts';
import type { DrawablePicture } from '../metafile-pictures.ts';
import {
  PART_PIXEL_BUDGET,
  partPictureScale,
  ScaledPictures,
  withStoredOrientation,
} from '../scaled-pictures.ts';

const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const SLIDE = '/ppt/slides/slide1.xml';
const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));

/** A PNG header with the given size: the signature and an IHDR chunk. */
function png(width: number, height: number, seed = 0): Uint8Array {
  const b = new Uint8Array(34);
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52]);
  new DataView(b.buffer).setUint32(16, width);
  new DataView(b.buffer).setUint32(20, height);
  b[24] = 8;
  b[25] = 6;
  b[33] = seed;
  return b;
}

const pic = (id: number, rId: string): string =>
  `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="P${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
  `<p:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';

/** Slide 1 shows `count` distinct 8000x8000 pictures (64 M pixels each); slide 2 shows one. */
function deck(count: number): Uint8Array {
  const ids = Array.from({ length: count }, (_, i) => i);
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        rels: [
          ...ids.map((i): [string, string, string] => [`rIdP${i}`, R_IMAGE, `../media/p${i}.png`]),
          // A second relationship to the same part counts once.
          ['rIdAgain', R_IMAGE, '../media/p0.png'],
        ],
        shapes: ids.map((i) => pic(i + 2, `rIdP${i}`)).join(''),
      },
      {
        layout: 'blank',
        rels: [['rIdOne', R_IMAGE, '../media/p0.png']],
        shapes: pic(2, 'rIdOne'),
      },
    ],
    media: ids.map((i) => ({
      name: `p${i}.png`,
      bytes: png(8000, 8000, i),
      contentType: 'image/png',
    })),
  });
}

test('a part over the pixel budget scales its pictures to fit it', () => {
  const read = readPptx(deck(3));
  if (!read.ok) throw new Error('deck did not read');
  const scale = partPictureScale(read.pkg, SLIDE);
  expect(scale).toBeCloseTo(Math.sqrt(PART_PIXEL_BUDGET / (3 * 64_000_000)), 6);
  // One picture of 64 M pixels is within the budget.
  expect(partPictureScale(read.pkg, '/ppt/slides/slide2.xml')).toBe(1);
  // A part with no pictures.
  expect(partPictureScale(read.pkg, '/ppt/presentation.xml')).toBe(1);
});

/** A JPEG start with an EXIF block whose orientation is `value`. */
function exifJpeg(value: number, little: boolean): Uint8Array {
  const u16 = (n: number) => (little ? [n & 255, n >> 8] : [n >> 8, n & 255]);
  const u32 = (n: number) => (little ? [n, 0, 0, 0] : [0, 0, 0, n]);
  const tiff = [
    ...(little ? [0x49, 0x49] : [0x4d, 0x4d]),
    ...u16(42),
    ...u32(8),
    ...u16(1),
    ...u16(0x0112),
    ...u16(3),
    ...u32(1),
    ...u16(value),
    0,
    0,
    ...u32(0),
  ];
  const payload = [0x45, 0x78, 0x69, 0x66, 0, 0, ...tiff];
  const length = payload.length + 2;
  return new Uint8Array([
    0xff,
    0xd8,
    0xff,
    0xe1,
    length >> 8,
    length & 255,
    ...payload,
    0xff,
    0xda,
  ]);
}

test('a JPEG decodes to its stored pixels: its EXIF orientation becomes 1 in a copy', () => {
  for (const little of [false, true]) {
    const turned = exifJpeg(6, little);
    const stored = withStoredOrientation(turned);
    expect(stored).not.toBe(turned);
    expect(stored).toEqual(exifJpeg(1, little));
    // The part's bytes stay as they are.
    expect(turned).toEqual(exifJpeg(6, little));
    const upright = exifJpeg(1, little);
    expect(withStoredOrientation(upright)).toBe(upright);
  }
  const other = png(10, 10);
  expect(withStoredOrientation(other)).toBe(other);
  // A cut EXIF block is read no further than its end.
  const cut = exifJpeg(6, false).slice(0, 20);
  expect(withStoredOrientation(cut)).toBe(cut);
});

test('a scaled picture paints the placeholder, then its scaled bytes after one repaint', async () => {
  const calls: number[][] = [];
  let finish = (_: Uint8Array | null) => {};
  let changed = 0;
  const scaled = new ScaledPictures(
    () => (changed += 1),
    (_bytes, _mime, w, h) => {
      calls.push([w, h]);
      return new Promise((r) => (finish = r));
    },
    () => true
  );
  const picture: DrawablePicture = { bytes: png(8000, 4000), mime: 'image/png' };
  expect(scaled.shown(picture, 0.5, SLIDE)).toBeUndefined();
  // A second paint, also from another part, waits for the same work.
  expect(scaled.shown(picture, 0.5, '/ppt/slideLayouts/slideLayout1.xml')).toBeUndefined();
  expect(calls).toEqual([[4000, 2000]]);
  expect(scaled.busy).toBe(true);
  expect(scaled.serialOf(SLIDE)).toBe('');
  const out = new Uint8Array([1, 2, 3]);
  finish(out);
  await wait(10);
  expect(changed).toBe(1);
  expect(scaled.busy).toBe(false);
  expect(scaled.serialOf(SLIDE)).not.toBe('');
  expect(scaled.serialOf('/ppt/slideLayouts/slideLayout1.xml')).not.toBe('');
  expect(scaled.shown(picture, 0.5, SLIDE)).toEqual({
    picture: { bytes: out, mime: 'image/png' },
    key: '4000x2000',
  });
  // Within the budget, an SVG picture, and a runtime that cannot scale: the stored bytes.
  expect(scaled.shown(picture, 1, SLIDE)).toEqual({ picture, key: '' });
  const svg = { bytes: new TextEncoder().encode('<svg/>'), mime: 'image/svg+xml' };
  expect(scaled.shown(svg, 0.5, SLIDE)).toEqual({ picture: svg, key: '' });
  const plain = new ScaledPictures(
    () => {},
    undefined,
    () => false
  );
  expect(plain.shown(picture, 0.5, SLIDE)).toEqual({ picture, key: '' });
});

test('a picture that does not scale paints the placeholder and is not tried again', async () => {
  let calls = 0;
  const scaled = new ScaledPictures(
    () => {},
    async () => {
      calls += 1;
      throw new Error('cannot decode');
    },
    () => true
  );
  const picture: DrawablePicture = { bytes: png(8000, 8000), mime: 'image/png' };
  expect(scaled.shown(picture, 0.25, SLIDE)).toBeUndefined();
  await wait(10);
  expect(scaled.shown(picture, 0.25, SLIDE)).toBeUndefined();
  expect(calls).toBe(1);
});

test('a disposed cache starts no work and stays quiet', async () => {
  let changed = 0;
  let finish = (_: Uint8Array | null) => {};
  const scaled = new ScaledPictures(
    () => (changed += 1),
    () => new Promise((r) => (finish = r)),
    () => true
  );
  const picture: DrawablePicture = { bytes: png(8000, 8000), mime: 'image/png' };
  scaled.shown(picture, 0.5, SLIDE);
  scaled.dispose();
  finish(new Uint8Array([1]));
  await wait(10);
  expect(changed).toBe(0);
  expect(scaled.shown(picture, 0.5, SLIDE)).toBeUndefined();
});

const globals = globalThis as unknown as Record<string, unknown>;
const saved = {
  createImageBitmap: globals.createImageBitmap,
  OffscreenCanvas: globals.OffscreenCanvas,
};
afterEach(() => Object.assign(globals, saved));

test('the editor draws a slide over the budget scaled, and exports the stored pictures', async () => {
  // A browser that scales: each picture becomes the bytes of its new size.
  const sizes: string[] = [];
  globals.createImageBitmap = async (
    _blob: Blob,
    o: { resizeWidth: number; resizeHeight: number }
  ) => {
    sizes.push(`${o.resizeWidth}x${o.resizeHeight}`);
    return { width: o.resizeWidth, height: o.resizeHeight, close() {} };
  };
  globals.OffscreenCanvas = class {
    getContext() {
      return { drawImage() {} };
    }
    async convertToBlob() {
      return new Blob([png(10, 10, 99).slice().buffer], { type: 'image/png' });
    }
  };
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(deck(3))).ok).toBe(true);
  const session = editor.session!;
  let repaints = 0;
  editor.subscribe(() => (repaints += 1));
  const keys = [SLIDE, '/ppt/slides/slide2.xml'].map((p) => session.paintKey(p));
  // Slide 2 is within the budget: its picture draws at once, as stored.
  expect(session.imageUrl('/ppt/slides/slide2.xml', 'rIdOne')).toStartWith('blob:');
  for (const rId of ['rIdP0', 'rIdP1', 'rIdP2'])
    expect(session.imageUrl(SLIDE, rId)).toBeUndefined();
  await wait(50);
  expect(repaints).toBeGreaterThan(0);
  const side = Math.floor(8000 * Math.sqrt(PART_PIXEL_BUDGET / (3 * 64_000_000)));
  expect(sizes).toEqual([`${side}x${side}`, `${side}x${side}`, `${side}x${side}`]);
  for (const rId of ['rIdP0', 'rIdP1', 'rIdP2'])
    expect(session.imageUrl(SLIDE, rId)).toStartWith('blob:');
  // The slide repaints; the other slide keeps its key.
  expect(session.paintKey(SLIDE)).not.toBe(keys[0]);
  expect(session.paintKey('/ppt/slides/slide2.xml')).toBe(keys[1]!);
  // The same part keeps one URL at each size.
  expect(session.imageUrl(SLIDE, 'rIdP0')).not.toBe(
    session.imageUrl('/ppt/slides/slide2.xml', 'rIdOne')
  );
  // Export and copy read the stored bytes.
  expect(session.picture('/ppt/media/p0.png')?.bytes.length).toBe(34);
  editor.destroy();
});

test('a part keeps its factor while its pictures change a little', () => {
  const scaled = new ScaledPictures(() => {});
  expect(scaled.steady(SLIDE, 0.5)).toBe(0.5);
  // Within 25% of the factor before: the pictures keep their size.
  expect(scaled.steady(SLIDE, 0.45)).toBe(0.5);
  expect(scaled.steady(SLIDE, 0.6)).toBe(0.5);
  // Further: they scale again, to the new factor.
  expect(scaled.steady(SLIDE, 0.39)).toBe(0.39);
  // Within the budget: the stored size, and no factor to keep.
  expect(scaled.steady(SLIDE, 1)).toBe(1);
  expect(scaled.steady(SLIDE, 0.9)).toBe(0.9);
  expect(scaled.steady('/ppt/slides/slide2.xml', 0.45)).toBe(0.45);
});

test('a picture draws at its last scaled size while a new size scales', async () => {
  const finish: ((b: Uint8Array | null) => void)[] = [];
  const scaled = new ScaledPictures(
    () => {},
    () => new Promise((r) => finish.push(r)),
    () => true
  );
  const picture: DrawablePicture = { bytes: png(8000, 8000), mime: 'image/png' };
  scaled.shown(picture, 0.5, SLIDE);
  const first = new Uint8Array([1]);
  finish[0]!(first);
  await wait(10);
  expect(scaled.shown(picture, 0.25, SLIDE)).toEqual({
    picture: { bytes: first, mime: 'image/png' },
    key: '4000x4000',
  });
  const second = new Uint8Array([2]);
  finish[1]!(second);
  await wait(10);
  expect(scaled.shown(picture, 0.25, SLIDE)?.picture.bytes).toBe(second);
  // The size before stays ready.
  expect(scaled.shown(picture, 0.5, SLIDE)?.picture.bytes).toBe(first);
});

test('pictures keep their URLs when a picture is inserted, and old sizes are revoked', async () => {
  globals.createImageBitmap = async () => ({ close() {} });
  let serial = 0;
  globals.OffscreenCanvas = class {
    getContext() {
      return { drawImage() {} };
    }
    async convertToBlob() {
      return new Blob([png(10, 10, (serial += 1)).slice().buffer], { type: 'image/png' });
    }
  };
  const revoked: string[] = [];
  const revoke = URL.revokeObjectURL;
  URL.revokeObjectURL = (url: string) => (revoked.push(url), revoke(url));
  try {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(deck(3))).ok).toBe(true);
    const session = editor.session!;
    const ids = ['rIdP0', 'rIdP1', 'rIdP2'];
    const urls = () => ids.map((rId) => session.imageUrl(SLIDE, rId));
    urls();
    await wait(50);
    const before = urls();
    expect(before.every((u) => u?.startsWith('blob:'))).toBe(true);
    const insert = (seed: number) =>
      editor.exec({
        type: 'insertImage',
        bytes: png(8000, 8000, seed),
        mime: 'image/png',
        width: 8000,
        height: 8000,
      });
    // One more picture moves the factor less than 25%: nothing scales again.
    expect(insert(50).ok).toBe(true);
    expect(urls()).toEqual(before);
    await wait(50);
    expect(urls()).toEqual(before);
    expect(revoked).toEqual([]);
    // One more moves it further: each picture scales again, and draws its old URL meanwhile.
    expect(insert(51).ok).toBe(true);
    expect(urls()).toEqual(before);
    await wait(50);
    const after = urls();
    expect(after.every((u, i) => u?.startsWith('blob:') && u !== before[i])).toBe(true);
    for (const url of before) expect(revoked).toContain(url!);
    editor.destroy();
  } finally {
    URL.revokeObjectURL = revoke;
  }
});
