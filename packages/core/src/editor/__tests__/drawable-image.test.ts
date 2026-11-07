// Picture bytes reach the browser only when their signature and header give a picture within
// the limits; anything else paints the placeholder.

import { expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { drawableImageMime } from '../drawable-image.ts';
import { paintLayout } from '../export-image.ts';
import { createPptxEditor } from '../index.ts';

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

test('an export paints a damaged picture as the placeholder', async () => {
  const IN = 914_400;
  const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
  const pic = (id: number, rId: string) =>
    `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="P${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr><a:xfrm><a:off x="${id * IN}" y="${IN}"/><a:ext cx="${IN}" cy="${IN}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
  const deck = buildPptx({
    slides: [
      {
        layout: 'blank',
        rels: [
          ['rIdGood', R_IMAGE, '../media/good.png'],
          ['rIdBad', R_IMAGE, '../media/bad.png'],
        ],
        shapes: pic(2, 'rIdGood') + pic(4, 'rIdBad'),
      },
    ],
    media: [
      { name: 'good.png', bytes: png(10, 10), contentType: 'image/png' },
      { name: 'bad.png', bytes: text('<svg onload="x"/>'), contentType: 'image/png' },
    ],
  });
  const win = new GlobalWindow();
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(deck)).ok).toBe(true);
  const session = editor.session!;
  const doc = win.document as unknown as Document;
  const svg = paintLayout(session, session.layout(0)!, 100, 100, () => false, undefined, doc);
  const hrefs = [...svg.querySelectorAll('image')].map((i) => i.getAttribute('href') ?? '');
  expect(hrefs.filter((h) => h.startsWith('data:image/png'))).toHaveLength(1);
  expect(hrefs.some((h) => h.includes(btoa('<svg')))).toBe(false);
  editor.destroy();
  await win.happyDOM.close();
});
