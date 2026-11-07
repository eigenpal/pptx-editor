// Picture bytes reach the browser only when their signature and header give a picture within
// the limits; anything else paints the placeholder.

import { expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { webpHeader } from '../../../../../scripts/fixtures/picture-bytes.ts';
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
  expect(drawableImageMime(webpHeader('VP8L', 10, 10), 'image/png')).toBe('image/webp');
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

test('a WebP passes the same size checks', () => {
  expect(drawableImageMime(webpHeader('VP8 ', 640, 480), 'image/webp')).toBe('image/webp');
  expect(drawableImageMime(webpHeader('VP8X', 20_000, 20_000), 'image/webp')).toBeUndefined();
  expect(drawableImageMime(webpHeader('VP8X', 40_000, 10), 'image/webp')).toBeUndefined();
  expect(drawableImageMime(text('RIFF\0\0\0\0WEBPVP8 '), 'image/webp')).toBeUndefined();
});

const svgDoc = (body: string) =>
  text(
    '<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink">' +
      `${body}</svg>`
  );

test('an SVG picture whose <use> links expand past the limit draws nothing', () => {
  // Six levels of ten copies: 10^6 drawn rectangles from 60 elements.
  let defs = '<rect id="l0" width="1" height="1"/>';
  for (let i = 1; i <= 6; i += 1)
    defs += `<g id="l${i}">${`<use xlink:href="#l${i - 1}"/>`.repeat(10)}</g>`;
  const bomb = svgDoc(`<defs>${defs}</defs><use href="#l6"/>`);
  expect(drawableImageMime(bomb, 'image/svg+xml')).toBeUndefined();
  // Four levels draw 10^4 rectangles: within the limit.
  const icon = svgDoc(`<defs>${defs}</defs><use href="#l4"/>`);
  expect(drawableImageMime(icon, 'image/svg+xml')).toBe('image/svg+xml');
});

test('an SVG picture with a <use> cycle, an entity, or deep nesting draws nothing', () => {
  const cycle = svgDoc('<g id="a"><rect width="1" height="1"/><use href="#a"/></g>');
  expect(drawableImageMime(cycle, 'image/svg+xml')).toBeUndefined();
  const entity = text('<!DOCTYPE svg [<!ENTITY a "aaaa">]><svg><text>&a;</text></svg>');
  expect(drawableImageMime(entity, 'image/svg+xml')).toBeUndefined();
  const deep = svgDoc(`${'<g>'.repeat(1000)}${'</g>'.repeat(1000)}`);
  expect(drawableImageMime(deep, 'image/svg+xml')).toBeUndefined();
});

const svg = (bytes: Uint8Array) => drawableImageMime(bytes, 'image/svg+xml');

test('an SVG picture with a nested SVG document in a data: URI draws nothing', () => {
  const inner = btoa('<svg xmlns="http://www.w3.org/2000/svg"><rect/></svg>');
  const uri = `data:image/svg+xml;base64,${inner}`;
  expect(svg(svgDoc(`<image href="${uri}"/>`))).toBeUndefined();
  expect(svg(svgDoc(`<image xlink:href=" DATA:Image/SVG+XML,%3Csvg/%3E"/>`))).toBeUndefined();
  expect(svg(svgDoc(`<filter id="f"><feImage href="${uri}"/></filter>`))).toBeUndefined();
  expect(svg(svgDoc(`<rect style="fill: url('${uri}')"/>`))).toBeUndefined();
  expect(svg(svgDoc(`<style>rect { fill: url("${uri}") }</style><rect/>`))).toBeUndefined();
  // A CSS escape spells the same URI.
  expect(svg(svgDoc(`<style>rect { fill: url(\\64 ata:image/svg+xml,x) }</style>`))).toBe(
    undefined
  );
  // A raster picture in a data: URI draws.
  const png = 'data:image/png;base64,iVBORw0KGgo=';
  expect(svg(svgDoc(`<image href="${png}" width="1" height="1"/>`))).toBe('image/svg+xml');
});

/** A path of `n` vertices. */
const path = (n: number, attrs = '') =>
  `<path d="M0 0${' L1 1'.repeat(n - 1)}"${attrs ? ` ${attrs}` : ''}/>`;

test('an SVG picture whose markers draw past the limit at the vertices draws nothing', () => {
  const marker = (body: string) => `<defs><marker id="m">${body}</marker></defs>`;
  // A marker of 100 elements at each of 5,000 vertices draws 500,000 elements.
  const dots = '<circle r="1"/>'.repeat(99);
  expect(svg(svgDoc(marker(dots) + path(5_000, 'marker-mid="url(#m)"')))).toBeUndefined();
  // The same marker from a style, from an ancestor, or from a style sheet.
  expect(svg(svgDoc(marker(dots) + path(5_000, 'style="marker-mid:url(#m)"')))).toBeUndefined();
  expect(svg(svgDoc(`${marker(dots)}<g marker-end="url(#m)">${path(5_000)}</g>`))).toBe(undefined);
  const sheet = '<style>path { marker: url(#m) }</style>';
  expect(svg(svgDoc(sheet + marker(dots) + path(5_000)))).toBeUndefined();
  // A polyline counts its points.
  const points = Array.from({ length: 5_000 }, (_, i) => `${i},0`).join(' ');
  expect(
    svg(svgDoc(`${marker(dots)}<polyline points="${points}" marker-mid="url(#m)"/>`))
  ).toBeUndefined();
  // An arrowhead at each end of a few hundred lines draws.
  const arrow = marker('<path d="M0 0 L10 5 L0 10 z"/>');
  const lines = Array.from(
    { length: 300 },
    (_, i) => `<line x1="0" y1="${i}" x2="10" y2="${i}" marker-end="url(#m)"/>`
  ).join('');
  expect(svg(svgDoc(arrow + lines))).toBe('image/svg+xml');
});

test('an SVG picture with a marker inside marker content draws nothing', () => {
  const nested =
    '<defs><marker id="a"><circle r="1"/></marker>' +
    `<marker id="b">${path(3, 'marker-mid="url(#a)"')}</marker></defs>` +
    path(3, 'marker-mid="url(#b)"');
  expect(svg(svgDoc(nested))).toBeUndefined();
  // A marker property on the marker's ancestor reaches its content.
  const inherited = `<g marker-mid="url(#a)"><marker id="a">${path(3)}</marker></g>${path(3)}`;
  expect(svg(svgDoc(inherited))).toBeUndefined();
  // Marker content that links to a path with a marker.
  const linked =
    `<defs>${path(3, 'id="p" marker-mid="url(#a)"')}` +
    '<marker id="a"><use href="#p"/></marker></defs><use href="#p"/>';
  expect(svg(svgDoc(linked))).toBeUndefined();
});

test('an SVG picture with a plain document type declaration draws', () => {
  const doctype =
    '<!DOCTYPE svg PUBLIC "-//W3C//DTD SVG 1.1//EN" ' +
    '"http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd">';
  const plain = text(`${doctype}<svg><rect/></svg>`);
  expect(drawableImageMime(plain, 'image/svg+xml')).toBe('image/svg+xml');
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
