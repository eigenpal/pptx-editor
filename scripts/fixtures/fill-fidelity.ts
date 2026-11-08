// The `fill-fidelity` fixture: one slide per fill rule, scored
// against its reference PDF (`fixtures/fill-fidelity.pdf`).

import { crc32 } from 'node:zlib';
import { shape, type FixtureOptions } from './builder.ts';

const IN = 914_400;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const IMG: [string, string, string] = ['rIdImg', R_IMAGE, '../media/image1.png'];
const IMG_DPI: [string, string, string] = ['rIdDpi', R_IMAGE, '../media/image2.png'];

/** The same PNG with a `pHYs` chunk (pixels per metre) after `IHDR`: a stated resolution. */
export function withResolution(png: Uint8Array, dpi: number): Uint8Array {
  const ppm = Math.round(dpi / 0.0254);
  const body = new Uint8Array(13);
  body.set([0x70, 0x48, 0x59, 0x73]); // pHYs
  const view = new DataView(body.buffer);
  view.setUint32(4, ppm);
  view.setUint32(8, ppm);
  body[12] = 1; // the unit is the metre
  const chunk = new Uint8Array(4 + 13 + 4);
  const cv = new DataView(chunk.buffer);
  cv.setUint32(0, 9);
  chunk.set(body, 4);
  cv.setUint32(17, crc32(body));
  const ihdrEnd = 8 + 4 + 4 + 13 + 4;
  const out = new Uint8Array(png.length + chunk.length);
  out.set(png.subarray(0, ihdrEnd));
  out.set(chunk, ihdrEnd);
  out.set(png.subarray(ihdrEnd), ihdrEnd + chunk.length);
  return out;
}

const srgb = (hex: string, mods = ''): string => `<a:srgbClr val="${hex}">${mods}</a:srgbClr>`;
const stops = (...s: [number, string][]): string =>
  `<a:gsLst>${s.map(([pos, c]) => `<a:gs pos="${pos * 1000}">${c}</a:gs>`).join('')}</a:gsLst>`;
const BLUE_ORANGE = stops([0, srgb('2E4BE6')], [100, srgb('F2A23A')]);
const WHITE_RED = stops([0, srgb('FFFFFF')], [100, srgb('C00000')]);
const LIGHT_DARK = stops([0, srgb('DEEBF7')], [60, srgb('5B9BD5')], [100, srgb('1F4E79')]);
const lin = (ang: number, scaled: 0 | 1, s = BLUE_ORANGE): string =>
  `<a:gradFill rotWithShape="1">${s}<a:lin ang="${ang * 60000}" scaled="${scaled}"/></a:gradFill>`;
const pathFill = (kind: string, ftr: string, s = WHITE_RED, tile = ''): string =>
  `<a:gradFill rotWithShape="1">${s}<a:path path="${kind}"><a:fillToRect ${ftr}/></a:path>${tile}</a:gradFill>`;
const CENTER = 'l="50000" t="50000" r="50000" b="50000"';
const at = (id: number, x: number, y: number, cx: number, cy: number, fill: string, prst = 'rect'): string =>
  shape({ id, x: x * IN, y: y * IN, cx: cx * IN, cy: cy * IN, prst, fill });

const pic = (id: number, x: number, y: number, cx: number, cy: number, blip: string): string =>
  `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
  `<p:blipFill>${blip}</p:blipFill><p:spPr><a:xfrm><a:off x="${Math.round(x * IN)}" y="${Math.round(y * IN)}"/>` +
  `<a:ext cx="${Math.round(cx * IN)}" cy="${Math.round(cy * IN)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;

/**
 * The resolution of a tiled picture in the fixtures. A PNG without `pHYs` has no stated
 * resolution, and presentation apps tile it at different resolutions, so every tiled PNG states one.
 */
export const TILE_FIXTURE_DPI = 144;

/** The deck: linear gradients, path gradients, tiles and overlays, stretched pictures, backgrounds. */
export function fillFidelityDeck(image: Uint8Array): FixtureOptions {
  return {
    media: [
      { name: 'image1.png', bytes: withResolution(image, TILE_FIXTURE_DPI), contentType: 'image/png' },
      { name: 'image2.png', bytes: withResolution(image, 150), contentType: 'image/png' },
    ],
    slides: [
      {
        // Linear gradients: scaled stretches with the box (D1); a tile rectangle from a corner (D4).
        shapes:
          at(2, 0.4, 0.4, 6, 1.6, lin(315, 1)) +
          at(3, 0.4, 2.3, 6, 1.6, lin(315, 0)) +
          at(4, 0.4, 4.2, 6, 1.6, lin(30, 1, LIGHT_DARK)) +
          at(5, 6.9, 0.4, 1.6, 3.4, lin(45, 1)) +
          at(6, 8.8, 0.4, 1.6, 3.4, lin(45, 0)) +
          at(7, 6.9, 4.2, 2.6, 2.6, pathFill('circle', 'r="100000" b="100000"', LIGHT_DARK, '<a:tileRect r="-100000" b="-100000"/>')) +
          at(8, 9.8, 4.2, 3, 1.4, pathFill('circle', 't="100000" r="100000"', LIGHT_DARK, '<a:tileRect l="-100000" b="-100000"/>')),
      },
      {
        // Path gradients: centered, outside, and off-center foci (D2), shape outlines (D3).
        shapes:
          at(2, 0.4, 0.4, 4, 2.4, pathFill('circle', CENTER)) +
          at(3, 4.7, 0.4, 4, 2.4, pathFill('circle', 'l="50000" t="-80000" r="50000" b="180000"')) +
          at(4, 9, 0.4, 3.9, 2.4, pathFill('circle', 'l="25000" t="50000" r="75000" b="50000"')) +
          at(5, 0.4, 3.2, 4, 2.4, pathFill('rect', 'l="30000" t="30000" r="70000" b="70000"', LIGHT_DARK)) +
          at(6, 4.7, 3.2, 2.6, 2.4, pathFill('shape', CENTER), 'triangle') +
          at(7, 7.6, 3.2, 2.6, 2.4, pathFill('shape', CENTER, LIGHT_DARK), 'star5') +
          at(8, 10.5, 3.2, 2.4, 2.4, pathFill('shape', CENTER, stops([0, srgb('70AD47')], [100, srgb('70AD47', '<a:alpha val="20000"/>')])), 'roundRect'),
      },
      {
        // Tiles: 72 DPI without a stated resolution, a stated 150 DPI, flips, a duotone (D6);
        // fill overlays (D5).
        rels: [IMG, IMG_DPI],
        shapes:
          at(2, 0.4, 0.4, 4, 2.6, '<a:blipFill><a:blip r:embed="rIdImg"/><a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/></a:blipFill>') +
          at(3, 4.7, 0.4, 4, 2.6, '<a:blipFill><a:blip r:embed="rIdDpi"/><a:tile tx="76200" ty="38100" sx="100000" sy="100000" flip="x" algn="br"/></a:blipFill>') +
          at(4, 9, 0.4, 3.9, 2.6, '<a:blipFill><a:blip r:embed="rIdImg"><a:duotone><a:srgbClr val="385723"/><a:srgbClr val="FFFFFF"/></a:duotone></a:blip><a:tile tx="0" ty="0" sx="50000" sy="50000" flip="y" algn="ctr"/></a:blipFill>') +
          shape({
            id: 5, x: 0.4 * IN, y: 3.4 * IN, cx: 4 * IN, cy: 3.6 * IN,
            fill: '<a:blipFill><a:blip r:embed="rIdImg"><a:alphaModFix amt="43000"/></a:blip><a:tile tx="0" ty="0" sx="50000" sy="50000" flip="none" algn="tl"/></a:blipFill>',
            effects: `<a:effectLst><a:fillOverlay blend="mult">${pathFill('circle', 'l="50000" t="120000" r="50000" b="-20000"', stops([0, srgb('E8A0DA')], [100, srgb('3D0634')]))}</a:fillOverlay></a:effectLst>`,
          }) +
          shape({
            id: 6, x: 4.7 * IN, y: 3.4 * IN, cx: 4 * IN, cy: 3.6 * IN, prst: 'ellipse',
            fill: '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>',
            effects: `<a:effectLst><a:fillOverlay blend="screen">${lin(90, 0, stops([0, srgb('000000')], [100, srgb('FFC000')]))}</a:fillOverlay></a:effectLst>`,
          }) +
          shape({
            id: 7, x: 9 * IN, y: 3.4 * IN, cx: 3.9 * IN, cy: 3.6 * IN,
            fill: '<a:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>',
            effects: `<a:effectLst><a:fillOverlay blend="darken"><a:solidFill>${srgb('808080')}</a:solidFill></a:fillOverlay></a:effectLst>`,
          }),
      },
      {
        // Stretched pictures: fillRect insets and extensions, srcRect, and a picture that starts
        // left of the slide (D7).
        rels: [IMG],
        shapes:
          pic(2, -2, 0.4, 7, 3.2, '<a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch>') +
          at(3, 5.6, 0.4, 3.4, 3.2, '<a:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect l="10000" t="20000" r="10000" b="20000"/></a:stretch></a:blipFill>', 'ellipse') +
          at(4, 9.4, 0.4, 3.4, 3.2, '<a:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect l="-25000" r="-25000"/></a:stretch></a:blipFill>') +
          pic(5, 0.4, 4, 5.6, 3, '<a:blip r:embed="rIdImg"/><a:srcRect l="-10000" t="15000" r="-10000" b="15000"/><a:stretch><a:fillRect/></a:stretch>') +
          pic(6, 6.4, 4, 6.5, 3, '<a:blip r:embed="rIdImg"/><a:srcRect l="25000" t="10000" r="10000" b="20000"/><a:stretch><a:fillRect/></a:stretch>'),
      },
      {
        // A background path gradient uses the same gradient paint servers.
        background: `<p:bg><p:bgPr>${pathFill('circle', CENTER, LIGHT_DARK)}<a:effectLst/></p:bgPr></p:bg>`,
        shapes: at(2, 4.2, 2.2, 5, 3, lin(315, 1, WHITE_RED)),
      },
      {
        // A tiled background picture.
        rels: [IMG],
        background: '<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rIdImg"/><a:srcRect/><a:tile tx="0" ty="0" sx="75000" sy="75000" flip="xy" algn="tl"/></a:blipFill><a:effectLst/></p:bgPr></p:bg>',
        shapes: '',
      },
    ],
  };
}
