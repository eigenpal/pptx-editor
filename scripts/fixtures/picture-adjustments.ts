// The `picture-adjustments` fixture: picture borders with dashes, caps, and joins, the 2010
// picture corrections (`a14:imgProps`: saturation, color temperature, brightness and contrast),
// and the brightness and contrast of `a:lum`.
// Pictures draw at 96 DPI, near their pixel size, so image scaling stays out of the measurement.

import { PNG } from 'pngjs';
import type { FixtureOptions } from './builder.ts';

const IN = 914_400;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const A14 = 'http://schemas.microsoft.com/office/drawing/2010/main';
const IMG_PROPS_EXT = '{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}';

/** The image size in pixels: 2.5 × 1.5 in at 96 DPI. */
const W = 240;
const H = 144;

/** Twelve flat color blocks, six per row: mid-saturated colors above, light and gray below. */
const BLOCKS = [
  ['C04040', '40A040', '4060C0', 'D0A030', '30A0A0', 'A040A0'],
  ['F0C0A0', '808080', '404040', 'C0C0C0', '80C0E0', 'E08080'],
];

/** A deterministic 240 × 144 PNG of flat color blocks. */
export function adjustmentsPng(): Uint8Array {
  const png = new PNG({ width: W, height: H });
  for (let y = 0; y < H; y += 1)
    for (let x = 0; x < W; x += 1) {
      const hex = BLOCKS[Math.floor(y / (H / 2))]![Math.floor(x / (W / 6))]!;
      const i = (y * W + x) * 4;
      png.data[i] = parseInt(hex.slice(0, 2), 16);
      png.data[i + 1] = parseInt(hex.slice(2, 4), 16);
      png.data[i + 2] = parseInt(hex.slice(4, 6), 16);
      png.data[i + 3] = 255;
    }
  return new Uint8Array(PNG.sync.write(png));
}

/** A second image, one flat teal block. */
function tealPng(): Uint8Array {
  const png = new PNG({ width: W, height: H });
  for (let i = 0; i < W * H * 4; i += 4) png.data.set([0, 128, 128, 255], i);
  return new Uint8Array(PNG.sync.write(png));
}

/**
 * The 2010 picture layer of a blip: the image `layer` with `effects`, each in its own
 * `a14:imgEffect`, which holds one effect.
 */
const imgProps = (effects: string | readonly string[], layer = 'rIdImg'): string =>
  `<a:extLst><a:ext uri="${IMG_PROPS_EXT}"><a14:imgProps xmlns:a14="${A14}">` +
  `<a14:imgLayer r:embed="${layer}">` +
  [effects].flat().map((e) => `<a14:imgEffect>${e}</a14:imgEffect>`).join('') +
  '</a14:imgLayer></a14:imgProps></a:ext></a:extLst>';

/** One picture of 2.5 × 1.5 in at column `col` and row `row` of a 3 × 2 grid. */
const pic = (id: number, col: number, row: number, blip: string, line = ''): string =>
  `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
  `<p:blipFill>${blip}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  `<p:spPr><a:xfrm><a:off x="${Math.round((1 + col * 4) * IN)}" y="${Math.round((1 + row * 3) * IN)}"/>` +
  `<a:ext cx="${2.5 * IN}" cy="${1.5 * IN}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${line}</p:spPr></p:pic>`;

const blip = (inner = '', embed = 'rIdImg'): string => `<a:blip r:embed="${embed}">${inner}</a:blip>`;

const border = (w: number, dash: string, cap: string, join = '<a:miter lim="800000"/>'): string =>
  `<a:ln w="${w}" cap="${cap}"><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill>` +
  `<a:prstDash val="${dash}"/>${join}</a:ln>`;

/** One slide of six pictures from `[blip inner XML, outline XML]` pairs. */
const grid = (items: readonly (readonly [string, string?])[]): string =>
  items.map(([inner, line], i) => pic(2 + i, i % 3, Math.floor(i / 3), inner, line)).join('');

export function pictureAdjustmentsDeck(): FixtureOptions {
  const rels = [['rIdImg', R_IMAGE, '../media/image1.png']] as [string, string, string][];
  return {
    slides: [
      {
        // Border dashes and caps at 3 pt and 9 pt.
        rels,
        shapes: grid([
          [blip(), border(38100, 'sysDot', 'rnd')],
          [blip(), border(38100, 'dash', 'flat')],
          [blip(), border(38100, 'lgDashDot', 'sq')],
          [blip(), border(114300, 'sysDot', 'rnd')],
          [blip(), border(114300, 'dash', 'flat')],
          [blip(), border(114300, 'lgDashDot', 'rnd')],
        ]),
      },
      {
        // More caps at 9 pt, and the round and bevel joins of solid borders.
        rels,
        shapes: grid([
          [blip(), border(114300, 'dash', 'sq')],
          [blip(), border(114300, 'sysDash', 'rnd')],
          [blip(), border(114300, 'dashDot', 'rnd')],
          [blip(), border(114300, 'sysDot', 'flat')],
          [blip(), border(114300, 'solid', 'rnd', '<a:round/>')],
          [blip(), border(114300, 'solid', 'flat', '<a:bevel/>')],
        ]),
      },
      {
        // Saturation 0 %, 50 %, and 200 %, temperatures 4700 K and 11200 K, and no change.
        // The blip's own image draws: a writer stores it with the effects applied.
        rels,
        shapes: grid([
          [blip(imgProps('<a14:saturation sat="0"/>'))],
          [blip(imgProps('<a14:saturation sat="50000"/>'))],
          [blip(imgProps('<a14:saturation sat="200000"/>'))],
          [blip(imgProps('<a14:colorTemperature colorTemp="4700"/>'))],
          [blip(imgProps('<a14:colorTemperature colorTemp="11200"/>'))],
          [blip()],
        ]),
      },
      {
        // Brightness and contrast, alone, together, and after saturation and temperature.
        rels,
        shapes: grid([
          [blip(imgProps('<a14:brightnessContrast bright="20000"/>'))],
          [blip(imgProps('<a14:brightnessContrast contrast="40000"/>'))],
          [blip(imgProps('<a14:brightnessContrast bright="-20000" contrast="-40000"/>'))],
          [blip(imgProps('<a14:brightnessContrast bright="40000" contrast="40000"/>'))],
          [blip(imgProps(['<a14:saturation sat="50000"/>', '<a14:brightnessContrast bright="-20000" contrast="20000"/>']))],
          [blip(imgProps(['<a14:colorTemperature colorTemp="11200"/>', '<a14:saturation sat="200000"/>']))],
        ]),
      },
      {
        // A layer image that differs from the blip's image: the blip's image draws. Then the
        // brightness and contrast of `a:lum`, which the blip's image does not include.
        rels: [...rels, ['rIdTeal', R_IMAGE, '../media/image2.png']],
        shapes: grid([
          [blip(imgProps('<a14:saturation sat="0"/>', 'rIdTeal'))],
          [blip(imgProps('<a14:saturation sat="0"/>'), 'rIdTeal')],
          [blip('<a:lum bright="20000"/>')],
          [blip('<a:lum contrast="40000"/>')],
          [blip('<a:lum bright="-20000" contrast="-40000"/>')],
          [blip('<a:lum bright="40000" contrast="40000"/>')],
        ]),
      },
    ],
    media: [
      { name: 'image1.png', bytes: adjustmentsPng(), contentType: 'image/png' },
      { name: 'image2.png', bytes: tealPng(), contentType: 'image/png' },
    ],
  };
}
