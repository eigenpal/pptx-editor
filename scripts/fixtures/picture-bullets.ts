// Picture bullets (`a:buBlip`, ECMA-376 Part 1, 21.1.2.4.2): a square and a wide picture at
// several bullet sizes (`a:buSzPct`, `a:buSzPts`), text sizes, levels, and alignments, and a
// picture bullet from the text body's list style.

import { PNG } from 'pngjs';
import { para, run, shape, solid, type FixtureOptions } from './builder.ts';

const IN = 914_400;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

/** A 32×32 picture: a red disc with a dark ring on a transparent ground. */
function discPng(): Uint8Array {
  const png = new PNG({ width: 32, height: 32 });
  for (let y = 0; y < 32; y += 1)
    for (let x = 0; x < 32; x += 1) {
      const i = (y * 32 + x) * 4;
      const r = Math.hypot(x - 15.5, y - 15.5);
      const ring = r > 12 && r <= 15.5;
      const disc = r <= 12;
      png.data.set(ring ? [60, 20, 20, 255] : disc ? [220, 40, 40, 255] : [0, 0, 0, 0], i);
    }
  return new Uint8Array(PNG.sync.write(png));
}

/** A 64×32 picture: a blue left half and an orange right half. */
function widePng(): Uint8Array {
  const png = new PNG({ width: 64, height: 32 });
  for (let y = 0; y < 32; y += 1)
    for (let x = 0; x < 64; x += 1)
      png.data.set(x < 32 ? [40, 90, 200, 255] : [240, 150, 30, 255], (y * 64 + x) * 4);
  return new Uint8Array(PNG.sync.write(png));
}

const blip = (rId: string): string => `<a:buBlip><a:blip r:embed="${rId}"/></a:buBlip>`;
const item = (text: string, pPr: string, rPr = 'sz="1800"'): string =>
  para(run(text, rPr), `<a:pPr marL="342900" indent="-342900"${pPr}`);

const box = (id: number, x: number, body: string, lstStyle = ''): string =>
  shape({
    id,
    x,
    y: 0.5 * IN,
    cx: 5.6 * IN,
    cy: 6.2 * IN,
    fill: solid('F2F2F2'),
    text: body,
    bodyPr: '<a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720"><a:noAutofit/></a:bodyPr>',
    ...(lstStyle ? { lstStyle } : {}),
  });

const faces = (face: string): string => {
  const font = `<a:latin typeface="${face}"/>`;
  return (
    item(`${face} 18 pt`, `>${blip('rIdDisc')}</a:pPr>`, `sz="1800"`).replace('</a:rPr>', `${font}</a:rPr>`) +
    item(`${face} 28 pt`, `>${blip('rIdDisc')}</a:pPr>`, `sz="2800"`).replace('</a:rPr>', `${font}</a:rPr>`) +
    item(`${face} 40 pt`, `>${blip('rIdWide')}</a:pPr>`, `sz="4000"`).replace('</a:rPr>', `${font}</a:rPr>`) +
    item(`${face} 12 pt, 300%`, `><a:buSzPct val="300000"/>${blip('rIdDisc')}</a:pPr>`, `sz="1200"`).replace('</a:rPr>', `${font}</a:rPr>`)
  );
};

/** The picture bullets deck. */
export function pictureBulletsDeck(): FixtureOptions {
  const left =
    item('Disc at 100%', `>${blip('rIdDisc')}</a:pPr>`) +
    item('Disc at 150%', `><a:buSzPct val="150000"/>${blip('rIdDisc')}</a:pPr>`) +
    item('Disc at 60%', `><a:buSzPct val="60000"/>${blip('rIdDisc')}</a:pPr>`) +
    item('Wide picture at 100%', `>${blip('rIdWide')}</a:pPr>`) +
    item('Disc at 24 pt', `><a:buSzPts val="2400"/>${blip('rIdDisc')}</a:pPr>`) +
    item('Disc, 32 pt text', `>${blip('rIdDisc')}</a:pPr>`, 'sz="3200"') +
    item(
      'A long item that wraps to a second line under its picture bullet',
      `>${blip('rIdWide')}</a:pPr>`
    );
  // Level 2 takes its picture bullet from the list style of the text body.
  const lstStyle =
    '<a:lvl2pPr marL="800100" indent="-342900"><a:buSzPct val="120000"/>' +
    `${blip('rIdWide')}</a:lvl2pPr>`;
  const right =
    item('Level one with the disc', `>${blip('rIdDisc')}</a:pPr>`) +
    para(run('Level two from the list style', 'sz="1800"'), '<a:pPr lvl="1"/>') +
    para(run('Level two again', 'sz="1800"'), '<a:pPr lvl="1"/>') +
    item('Centered with the disc', ` algn="ctr">${blip('rIdDisc')}</a:pPr>`) +
    item('Right with the wide picture', ` algn="r">${blip('rIdWide')}</a:pPr>`) +
    item('Disc at 200%, 14 pt text', `><a:buSzPct val="200000"/>${blip('rIdDisc')}</a:pPr>`, 'sz="1400"');
  return {
    slides: [
      {
        shapes: box(2, 0.4 * IN, left) + box(3, 6.4 * IN, right, lstStyle),
        rels: [
          ['rIdDisc', R_IMAGE, '../media/image1.png'],
          ['rIdWide', R_IMAGE, '../media/image2.png'],
        ],
      },
      {
        // The same pictures with other faces: the picture size follows the bullet size only.
        shapes: box(2, 0.4 * IN, faces('Liberation Serif')) + box(3, 6.4 * IN, faces('Liberation Sans')),
        rels: [
          ['rIdDisc', R_IMAGE, '../media/image1.png'],
          ['rIdWide', R_IMAGE, '../media/image2.png'],
        ],
      },
    ],
    media: [
      { name: 'image1.png', bytes: discPng(), contentType: 'image/png' },
      { name: 'image2.png', bytes: widePng(), contentType: 'image/png' },
    ],
  };
}
