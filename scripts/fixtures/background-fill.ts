// The `background-fill` fixture: shapes with `useBgFill` (ECMA-376 Part 1, 19.3.1.43) over a
// solid, a gradient, and a stretched picture background. Each `useBgFill` shape has a style
// fill of `accent1`, which must not show, and covers part of an orange shape, which it hides.

import { para, run, shape, solid, type FixtureOptions } from './builder.ts';

const IN = 914_400;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

const STYLE =
  '<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr>' +
  '</a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>' +
  '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>' +
  '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>';

/** A rectangle that fills with the slide background, with no outline. */
function backgroundShape(id: number, x: number, y: number, cx: number, cy: number): string {
  return shape({
    id,
    name: `Background fill ${id}`,
    x,
    y,
    cx,
    cy,
    line: '<a:ln><a:noFill/></a:ln>',
    style: STYLE,
  }).replace('<p:sp>', '<p:sp useBgFill="1">');
}

function label(id: number, y: number, text: string, color = '1F3864'): string {
  return shape({
    id,
    x: 0.6 * IN,
    y,
    cx: 12 * IN,
    cy: 0.6 * IN,
    text: para(run(text, 'sz="2000"', `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`)),
    bodyPr: '<a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"><a:noAutofit/></a:bodyPr>',
  });
}

/** An orange shape, half under a `useBgFill` shape that hides the covered half. */
function covered(id: number): string {
  return (
    shape({ id, x: 1 * IN, y: 2 * IN, cx: 6 * IN, cy: 3.4 * IN, fill: solid('ED7D31') }) +
    backgroundShape(id + 1, 4 * IN, 1.6 * IN, 8 * IN, 4.4 * IN)
  );
}

const SOLID_BG =
  '<p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>';
const GRADIENT_BG =
  '<p:bg><p:bgPr><a:gradFill rotWithShape="1"><a:gsLst>' +
  '<a:gs pos="0"><a:srgbClr val="DDEBF7"/></a:gs><a:gs pos="100000"><a:srgbClr val="548235"/></a:gs>' +
  '</a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill><a:effectLst/></p:bgPr></p:bg>';
const PICTURE_BG =
  '<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rIdImg"/><a:srcRect/>' +
  '<a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>';

export function backgroundFillDeck(image: Uint8Array): FixtureOptions {
  return {
    media: [{ name: 'image1.png', bytes: image, contentType: 'image/png' }],
    slides: [
      {
        // The full-slide design element of a generated deck, under text in `accent1`.
        layout: 'blank',
        background: SOLID_BG,
        shapes:
          backgroundShape(2, 0, 0, 12_192_000, 6_858_000) +
          label(3, 0.5 * IN, 'Text in accent 1 over a background-filled shape', '4472C4') +
          covered(4) +
          label(6, 6.4 * IN, 'A solid background hides the right half of the orange shape'),
      },
      {
        layout: 'blank',
        background: GRADIENT_BG,
        shapes:
          label(2, 0.5 * IN, 'A gradient background keeps its place on the slide') +
          covered(3) +
          label(5, 6.4 * IN, 'The gradient continues across the edge of the covering shape'),
      },
      {
        layout: 'blank',
        background: PICTURE_BG,
        rels: [['rIdImg', R_IMAGE, '../media/image1.png']],
        shapes:
          label(2, 0.5 * IN, 'A stretched picture background keeps its place on the slide') +
          covered(3) +
          label(5, 6.4 * IN, 'The picture continues across the edge of the covering shape'),
      },
    ],
  };
}
