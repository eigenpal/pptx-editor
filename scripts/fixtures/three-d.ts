// The 3-D fixture: shape bevels (`a:sp3d/a:bevelT`, ECMA-376 Part 1, 20.1.5.9), contours,
// materials, and light rigs under the front cameras, camera presets with an extrusion
// (`a:scene3d/a:camera`, 20.1.5.5), preset shadows (`a:prstShdw`, 20.1.8.49), and blur
// (`a:blur`, 20.1.8.15). The shapes hold no text, so the pixels alone measure the paint.

import { shape, solid, type FixtureSlide } from './builder.ts';

const IN = 914_400;
const BLUE = '4472C4';
const OUTLINE = '<a:ln w="12700"><a:solidFill><a:srgbClr val="2F528F"/></a:solidFill></a:ln>';

/** `a:scene3d` with a camera preset, its rotation, and a light rig. */
export function scene(camera = 'orthographicFront', rig = 'threePt', dir = 't', rot = ''): string {
  return `<a:scene3d><a:camera prst="${camera}">${rot}</a:camera><a:lightRig rig="${rig}" dir="${dir}"/></a:scene3d>`;
}

/** `a:sp3d` with a top bevel and optional attributes and children. */
export function sp3d(o: {
  bevel?: string;
  w?: number;
  h?: number;
  attrs?: string;
  extra?: string;
}): string {
  const bevel =
    o.bevel === undefined
      ? ''
      : `<a:bevelT w="${o.w ?? 76_200}" h="${o.h ?? 76_200}"${o.bevel === 'circle' ? '' : ` prst="${o.bevel}"`}/>`;
  return `<a:sp3d${o.attrs ?? ''}>${bevel}${o.extra ?? ''}</a:sp3d>`;
}

interface Cell {
  readonly threeD: string;
  readonly prst?: string;
  readonly fill?: string;
  readonly line?: string;
  readonly effects?: string;
}

/** A grid of `cols` columns; each shape is `w` by `h` inches. */
function grid(cells: readonly Cell[], cols: number, w = 1.9, h = 1.3): FixtureSlide {
  const gapX = (13.333 - cols * w) / (cols + 1);
  const rows = Math.ceil(cells.length / cols);
  const gapY = (7.5 - rows * h) / (rows + 1);
  return {
    layout: 'blank',
    shapes: cells
      .map((c, i) =>
        shape({
          id: 2 + i,
          x: (gapX + (i % cols) * (w + gapX)) * IN,
          y: (gapY + Math.floor(i / cols) * (h + gapY)) * IN,
          cx: w * IN,
          cy: h * IN,
          prst: c.prst ?? 'rect',
          fill: solid(c.fill ?? BLUE),
          line: c.line,
          effects: `${c.effects ?? ''}${c.threeD}`,
        })
      )
      .join(''),
  };
}

export const BEVEL_PRESETS = [
  'circle',
  'relaxedInset',
  'cross',
  'coolSlant',
  'angle',
  'softRound',
  'convex',
  'slope',
  'divot',
  'riblet',
  'hardEdge',
  'artDeco',
] as const;

const front = scene();

/** The cameras of the camera slides, each with an extrusion of 18 pt. */
const CAMERAS = [
  'isometricTopUp',
  'isometricTopDown',
  'isometricBottomUp',
  'isometricBottomDown',
  'isometricLeftUp',
  'isometricLeftDown',
  'isometricRightUp',
  'isometricRightDown',
  'isometricOffAxis1Left',
  'isometricOffAxis1Right',
  'isometricOffAxis2Left',
  'isometricOffAxis2Right',
  'obliqueTopLeft',
  'obliqueTop',
  'obliqueBottomRight',
  'obliqueRight',
] as const;

const black = (alpha: number): string =>
  `<a:prstClr val="black"><a:alpha val="${alpha}"/></a:prstClr>`;

export const threeDFixtures: Record<string, FixtureSlide[]> = {
  'three-d-effects': [
    // The twelve bevel presets at 6 pt, front camera, three-point light from the top.
    grid(
      BEVEL_PRESETS.map((bevel) => ({ threeD: front + sp3d({ bevel }) })),
      4
    ),
    // Sizes, geometries, outlines, contours, materials, and light rigs with the circle bevel.
    grid(
      [
        { threeD: front + sp3d({ bevel: 'circle', w: 152_400, h: 152_400 }) },
        { threeD: front + sp3d({ bevel: 'circle', w: 38_100, h: 38_100 }) },
        { threeD: front + sp3d({ bevel: 'angle', w: 127_000, h: 50_800 }) },
        { prst: 'ellipse', threeD: front + sp3d({ bevel: 'circle' }) },
        { prst: 'roundRect', threeD: front + sp3d({ bevel: 'softRound' }) },
        { prst: 'triangle', threeD: front + sp3d({ bevel: 'angle' }) },
        { line: OUTLINE, threeD: front + sp3d({ bevel: 'circle' }) },
        {
          threeD:
            front +
            sp3d({
              bevel: 'circle',
              attrs: ' contourW="38100"',
              extra: '<a:contourClr><a:srgbClr val="C00000"/></a:contourClr>',
            }),
        },
        { threeD: front + sp3d({ bevel: 'circle', attrs: ' prstMaterial="matte"' }) },
        { threeD: front + sp3d({ bevel: 'circle', attrs: ' prstMaterial="metal"' }) },
        { threeD: front + sp3d({ bevel: 'circle', attrs: ' prstMaterial="flat"' }) },
        { threeD: front + sp3d({ bevel: 'circle', attrs: ' prstMaterial="warmMatte"' }) },
        { threeD: scene('orthographicFront', 'balanced', 't') + sp3d({ bevel: 'circle' }) },
        { threeD: scene('orthographicFront', 'harsh', 't') + sp3d({ bevel: 'circle' }) },
        { threeD: scene('orthographicFront', 'threePt', 'tl') + sp3d({ bevel: 'circle' }) },
        { threeD: scene('perspectiveFront') + sp3d({ bevel: 'circle' }) },
      ],
      4,
      1.9,
      1.1
    ),
    // Camera presets with an 18 pt extrusion; no bevel.
    grid(
      CAMERAS.map((camera) => ({
        threeD: scene(camera) + sp3d({ attrs: ' extrusionH="228600"' }),
      })),
      4,
      1.6,
      1.0
    ),
    // Camera rotations with a bevel, a contour, and an extrusion color; front cameras with an
    // extrusion, which does not show.
    grid(
      [
        {
          threeD:
            scene('isometricOffAxis1Right') +
            sp3d({ bevel: 'circle', attrs: ' extrusionH="152400"' }),
        },
        {
          threeD:
            scene('isometricLeftDown') +
            sp3d({
              attrs: ' extrusionH="228600"',
              extra: '<a:extrusionClr><a:srgbClr val="ED7D31"/></a:extrusionClr>',
            }),
        },
        {
          prst: 'ellipse',
          threeD: scene('isometricTopUp') + sp3d({ attrs: ' extrusionH="228600"' }),
        },
        {
          threeD:
            scene('orthographicFront', 'threePt', 't', '<a:rot lat="0" lon="2700000" rev="0"/>') +
            sp3d({ attrs: ' extrusionH="228600"' }),
        },
        { threeD: front + sp3d({ attrs: ' extrusionH="228600"' }) },
        { threeD: scene('perspectiveFront') + sp3d({ attrs: ' extrusionH="228600"' }) },
        { threeD: scene('isometricTopUp') },
        {
          line: OUTLINE,
          threeD: scene('obliqueBottomLeft') + sp3d({ attrs: ' extrusionH="228600"' }),
        },
      ],
      4,
      1.9,
      1.3
    ),
    // The twenty preset shadows, at their default distance and direction, and a set distance.
    grid(
      Array.from({ length: 20 }, (_, i) => ({
        effects: `<a:effectLst><a:prstShdw prst="shdw${i + 1}" dist="76200" dir="2700000">${black(50_000)}</a:prstShdw></a:effectLst>`,
        threeD: '',
        fill: 'FFFFFF',
        line: OUTLINE,
      })),
      5,
      1.5,
      1.0
    ),
    // Blur: radii, the grow attribute, and a blur on an outlined shape.
    grid(
      [
        { effects: '<a:effectLst><a:blur rad="38100"/></a:effectLst>', threeD: '' },
        { effects: '<a:effectLst><a:blur rad="127000"/></a:effectLst>', threeD: '' },
        { effects: '<a:effectLst><a:blur rad="127000" grow="0"/></a:effectLst>', threeD: '' },
        {
          prst: 'ellipse',
          line: OUTLINE,
          effects: '<a:effectLst><a:blur rad="63500"/></a:effectLst>',
          threeD: '',
        },
      ],
      4,
      2.2,
      1.5
    ),
  ],
};
