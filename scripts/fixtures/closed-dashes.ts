// The `closed-dashes` fixture: dashed outlines with round and square caps on closed shapes
// other than a rectangle, where the dash pattern starts (`core/layout`, `line-strokes.ts`).
// One slide per line width (3 pt and 9 pt): six shapes across, three dash styles down.

import { shape, type FixtureOptions } from './builder.ts';

const IN = 914_400;

/** The dash styles, one per row: `[prstDash, cap]`. */
const STYLES: readonly (readonly [string, string])[] = [
  ['sysDot', 'rnd'],
  ['dash', 'sq'],
  ['lgDashDot', 'rnd'],
];

const SHAPES = ['ellipse', 'roundRect', 'star5', 'triangle', 'hexagon', 'custom'] as const;

/** A closed custom path: lines, a cubic curve, and a close back to its first point's segment. */
const CUSTOM_GEOMETRY =
  '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>' +
  '<a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="10" y="20"/></a:moveTo>' +
  '<a:lnTo><a:pt x="70" y="0"/></a:lnTo><a:lnTo><a:pt x="100" y="55"/></a:lnTo>' +
  '<a:cubicBezTo><a:pt x="90" y="90"/><a:pt x="50" y="110"/><a:pt x="30" y="90"/></a:cubicBezTo>' +
  '<a:lnTo><a:pt x="0" y="70"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>';

const line = (w: number, dash: string, cap: string): string =>
  `<a:ln w="${w}" cap="${cap}"><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill>` +
  `<a:prstDash val="${dash}"/><a:round/></a:ln>`;

function slide(w: number): string {
  let id = 2;
  let out = '';
  STYLES.forEach(([dash, cap], row) =>
    SHAPES.forEach((prst, col) => {
      const xml = shape({
        id: id++,
        x: (0.5 + col * 2.1) * IN,
        y: (0.5 + row * 2.35) * IN,
        cx: 1.7 * IN,
        cy: 1.7 * IN,
        prst: prst === 'custom' ? 'rect' : prst,
        fill: '<a:noFill/>',
        line: line(w, dash, cap),
      });
      out +=
        prst === 'custom'
          ? xml.replace('<a:prstGeom prst="rect"><a:avLst></a:avLst></a:prstGeom>', CUSTOM_GEOMETRY)
          : xml;
    })
  );
  return out;
}

/** Circles of 0.8, 1.7, and 3 in and a 3 × 1.2 in ellipse: 9 pt dashes, square caps above, flat below. */
function ellipses(): string {
  let id = 2;
  let out = '';
  let x = 0.4;
  for (const [cx, cy] of [
    [0.8, 0.8],
    [1.7, 1.7],
    [3, 3],
    [3, 1.2],
  ] as const) {
    for (const [y, cap] of [
      [0.3, 'sq'],
      [3.8, 'flat'],
    ] as const)
      out += shape({
        id: id++,
        x: x * IN,
        y: y * IN,
        cx: cx * IN,
        cy: cy * IN,
        prst: 'ellipse',
        fill: '<a:noFill/>',
        line: line(114300, 'dash', cap),
      });
    x += cx + 0.4;
  }
  return out;
}

export function closedDashesDeck(): FixtureOptions {
  return {
    slides: [{ shapes: slide(38100) }, { shapes: slide(114300) }, { shapes: ellipses() }],
  };
}
