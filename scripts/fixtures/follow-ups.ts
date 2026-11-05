// The `follow-ups` fixture: a custom path that starts with `lnTo` and then moves (`lnTo`,
// `moveTo`, `lnTo`; ECMA-376 Part 1, 20.1.9.15 path), the `gray` color transform on pure red,
// green, and blue (20.1.2.3.8 gray), every preset pattern (`ST_PresetPatternVal`, 20.1.10.51) on
// an unrotated shape, and patterns on rotated and flipped shapes (20.1.8.47 pattFill).

import { shape, solid, type FixtureSlide } from './builder.ts';

const IN = 914_400;

/** The 54 values of `ST_PresetPatternVal`, in schema order. */
export const PRESET_PATTERNS = [
  'pct5', 'pct10', 'pct20', 'pct25', 'pct30', 'pct40', 'pct50', 'pct60', 'pct70', 'pct75',
  'pct80', 'pct90', 'horz', 'vert', 'ltHorz', 'ltVert', 'dkHorz', 'dkVert', 'narHorz', 'narVert',
  'dashHorz', 'dashVert', 'cross', 'dnDiag', 'upDiag', 'ltDnDiag', 'ltUpDiag', 'dkDnDiag',
  'dkUpDiag', 'wdDnDiag', 'wdUpDiag', 'dashDnDiag', 'dashUpDiag', 'diagCross', 'smCheck',
  'lgCheck', 'smGrid', 'lgGrid', 'dotGrid', 'smConfetti', 'lgConfetti', 'horzBrick', 'diagBrick',
  'solidDmnd', 'openDmnd', 'dotDmnd', 'plaid', 'sphere', 'weave', 'divot', 'shingle', 'wave',
  'trellis', 'zigZag',
] as const; // prettier-ignore

const patt = (prst: string, fg = '1F3864', bg = 'FFFFFF'): string =>
  `<a:pattFill prst="${prst}"><a:fgClr><a:srgbClr val="${fg}"/></a:fgClr><a:bgClr><a:srgbClr val="${bg}"/></a:bgClr></a:pattFill>`;
const outline = (hex = '404040', w = 6350): string =>
  `<a:ln w="${w}"><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></a:ln>`;

/** A freeform from path commands in a 100 × 100 path space. */
function freeform(id: number, x: number, y: number, path: string, pathAttrs: string, fill: string, ln: string): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Freeform ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${Math.round(x * IN)}" y="${Math.round(y * IN)}"/><a:ext cx="${Math.round(2.8 * IN)}" cy="${Math.round(2.8 * IN)}"/></a:xfrm>` +
    '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>' +
    `<a:pathLst><a:path w="100" h="100"${pathAttrs}>${path}</a:path></a:pathLst></a:custGeom>${fill}${ln}</p:spPr></p:sp>`
  );
}

const pt = (x: number, y: number): string => `<a:pt x="${x}" y="${y}"/>`;
const LN_MOVE_LN = `<a:lnTo>${pt(50, 0)}</a:lnTo><a:moveTo>${pt(0, 100)}</a:moveTo><a:lnTo>${pt(100, 20)}</a:lnTo>`;
const LN_MOVE_LN_CLOSED =
  `<a:lnTo>${pt(100, 0)}</a:lnTo><a:moveTo>${pt(10, 90)}</a:moveTo><a:lnTo>${pt(90, 90)}</a:lnTo>` +
  `<a:lnTo>${pt(50, 10)}</a:lnTo><a:close/>`;
const MOVE_LN_MOVE_LN =
  `<a:moveTo>${pt(0, 20)}</a:moveTo><a:lnTo>${pt(100, 20)}</a:lnTo>` +
  `<a:moveTo>${pt(0, 80)}</a:moveTo><a:lnTo>${pt(100, 50)}</a:lnTo>`;
const STROKE = outline('843C0C', 38100);

/** Rotated and flipped shapes with a pattern: `[preset, geometry, rot in degrees, flipH, flipV]`. */
const TURNED: readonly (readonly [string, string, number, boolean, boolean])[] = [
  ['dkDnDiag', 'rect', 30, false, false],
  ['horz', 'rect', 45, false, false],
  ['smCheck', 'rect', 90, false, false],
  ['wave', 'rect', 0, true, false],
  ['horzBrick', 'rect', 0, false, true],
  ['ltVert', 'rect', 0, true, true],
  ['diagBrick', 'rect', 15, true, false],
  ['zigZag', 'rect', 120, false, false],
  ['lgCheck', 'ellipse', 30, false, false],
  ['weave', 'triangle', 0, false, true],
  ['trellis', 'rightArrow', 200, false, false],
  ['upDiag', 'roundRect', 330, true, true],
];

const followUps: FixtureSlide[] = [
  {
    // Paths that start with `lnTo` and then move: stroked only, filled, and closed; and two
    // subpaths, each started by `moveTo`. Then pure red, green, and blue, above the same colors
    // with `a:gray`.
    shapes:
      freeform(2, 0.4, 0.3, LN_MOVE_LN, ' fill="none"', '', STROKE) +
      freeform(3, 3.5, 0.3, LN_MOVE_LN, '', solid('F4B183'), STROKE) +
      freeform(4, 6.6, 0.3, LN_MOVE_LN_CLOSED, '', solid('A9D18E'), STROKE) +
      freeform(5, 9.7, 0.3, MOVE_LN_MOVE_LN, ' fill="none"', '', STROKE) +
      ['FF0000', '00FF00', '0000FF']
        .map(
          (hex, i) =>
            shape({ id: 10 + i, x: (0.4 + i * 4.2) * IN, y: 3.5 * IN, cx: 3.9 * IN, cy: 1.7 * IN, fill: solid(hex) }) +
            shape({
              id: 20 + i,
              x: (0.4 + i * 4.2) * IN,
              y: 5.4 * IN,
              cx: 3.9 * IN,
              cy: 1.7 * IN,
              fill: `<a:solidFill><a:srgbClr val="${hex}"><a:gray/></a:srgbClr></a:solidFill>`,
            })
        )
        .join(''),
  },
  {
    // Every preset pattern on an unrotated rectangle, in schema order: 9 columns, 6 rows.
    shapes: PRESET_PATTERNS.map((prst, i) =>
      shape({
        id: 2 + i,
        x: (0.3 + (i % 9) * 1.42) * IN,
        y: (0.3 + Math.floor(i / 9) * 1.17) * IN,
        cx: 1.27 * IN,
        cy: 1.02 * IN,
        fill: patt(prst),
        line: outline(),
      })
    ).join(''),
  },
  {
    // Patterns on rotated and flipped shapes, with other colors.
    shapes: TURNED.map(([prst, geometry, rot, flipH, flipV], i) =>
      shape({
        id: 2 + i,
        x: (0.6 + (i % 4) * 3.15) * IN,
        y: (0.5 + Math.floor(i / 4) * 2.35) * IN,
        cx: 2.4 * IN,
        cy: 1.5 * IN,
        prst: geometry,
        rot: rot * 60000,
        flipH,
        flipV,
        fill: patt(prst, 'C00000', 'FFF2CC'),
        line: outline('7F6000', 12700),
      })
    ).join(''),
  },
];

export const followUpFixtures: Record<string, FixtureSlide[]> = { 'follow-ups': followUps };
