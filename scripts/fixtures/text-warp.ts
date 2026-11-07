// The text warp fixture: preset text warps (`a:prstTxWarp`, ECMA-376 Part 1, 20.1.9.19) on
// open-font text. Each shape has a thin outline, so the reference shows where its box is. The
// first slide holds the warps that follow one path, the second and third the warps between two
// paths, with one and two lines, adjust values, and other anchors.

import { para, run, shape, solid, type FixtureSlide } from './builder.ts';

const IN = 914_400;
const OUTLINE = '<a:ln w="6350"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln>';
const FONT = '<a:latin typeface="Liberation Sans"/>';

interface Warp {
  readonly prst: string;
  readonly lines: readonly string[];
  readonly color: string;
  readonly av?: string;
  readonly sz?: number;
  readonly algn?: string;
  readonly anchor?: string;
  readonly wrap?: 'none' | 'square';
}

/** One warped shape in a 3 by 2 grid: `col` from 0 to 2, `row` from 0 to 1. */
function warped(id: number, col: number, row: number, w: Warp): string {
  const runProps = (color: string): string => `${solid(color)}${FONT}`;
  const text = w.lines
    .map((line) =>
      para(run(line, `sz="${w.sz ?? 3600}" b="1"`, runProps(w.color)), `<a:pPr algn="${w.algn ?? 'ctr'}"/>`)
    )
    .join('');
  return shape({
    id,
    name: `${w.prst} ${id}`,
    x: (0.45 + col * 4.25) * IN,
    y: (0.55 + row * 3.5) * IN,
    cx: 3.9 * IN,
    cy: 2.7 * IN,
    line: OUTLINE,
    bodyPr:
      `<a:bodyPr wrap="${w.wrap ?? 'none'}" rtlCol="0" anchor="${w.anchor ?? 't'}">` +
      `<a:prstTxWarp prst="${w.prst}"><a:avLst>${w.av ?? ''}</a:avLst></a:prstTxWarp><a:noAutofit/></a:bodyPr>`,
    text,
  });
}

const gd = (name: string, value: number): string => `<a:gd name="${name}" fmla="val ${value}"/>`;

function grid(warps: readonly Warp[]): FixtureSlide {
  return { shapes: warps.map((w, i) => warped(2 + i, i % 3, Math.floor(i / 3), w)).join('') };
}

export const textWarpFixtures: Record<string, FixtureSlide[]> = {
  'text-warp': [
    grid([
      { prst: 'textArchUp', lines: ['Arch up'], color: '1F4E79' },
      { prst: 'textArchDown', lines: ['Arch down'], color: 'C00000' },
      { prst: 'textCircle', lines: ['Text around a circle path'], color: '385723', sz: 2800 },
      { prst: 'textArchUp', lines: ['First line', 'Second line'], color: '7030A0', sz: 2800 },
      { prst: 'textButton', lines: ['Top', 'Middle', 'Bottom'], color: 'ED7D31', sz: 2400 },
      { prst: 'textArchUpPour', lines: ['Pour one', 'Pour two'], color: '2E75B6', sz: 2800 },
    ]),
    grid([
      { prst: 'textWave1', lines: ['Wave text'], color: '1F4E79', av: gd('adj1', 20000) + gd('adj2', 0) },
      { prst: 'textInflate', lines: ['Inflate', 'two lines'], color: 'C00000' },
      { prst: 'textDeflate', lines: ['Deflate'], color: '385723', anchor: 'ctr' },
      { prst: 'textSlantUp', lines: ['Slant up'], color: '7030A0' },
      { prst: 'textTriangle', lines: ['Triangle'], color: 'ED7D31' },
      { prst: 'textChevron', lines: ['Chevron'], color: '2E75B6', anchor: 'b' },
    ]),
    grid([
      { prst: 'textCanUp', lines: ['Can up'], color: '1F4E79' },
      { prst: 'textPlain', lines: ['Plain text in a box'], color: 'C00000', sz: 2400, wrap: 'square' },
      { prst: 'textDeflateInflate', lines: ['Upper', 'Lower'], color: '385723' },
      { prst: 'textCurveUp', lines: ['Curve up'], color: '7030A0', av: gd('adj', 30000) },
      { prst: 'textFadeRight', lines: ['Fade right'], color: 'ED7D31', algn: 'l' },
      { prst: 'textStop', lines: ['Stop sign'], color: '2E75B6' },
    ]),
  ],
};
