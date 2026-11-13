// The `underline-styles` fixture: every underline type (`ST_TextUnderlineType`, ECMA-376 Part 1,
// 20.1.10.82) at 18 pt and 36 pt; underline line properties (`a:uLn`, `a:uLnTx`) with widths,
// dashes, and fills; underline fills (`a:uFill`, `a:uFillTx`); and underlined runs of two sizes
// on one line. The `caps-scripts` fixture: small caps (`cap="small"`) and superscript and
// subscript (`baseline`) at several sizes and offsets.

import { para, run, shape, solid, type FixtureSlide } from './builder.ts';

const IN = 914_400;

/** Every underline type but `none`, in schema order. */
export const UNDERLINE_TYPES = [
  'words', 'sng', 'dbl', 'heavy', 'dotted', 'dottedHeavy', 'dash', 'dashHeavy', 'dashLong',
  'dashLongHeavy', 'dotDash', 'dotDashHeavy', 'dotDotDash', 'dotDotDashHeavy', 'wavy',
  'wavyHeavy', 'wavyDbl',
] as const; // prettier-ignore

const BODY = '<a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720"><a:noAutofit/></a:bodyPr>';
const SPACED = '<a:pPr><a:spcBef><a:spcPts val="600"/></a:spcBef></a:pPr>';

const box = (id: number, x: number, y: number, cx: number, cy: number, text: string, fill: string): string =>
  shape({ id, x, y, cx, cy, text, bodyPr: BODY, fill: solid(fill) });

const red = '<a:solidFill><a:srgbClr val="C00000"/></a:solidFill>';
const blue = '<a:solidFill><a:srgbClr val="0070C0"/></a:solidFill>';
const green = '<a:solidFill><a:srgbClr val="00B050"/></a:solidFill>';
const uLn = (w: number, inner = ''): string => `<a:uLn w="${w}">${inner}</a:uLn>`;

/** Every underline type at 18 pt, in two columns. */
function smallSlide(): FixtureSlide {
  const line = (u: string): string => para(run(`${u} under two words`, `u="${u}"`), SPACED);
  const half = Math.ceil(UNDERLINE_TYPES.length / 2);
  return {
    shapes:
      box(2, 0.4 * IN, 0.3 * IN, 6.1 * IN, 6.9 * IN, UNDERLINE_TYPES.slice(0, half).map(line).join(''), 'EEF3FA') +
      box(3, 6.8 * IN, 0.3 * IN, 6.1 * IN, 6.9 * IN, UNDERLINE_TYPES.slice(half).map(line).join(''), 'FDF2E9'),
  };
}

/** Every underline type at 36 pt, in two columns. */
function largeSlide(): FixtureSlide {
  const line = (u: string): string => para(run(`${u} a b`, `sz="3600" u="${u}"`), SPACED);
  const half = Math.ceil(UNDERLINE_TYPES.length / 2);
  return {
    shapes:
      box(2, 0.3 * IN, 0.2 * IN, 6.3 * IN, 7.1 * IN, UNDERLINE_TYPES.slice(0, half).map(line).join(''), 'EEF3FA') +
      box(3, 6.8 * IN, 0.2 * IN, 6.3 * IN, 7.1 * IN, UNDERLINE_TYPES.slice(half).map(line).join(''), 'FDF2E9'),
  };
}

/** Underline line properties and fills. */
function lineSlide(): FixtureSlide {
  const row = (text: string, attrs: string, children: string): string =>
    para(run(text, attrs, children), SPACED);
  return {
    shapes:
      box(2, 0.4 * IN, 0.3 * IN, 6.1 * IN, 6.9 * IN,
        row('Single, line 1 pt', 'u="sng"', uLn(12700)) +
        row('Single, line 3 pt', 'u="sng"', uLn(38100)) +
        row('Single, line 6 pt', 'u="sng"', uLn(76200)) +
        row('Single, dashed line', 'u="sng"', uLn(25400, '<a:prstDash val="dash"/>')) +
        row('Single, round dots', 'u="sng"', uLn(38100, '<a:prstDash val="sysDot"/>').replace('<a:uLn ', '<a:uLn cap="rnd" ')) +
        row('Single, long dash dot', 'u="sng"', uLn(25400, '<a:prstDash val="lgDashDot"/>')) +
        row('Double, line 2 pt', 'u="dbl"', uLn(25400)) +
        row('Heavy, line 1 pt', 'u="heavy"', uLn(12700)) +
        row('Dotted, line 3 pt', 'u="dotted"', uLn(38100)) +
        row('Wavy, line 2 pt', 'u="wavy"', uLn(25400)),
        'EEF3FA') +
      box(3, 6.8 * IN, 0.3 * IN, 6.1 * IN, 6.9 * IN,
        row('Line fill red, no uFill', 'u="sng"', uLn(25400, red)) +
        row('Line follows text', 'u="sng"', `${blue}<a:uLnTx/>`) +
        row('Fill follows text', 'u="sng"', `${blue}<a:uFillTx/>`) +
        row('Green fill, line 3 pt', 'u="sng"', `${uLn(38100)}<a:uFill>${green}</a:uFill>`) +
        row('Green fill, dashed', 'u="dash"', `<a:uFill>${green}</a:uFill>`) +
        row('Single 36 pt, line 1 pt', 'sz="3600" u="sng"', uLn(12700)) +
        row('Wavy 36 pt, line 1 pt', 'sz="3600" u="wavy"', uLn(12700)) +
        para(run('Large 36 and ', 'sz="3600" u="sng"') + run('small 14 underlined', 'sz="1400" u="sng"'), SPACED) +
        para(run('Small 14 and ', 'sz="1400" u="dbl"') + run('large 36 double', 'sz="3600" u="dbl"'), SPACED),
        'FDF2E9'),
  };
}

const SIZES = [1000, 1400, 1800, 2400, 3200, 4400, 6000] as const;
const OFFSETS = [10000, 30000, 50000, 80000, -10000, -25000, -50000] as const;

/** Small caps and superscript and subscript at several sizes, and scripts at several offsets. */
function capsScriptsSlide(): FixtureSlide {
  return {
    shapes:
      box(2, 0.2 * IN, 0.3 * IN, 4.3 * IN, 6.9 * IN,
        SIZES.map((sz) => para(run('Small Caps Hxhx', `sz="${sz}" cap="small"`))).join(''),
        'EEF3FA') +
      box(3, 4.6 * IN, 0.3 * IN, 4.4 * IN, 6.9 * IN,
        SIZES.map((sz) =>
          para(run('H', `sz="${sz}"`) + run('Hsup', `sz="${sz}" baseline="30000"`) +
            run(' H', `sz="${sz}"`) + run('Hsub', `sz="${sz}" baseline="-25000"`))).join(''),
        'FDF2E9') +
      box(4, 9.1 * IN, 0.3 * IN, 4.0 * IN, 6.9 * IN,
        OFFSETS.map((b) =>
          para(run(`Offset ${b / 1000} H`) + run('Hxh', `baseline="${b}"`) + run('H'), SPACED)).join('') +
        para(run('Caps and script H') + run('Hxh', 'baseline="30000" cap="small"'), SPACED),
        'E8F5E9'),
  };
}

export const underlineStylesFixtures: Record<string, FixtureSlide[]> = {
  'underline-styles': [smallSlide(), largeSlide(), lineSlide()],
  'caps-scripts': [capsScriptsSlide()],
};
