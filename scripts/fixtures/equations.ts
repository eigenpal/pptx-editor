// The equations fixture: Office Math (`m:oMathPara`, ECMA-376 Part 1, 22.1) in text boxes, in the
// markup-compatibility wrapper that a reader that understands `a14` selects (`mc:Choice
// Requires="a14"`). Every math run names STIX Two Math, an open font with an OpenType MATH table,
// and the fallback of each equation is the same text box with plain text. One equation for each
// shape: runs and scripts, fractions and radicals, delimiters and limits, matrices, bars, boxes,
// group characters, and equation arrays. N-ary operators, accents, and skewed and linear fractions
// are not here: the reference draws their glyphs with a math font that is not open.

import { escapeText, type FixtureSlide } from './builder.ts';

const IN = 914_400;
const A14 = 'http://schemas.microsoft.com/office/drawing/2010/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const MATH = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const MATH_FONT = '<a:latin typeface="STIX Two Math"/><a:ea typeface="STIX Two Math"/><a:cs typeface="STIX Two Math"/>';
const TEXT_FONT = '<a:latin typeface="Liberation Sans"/><a:ea typeface="Liberation Sans"/><a:cs typeface="Liberation Sans"/>';

/** A math run. Letters draw in math italic, digits and operators upright. */
export const r = (t: string, mrPr = ''): string =>
  `<m:r>${mrPr ? `<m:rPr>${mrPr}</m:rPr>` : ''}<a:rPr lang="en-US" i="1" dirty="0">${MATH_FONT}</a:rPr><m:t>${escapeText(t)}</m:t></m:r>`;
/** A math run in upright, plain style (`m:sty m:val="p"`). */
export const up = (t: string): string => r(t, '<m:sty m:val="p"/>');
const e = (body: string, tag = 'e'): string => `<m:${tag}>${body}</m:${tag}>`;
const chr = (c: string): string => `<m:chr m:val="${escapeText(c)}"/>`;

export const sup = (base: string, s: string): string => `<m:sSup>${e(base)}${e(s, 'sup')}</m:sSup>`;
export const sub = (base: string, s: string): string => `<m:sSub>${e(base)}${e(s, 'sub')}</m:sSub>`;
export const subSup = (base: string, b: string, t: string): string =>
  `<m:sSubSup>${e(base)}${e(b, 'sub')}${e(t, 'sup')}</m:sSubSup>`;
export const pre = (b: string, t: string, base: string): string =>
  `<m:sPre>${e(b, 'sub')}${e(t, 'sup')}${e(base)}</m:sPre>`;
export const frac = (n: string, d: string, type?: string): string =>
  `<m:f>${type ? `<m:fPr><m:type m:val="${type}"/></m:fPr>` : ''}${e(n, 'num')}${e(d, 'den')}</m:f>`;
export const rad = (body: string, deg?: string): string =>
  deg === undefined
    ? `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>${e(body)}</m:rad>`
    : `<m:rad>${e(deg, 'deg')}${e(body)}</m:rad>`;
export const nary = (c: string, lo: string, hi: string, body: string, limLoc?: string): string =>
  `<m:nary><m:naryPr>${chr(c)}${limLoc ? `<m:limLoc m:val="${limLoc}"/>` : ''}</m:naryPr>${e(lo, 'sub')}${e(hi, 'sup')}${e(body)}</m:nary>`;
export const delim = (items: readonly string[], beg?: string, end?: string, sep?: string): string => {
  const props =
    (beg !== undefined ? `<m:begChr m:val="${escapeText(beg)}"/>` : '') +
    (sep !== undefined ? `<m:sepChr m:val="${escapeText(sep)}"/>` : '') +
    (end !== undefined ? `<m:endChr m:val="${escapeText(end)}"/>` : '');
  return `<m:d>${props ? `<m:dPr>${props}</m:dPr>` : ''}${items.map((i) => e(i)).join('')}</m:d>`;
};
export const func = (name: string, body: string): string => `<m:func>${e(name, 'fName')}${e(body)}</m:func>`;
export const limLow = (body: string, lim: string): string => `<m:limLow>${e(body)}${e(lim, 'lim')}</m:limLow>`;
export const limUpp = (body: string, lim: string): string => `<m:limUpp>${e(body)}${e(lim, 'lim')}</m:limUpp>`;
export const matrix = (rows: readonly (readonly string[])[]): string =>
  `<m:m>${rows.map((row) => `<m:mr>${row.map((c) => e(c)).join('')}</m:mr>`).join('')}</m:m>`;
export const acc = (c: string, body: string): string => `<m:acc><m:accPr>${chr(c)}</m:accPr>${e(body)}</m:acc>`;
export const bar = (body: string, pos?: string): string =>
  `<m:bar>${pos ? `<m:barPr><m:pos m:val="${pos}"/></m:barPr>` : ''}${e(body)}</m:bar>`;
export const groupChr = (c: string, body: string, pos?: string): string =>
  `<m:groupChr><m:groupChrPr>${chr(c)}${pos ? `<m:pos m:val="${pos}"/><m:vertJc m:val="${pos === 'top' ? 'bot' : 'top'}"/>` : ''}</m:groupChrPr>${e(body)}</m:groupChr>`;
export const box = (body: string): string => `<m:box>${e(body)}</m:box>`;
export const borderBox = (body: string): string => `<m:borderBox>${e(body)}</m:borderBox>`;
export const eqArr = (rows: readonly string[]): string => `<m:eqArr>${rows.map((x) => e(x)).join('')}</m:eqArr>`;

/** One equation in a text box: `math` is the `m:oMath` content, `plain` its fallback text. */
function equation(id: number, col: number, row: number, math: string, plain: string): string {
  const x = Math.round((0.4 + col * 4.25) * IN);
  const y = Math.round((1.2 + row * 2) * IN);
  const nv = `<p:nvSpPr><p:cNvPr id="${id}" name="Equation ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`;
  const spPr =
    `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${Math.round(4 * IN)}" cy="${Math.round(1.6 * IN)}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>' +
    '<a:ln w="6350"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln></p:spPr>';
  const bodyPr = '<a:bodyPr wrap="square" rtlCol="0" anchor="ctr"><a:noAutofit/></a:bodyPr><a:lstStyle/>';
  const end = `<a:endParaRPr lang="en-US" sz="2400" dirty="0">${TEXT_FONT}</a:endParaRPr>`;
  return (
    `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:a14="${A14}" Requires="a14"><p:sp>${nv}${spPr}<p:txBody>${bodyPr}` +
    `<a:p><a14:m><m:oMathPara xmlns:m="${MATH}"><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr>` +
    `<m:oMath>${math}</m:oMath></m:oMathPara></a14:m>${end}</a:p></p:txBody></p:sp></mc:Choice>` +
    `<mc:Fallback><p:sp>${nv}${spPr}<p:txBody>${bodyPr}<a:p><a:pPr algn="ctr"/>` +
    `<a:r><a:rPr lang="en-US" sz="2400" dirty="0">${TEXT_FONT}</a:rPr><a:t>${escapeText(plain)}</a:t></a:r>${end}</a:p>` +
    '</p:txBody></p:sp></mc:Fallback></mc:AlternateContent>'
  );
}

/** A title in an open font. */
function title(text: string): string {
  return (
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
    `<p:spPr><a:xfrm><a:off x="${Math.round(0.4 * IN)}" y="${Math.round(0.3 * IN)}"/><a:ext cx="${Math.round(12.5 * IN)}" cy="${Math.round(0.7 * IN)}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>' +
    '<p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>' +
    `<a:p><a:r><a:rPr lang="en-US" sz="2800" b="1" dirty="0">${TEXT_FONT}</a:rPr><a:t>${escapeText(text)}</a:t></a:r></a:p></p:txBody></p:sp>`
  );
}

function slide(name: string, items: readonly (readonly [string, string])[]): FixtureSlide {
  return {
    shapes: title(name) + items.map(([math, plain], i) => equation(3 + i, i % 3, Math.floor(i / 3), math, plain)).join(''),
  };
}

const SERIF_FONT = '<a:latin typeface="Liberation Serif"/><a:ea typeface="Liberation Serif"/><a:cs typeface="Liberation Serif"/>';
/** A normal-text math run (`m:nor`) in STIX Two Math. */
const nor = (t: string): string => r(t, '<m:nor/>');
/**
 * A normal-text math run (`m:nor`) in Liberation Serif, a face with no MATH table. Math text in
 * such a face is not here: the presentation app draws it in a math font that is not open.
 */
const serif = (t: string): string =>
  `<m:r><m:rPr><m:nor/></m:rPr><a:rPr lang="en-US" i="1" dirty="0">${SERIF_FONT}</a:rPr><m:t>${escapeText(t)}</m:t></m:r>`;

/** One paragraph of a stacked box: math (`m:oMath` content) or plain text. */
type Para = { readonly math: string } | { readonly text: string };

/**
 * A top-anchored text box of several paragraphs, for the height of lines that hold equations:
 * equations stacked on one another, and text lines above and below them. `spacing` is the
 * paragraphs' line spacing in percent.
 */
function stacked(id: number, col: number, row: number, paras: readonly Para[], spacing?: number): string {
  const x = Math.round((0.4 + col * 4.25) * IN);
  const y = Math.round((1.1 + row * 3.1) * IN);
  const nv = `<p:nvSpPr><p:cNvPr id="${id}" name="Stack ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`;
  const spPr =
    `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${Math.round(4 * IN)}" cy="${Math.round(2.9 * IN)}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>' +
    '<a:ln w="6350"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln></p:spPr>';
  const bodyPr = '<a:bodyPr wrap="square" rtlCol="0" anchor="t"><a:noAutofit/></a:bodyPr><a:lstStyle/>';
  const pPr = spacing ? `<a:pPr><a:lnSpc><a:spcPct val="${spacing * 1000}"/></a:lnSpc></a:pPr>` : '';
  const end = `<a:endParaRPr lang="en-US" sz="1800" dirty="0">${TEXT_FONT}</a:endParaRPr>`;
  const run = (t: string) => `<a:r><a:rPr lang="en-US" sz="1800" dirty="0">${TEXT_FONT}</a:rPr><a:t>${escapeText(t)}</a:t></a:r>`;
  const choice = paras
    .map((p) =>
      'math' in p
        ? `<a:p>${pPr}<a14:m><m:oMathPara xmlns:m="${MATH}"><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr>` +
          `<m:oMath>${p.math}</m:oMath></m:oMathPara></a14:m>${end}</a:p>`
        : `<a:p>${pPr}${run(p.text)}${end}</a:p>`
    )
    .join('');
  const fallback = paras.map((p) => `<a:p>${pPr}${'text' in p ? run(p.text) : ''}${end}</a:p>`).join('');
  return (
    `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:a14="${A14}" Requires="a14"><p:sp>${nv}${spPr}<p:txBody>${bodyPr}` +
    `${choice}</p:txBody></p:sp></mc:Choice>` +
    `<mc:Fallback><p:sp>${nv}${spPr}<p:txBody>${bodyPr}${fallback}</p:txBody></p:sp></mc:Fallback></mc:AlternateContent>`
  );
}

function stackSlide(name: string, boxes: readonly (readonly [readonly Para[], number?])[]): FixtureSlide {
  return {
    shapes: title(name) + boxes.map(([paras, spacing], i) => stacked(3 + i, i % 3, Math.floor(i / 3), paras, spacing)).join(''),
  };
}

const fr = frac(r('a'), r('b'));
const emc = (run: (t: string) => string, script: (t: string) => string): string =>
  run('E') + run('=') + run('m') + sup(run('c'), script('2'));

export const equationFixtures: Record<string, FixtureSlide[]> = {
  // Lines that hold equations, in top-anchored boxes: stacked display fractions and scripts,
  // text above and below, and 150 % line spacing. The second slide holds normal text (`m:nor`)
  // in scripts and in a radical degree, also in a face with no MATH table.
  'equation-lines': [
    stackSlide('Equation lines', [
      [[{ math: fr }, { math: fr }]],
      [[{ text: 'Text above gyp' }, { math: fr }, { text: 'Text below Ég' }]],
      [[{ math: sup(r('x'), r('2')) }, { math: sup(r('x'), r('2')) }, { text: 'Text below gyp' }]],
      [[{ math: fr }, { math: fr }], 150],
      [[{ text: 'Text above gyp' }, { math: emc(r, r) }, { text: 'Text below Ég' }]],
      [[{ math: frac(frac(r('1'), r('x')), r('2')) }, { text: 'Text below gyp' }]],
    ]),
    stackSlide('Normal text in scripts', [
      [[{ text: 'Math base, normal script' }, { math: sub(r('v'), nor('max')) + r('+') + sup(r('x'), nor('n')) }]],
      [[{ text: 'Nested scripts' }, { math: sup(r('x'), sup(r('y'), r('z'))) + r('+') + sup(r('x'), sup(r('y'), nor('z'))) }]],
      [[{ text: 'Normal radical degree' }, { math: rad(r('x'), nor('3')) + r('+') + rad(r('x'), r('3')) }]],
      [[{ math: sub(r('x'), nor('i')) + r('+') + sup(r('y'), nor('k')) }, { text: 'Text below gyp' }]],
      [[{ text: 'Serif, normal text' }, { math: emc(serif, serif) }]],
      [[{ text: 'Math font, normal script' }, { math: emc(r, nor) }]],
    ]),
  ],

  equations: [
    slide('Runs and scripts', [
      [r('E') + r('=') + r('m') + sup(r('c'), r('2')), 'E = mc^2'],
      [subSup(r('x'), r('i'), r('2')) + r('+') + sub(r('y'), r('j')), 'x_i^2 + y_j'],
      [pre(r('6'), r('14'), up('C')) + r('→') + pre(r('7'), r('14'), up('N')), '14 6 C → 14 7 N'],
      [r('α') + r('+') + r('β') + r('=') + r('γ'), 'α + β = γ'],
      [func(up('sin'), r('θ')) + r('+') + func(up('cos'), r('θ')), 'sin θ + cos θ'],
      [r('area', '<m:nor/>') + r('=') + r('π') + sup(r('r'), r('2')), 'area = πr^2'],
    ]),
    slide('Fractions and radicals', [
      [r('y') + r('=') + frac(r('a'), r('b')), 'y = a/b'],
      [frac(frac(r('1'), r('x')), r('2')) + r('+') + r('1'), '(1/x)/2 + 1'],
      [delim([frac(r('n'), r('k'), 'noBar')]), '(n k)'],
      [rad(r('x') + r('+') + r('1')), '√(x+1)'],
      [rad(r('8'), r('3')) + r('=') + r('2'), '∛8 = 2'],
      [
        r('x') + r('=') + frac(r('−') + r('b') + r('±') + rad(sup(r('b'), r('2')) + r('−') + r('4') + r('a') + r('c')), r('2') + r('a')),
        'x = (−b ± √(b^2−4ac))/2a',
      ],
    ]),
    slide('Delimiters and limits', [
      [r('s') + r('=') + frac(r('n') + delim([r('n') + r('+') + r('1')]), r('2')), 's = n(n+1)/2'],
      [delim([r('x') + r('−') + r('1')], '|', '|') + r('≤') + r('2'), '|x − 1| ≤ 2'],
      [func(sup(up('sin'), r('2')), r('x')) + r('+') + sup(delim([r('a') + r('+') + r('b')]), r('2')), 'sin^2 x + (a+b)^2'],
      [delim([frac(r('a'), r('b'))], '[', ']'), '[a/b]'],
      [delim([r('x'), r('y')], '{', '}', '|'), '{x | y}'],
      [limLow(up('lim'), r('n') + r('→') + r('∞')) + frac(r('1'), r('n')), 'lim 1/n'],
    ]),
    slide('Matrices, bars, and boxes', [
      [delim([matrix([[r('a'), r('b')], [r('c'), r('d')]])]), '(a b; c d)'],
      [sub(r('a'), r('n') + r('+') + r('1')) + r('=') + rad(sub(r('a'), r('n')) + r('+') + r('1')), 'a_(n+1) = √(a_n + 1)'],
      [bar(r('z')) + r('+') + bar(r('w'), 'bot'), 'z̄ + w̲'],
      [groupChr('⏟', r('a') + r('+') + r('b'), 'bot') + r('+') + limUpp(r('x'), r('def')), 'a + b, x def'],
      [box(r('a') + r('=') + r('b')) + r('+') + borderBox(r('c')), 'a = b, c'],
      [eqArr([r('x') + r('+') + r('y') + r('=') + r('1'), r('x') + r('−') + r('y') + r('=') + r('0')]), 'x + y = 1, x − y = 0'],
    ]),
  ],
};
