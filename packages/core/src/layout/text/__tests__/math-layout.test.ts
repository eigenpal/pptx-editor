// The math layout (Office Math, ECMA-376 Part 1, 22.1) and the linear form of equations.

import { describe, expect, test } from 'bun:test';
import { parseXmlPart, type Element } from '../../../store/index.ts';
import { isSupportedMath, readMathZone } from '../../../cascade/math.ts';
import type { RunProps } from '../../../cascade/index.ts';
import { layoutMath } from '../math-layout.ts';
import { NORMAL_TEXT_LEADING } from '../math-line-extent.ts';
import { mathAlphanumeric, plainLetter } from '../math-glyphs.ts';
import { createFixedMeasurer } from '../measurer.ts';
import { oMathParaOf, parseLinear } from '../../../automation/commands/equation-linear.ts';
import { linearOfMath } from '../../../automation/commands/equation-text.ts';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const A14 = 'http://schemas.microsoft.com/office/drawing/2010/main';

const props: RunProps = {
  size: 1800,
  bold: false,
  italic: false,
  underline: 'none',
  strike: 'noStrike',
  caps: 'none',
  spacing: 0,
  baseline: 0,
  latin: 'Cambria Math',
  ea: '',
  cs: '',
  fill: { type: 'solid', color: { r: 0, g: 0, b: 0, a: 1 } },
};

/** An `a14:m` zone of the linear form `linear`. */
function zoneOf(linear: string): Element {
  const xml = `<a14:m xmlns:a14="${A14}" xmlns:a="${A}">${oMathParaOf(linear)}</a14:m>`;
  const r = parseXmlPart('/z.xml', xml);
  if (!r.ok) throw new Error('parse');
  return r.part.root;
}

const measurer = createFixedMeasurer();
const box = (linear: string) =>
  layoutMath(
    readMathZone(zoneOf(linear), () => props)!,
    measurer,
    1
  );
const EM = 18 * 12_700;

describe('math layout', () => {
  test('a superscript is smaller and raised, beside its base', () => {
    const b = box('x^2');
    expect(b.glyphs.map((g) => [g.text, g.sizePt])).toEqual([
      ['𝑥', 18],
      ['2', 13],
    ]);
    const [base, sup] = b.glyphs;
    expect(sup!.y).toBeGreaterThan(0.3 * EM);
    expect(sup!.x[0]).toBeCloseTo(base!.x[1]!, 0);
  });

  test('normal text keeps its size in a script, and its ink and leading make the line taller', () => {
    const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
    const run = (t: string, nor: boolean) =>
      `<m:r>${nor ? '<m:rPr><m:nor/></m:rPr>' : ''}<a:rPr/><m:t>${t}</m:t></m:r>`;
    const xml =
      `<a14:m xmlns:a14="${A14}" xmlns:a="${A}"><m:oMathPara xmlns:m="${M}"><m:oMath>` +
      `<m:sSup><m:e>${run('c', true)}</m:e><m:sup>${run('2', true)}</m:sup></m:sSup>` +
      `</m:oMath></m:oMathPara></a14:m>`;
    const r = parseXmlPart('/z.xml', xml);
    if (!r.ok) throw new Error('parse');
    const b = layoutMath(
      readMathZone(r.part.root, () => props)!,
      measurer,
      1
    );
    expect(b.glyphs.map((g) => [g.text, g.sizePt])).toEqual([
      ['c', 18],
      ['2', 18],
    ]);
    const sup = b.glyphs[1]!;
    expect(sup.y).toBeGreaterThan(0.3 * EM);
    // The line holds the ink of the raised digit and the leading of normal text above it.
    expect(b.lineAscent).toBeCloseTo(b.ascent + NORMAL_TEXT_LEADING * EM, 0);
  });

  /** The layout of `m:oMath` content `inner`, in OMML. */
  const omml = (inner: string) => {
    const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
    const xml =
      `<a14:m xmlns:a14="${A14}" xmlns:a="${A}"><m:oMathPara xmlns:m="${M}"><m:oMath>` +
      `${inner}</m:oMath></m:oMathPara></a14:m>`;
    const r = parseXmlPart('/z.xml', xml);
    if (!r.ok) throw new Error('parse');
    return layoutMath(
      readMathZone(r.part.root, () => props)!,
      measurer,
      1
    );
  };
  const mr = (t: string, nor = false) =>
    `<m:r>${nor ? '<m:rPr><m:nor/></m:rPr>' : ''}<a:rPr/><m:t>${t}</m:t></m:r>`;
  const sizes = (b: ReturnType<typeof omml>) => b.glyphs.map((g) => [g.text, g.sizePt]);

  test('normal text keeps its size in the scripts of a math base, also nested', () => {
    // v with a normal-text subscript, x with a normal-text superscript.
    const sub = omml(
      `<m:sSub><m:e>${mr('v')}</m:e><m:sub>${mr('max', true)}</m:sub></m:sSub>` +
        `<m:sSup><m:e>${mr('x')}</m:e><m:sup>${mr('n', true)}</m:sup></m:sSup>`
    );
    expect(sizes(sub)).toEqual([
      ['𝑣', 18],
      ['max', 18],
      ['𝑥', 18],
      ['n', 18],
    ]);
    // x^(y^z): math text shrinks at each level; normal text at the second level does not.
    const nested = (z: string) =>
      omml(
        `<m:sSup><m:e>${mr('x')}</m:e><m:sup><m:sSup><m:e>${mr('y')}</m:e>` +
          `<m:sup>${z}</m:sup></m:sSup></m:sup></m:sSup>`
      );
    expect(sizes(nested(mr('z')))).toEqual([
      ['𝑥', 18],
      ['𝑦', 13],
      ['𝑧', 11],
    ]);
    expect(sizes(nested(mr('z', true)))).toEqual([
      ['𝑥', 18],
      ['𝑦', 13],
      ['z', 18],
    ]);
  });

  test('a normal-text radical degree keeps its size', () => {
    const b = omml(`<m:rad><m:deg>${mr('3', true)}</m:deg><m:e>${mr('x')}</m:e></m:rad>`);
    expect(sizes(b)).toEqual([
      ['3', 18],
      ['𝑥', 18],
    ]);
  });

  test('the line holds the ink below the baseline, not the face descent', () => {
    // A subscript lowers its ink; the line descent is that ink, and no more.
    const b = omml(`<m:sSub><m:e>${mr('x')}</m:e><m:sub>${mr('i')}</m:sub></m:sSub>`);
    const sub = b.glyphs[1]!;
    expect(sub.y).toBeLessThan(0);
    expect(b.lineDescent).toBeCloseTo(b.descent, 0);
    expect(b.lineDescent).toBeLessThan(sub.face.metrics.descent * sub.sizePt * 12_700 - sub.y);
  });

  test('a fraction stacks its parts around one rule, centered', () => {
    const b = box('a/(b+c)');
    expect(b.marks).toHaveLength(1);
    const [num, ...den] = b.glyphs;
    expect(num!.y).toBeGreaterThan(0);
    for (const g of den) expect(g.y).toBeLessThan(0);
    expect(b.ascent).toBeGreaterThan(0.7 * EM);
  });

  test('binary operators and relations take medium and thick spaces', () => {
    const b = box('a+b=c');
    const x = b.glyphs.map((g) => g.x[0]!);
    // 0.5 em glyphs in the fixed measurer, then 4/18 and 5/18 em spaces.
    expect(x[1]! - x[0]!).toBeCloseTo(0.5 * EM + (4 / 18) * EM, 0);
    expect(x[3]! - x[2]!).toBeCloseTo(0.5 * EM + (5 / 18) * EM, 0);
    // A leading minus is a sign, with no space.
    const n = box('-b');
    expect(n.glyphs[0]!.text).toBe('−');
    expect(n.glyphs[1]!.x[0]! - n.glyphs[0]!.x[0]!).toBeCloseTo(0.5 * EM, 0);
  });

  test('a radical draws its sign and its rule over the content', () => {
    const b = box('\\sqrt[3](x)');
    expect(b.marks).toHaveLength(2);
    expect(b.glyphs.map((g) => [g.text, g.sizePt])).toEqual([
      ['3', 11],
      ['𝑥', 18],
    ]);
  });

  test('script glyphs advance by their script-style forms at the unrounded script size', () => {
    // A measurer whose script-style forms are 0.6 em (level 1) and 0.7 em (level 2) wide.
    const fixed = createFixedMeasurer();
    const styled = {
      ...fixed,
      resolve: (r: Parameters<typeof fixed.resolve>[0]) => ({
        ...fixed.resolve(r),
        hasMathTable: true as const,
      }),
      advances: (
        face: Parameters<typeof fixed.advances>[0],
        text: string,
        options?: { readonly mathScript?: 1 | 2 }
      ) => {
        const out = fixed.advances(face, text);
        const w = options?.mathScript === 2 ? 0.7 : options?.mathScript === 1 ? 0.6 : 0;
        if (w) for (let i = 0; i < out.length; i += 1) if (out[i]! > 0) out[i] = w;
        return out;
      },
    };
    const zone = readMathZone(zoneOf('x^(yz^w)'), () => props)!;
    const b = layoutMath(zone, styled, 1);
    const at = (t: string) => b.glyphs.find((g) => g.text === t)!;
    // Drawn at 13 pt and 11 pt (73 % and 60 % of 18 pt, to the nearest half point), placed
    // at 13.14 pt and 10.8 pt.
    expect(at('𝑦').sizePt).toBe(13);
    expect(at('𝑤').sizePt).toBe(11);
    expect(at('𝑧').x[0]! - at('𝑦').x[0]!).toBeCloseTo(0.6 * 13.14 * 12_700, 0);
    const w = at('𝑤');
    expect(w.x[w.x.length - 1]! - w.x[0]!).toBeCloseTo(0.7 * 10.8 * 12_700, 0);
    // The base keeps its plain advance.
    expect(at('𝑥').x[1]! - at('𝑥').x[0]!).toBeCloseTo(0.5 * EM, 0);
  });

  test('a face with no MATH table advances script glyphs at the size they draw', () => {
    // The fixed measurer's faces have no MATH table (a serif fallback, for example).
    const fixed = createFixedMeasurer();
    let asked = false;
    const plain = {
      ...fixed,
      advances: (
        face: Parameters<typeof fixed.advances>[0],
        text: string,
        options?: { readonly mathScript?: 1 | 2 }
      ) => {
        if (options?.mathScript) asked = true;
        return fixed.advances(face, text);
      },
    };
    const b = layoutMath(
      readMathZone(zoneOf('x^(yz)'), () => props)!,
      plain,
      1
    );
    const y = b.glyphs.find((g) => g.text === '𝑦')!;
    expect(y.sizePt).toBe(13);
    expect(y.x[1]! - y.x[0]!).toBeCloseTo(0.5 * 13 * 12_700, 0);
    expect(asked).toBe(false);
  });

  test('a script ends with the space after a script, and prescripts start with it', () => {
    const sup = box('x^2');
    const two = sup.glyphs[1]!;
    expect(sup.width - two.x[two.x.length - 1]!).toBeCloseTo(0.04 * EM, 0);
    const pre = omml(
      `<m:sPre><m:sub>${mr('6')}</m:sub><m:sup>${mr('14')}</m:sup><m:e>${mr('C')}</m:e></m:sPre>`
    );
    const first = Math.min(...pre.glyphs.map((g) => g.x[0]!));
    expect(first).toBeCloseTo(0.04 * EM, 0);
  });

  test('a radical ends with its radicand, and its degree sits on the kerns of the font', () => {
    const plain = box('\\sqrt(x)');
    const x = plain.glyphs[0]!;
    // The rule and the box end where the radicand ends.
    expect(plain.width).toBeCloseTo(x.x[1]!, 0);
    const rule = plain.marks[1]!.commands.filter((c) => c.op !== 'Z');
    expect(Math.max(...rule.map((c) => ('x' in c ? c.x : 0)))).toBeCloseTo(plain.width, 0);
    // The degree starts 0.065 em into the box; the sign starts 0.335 em before its end.
    const deg = box('\\sqrt[3](x)');
    const three = deg.glyphs[0]!;
    expect(three.x[0]!).toBeCloseTo(0.065 * EM, 0);
    const sign = deg.marks[0]!.commands.filter((c) => c.op !== 'Z');
    const left = Math.min(...sign.map((c) => ('x' in c ? c.x : Infinity)));
    expect(left).toBeCloseTo(three.x[1]! - 0.335 * EM + 0.019 * EM, 0);
  });

  test('a radical sign reaches below its box, and its line keeps the box', () => {
    // A subscript gives the radicand a descent; the sign reaches further down than the box.
    const b = omml(
      `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>` +
        `<m:sSub><m:e>${mr('a')}</m:e><m:sub>${mr('n')}</m:sub></m:sSub></m:e></m:rad>`
    );
    const sign = b.marks[0]!;
    const bottom = Math.min(...sign.commands.map((c) => ('y' in c ? c.y : 0)));
    expect(sign.claimBelow).toBeDefined();
    expect(-bottom).toBeGreaterThan(sign.claimBelow!);
    expect(b.descent).toBeCloseTo(sign.claimBelow!, 0);
    expect(b.lineDescent).toBeLessThan(-bottom);
  });

  test('delimiters grow with a matrix', () => {
    const small = box('(x)');
    const big = box('(\\matrix(a&b@c&d))');
    const height = (b: ReturnType<typeof box>) => b.ascent + b.descent;
    expect(height(big)).toBeGreaterThan(height(small) * 1.5);
    expect(big.marks).toHaveLength(2);
  });

  test('letters draw as plain italic letters in a face with no math alphanumerics', () => {
    // A host with no math font: every math family is substituted.
    const fixed = createFixedMeasurer();
    const noMath = {
      ...fixed,
      resolve: (r: Parameters<typeof fixed.resolve>[0]) => ({
        ...fixed.resolve(r),
        substituted: /math/i.test(r.family),
      }),
    };
    const zone = readMathZone(zoneOf('x+𝑦'), () => props)!;
    const b = layoutMath(zone, noMath, 1);
    expect(b.glyphs[2]!.text).toBe('y');
    expect(b.glyphs[2]!.face.italic).toBe(true);
    expect(b.glyphs[0]!.text).toBe('x');
    expect(b.glyphs[0]!.face.italic).toBe(true);
  });
});

describe('math alphanumerics', () => {
  test('map letters to their math forms and back', () => {
    expect(String.fromCodePoint(mathAlphanumeric(0x78, 'i'))).toBe('𝑥');
    expect(String.fromCodePoint(mathAlphanumeric(0x68, 'i'))).toBe('ℎ');
    expect(String.fromCodePoint(mathAlphanumeric(0x3b1, 'i'))).toBe('𝛼');
    expect(mathAlphanumeric(0x393, 'i')).toBe(0x393);
    expect(plainLetter('𝑥'.codePointAt(0)!)).toEqual({ cp: 0x78, style: 'i' });
    expect(plainLetter('𝐀'.codePointAt(0)!)).toEqual({ cp: 0x41, style: 'b' });
    expect(plainLetter('𝛼'.codePointAt(0)!)).toEqual({ cp: 0x3b1, style: 'i' });
    expect(plainLetter(0x78)).toBeUndefined();
  });
});

describe('the linear form', () => {
  const cases = [
    'E=mc^2',
    'x=(-b±\\sqrt(b^2-4ac))/(2a)',
    '\\sum_(k=1)^n k',
    'lim_(n→∞) (1+1/n)^n=e',
    '\\matrix(a&b@c&d)',
    'x_i^2+\\sqrt[3](8)',
    '\\hat(x)+\\bar(y)',
    'sin θ',
  ];
  for (const linear of cases)
    test(`${linear} reads back as itself`, () => {
      const zone = zoneOf(linear);
      expect(isSupportedMath(zone)).toBe(true);
      const back = linearOfMath(zone)!;
      expect(back).toBeDefined();
      // The linear form of the markup writes the same markup.
      expect(oMathParaOf(back)).toBe(oMathParaOf(linear));
    });

  test('refuses what the syntax does not cover', () => {
    for (const bad of ['', 'x^', '(a', 'a)', '\\nope', '\\matrix(a&b', 'a'.repeat(1001)])
      expect(() => parseLinear(bad)).toThrow();
  });

  test('names Greek letters and symbols', () => {
    expect(oMathParaOf('\\alpha\\le\\beta')).toContain('<m:t>α</m:t>');
    expect(oMathParaOf('\\alpha\\le\\beta')).toContain('<m:t>≤</m:t>');
  });

  test('an equation of other constructs has no linear form', () => {
    const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
    const xml = `<a14:m xmlns:a14="${A14}"><m:oMath xmlns:m="${M}"><m:groupChr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:groupChr></m:oMath></a14:m>`;
    const r = parseXmlPart('/z.xml', xml);
    if (!r.ok) throw new Error('parse');
    expect(isSupportedMath(r.part.root)).toBe(true);
    expect(linearOfMath(r.part.root)).toBeUndefined();
  });
});
