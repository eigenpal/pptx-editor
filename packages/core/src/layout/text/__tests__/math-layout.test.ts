// The math layout (Office Math, ECMA-376 Part 1, 22.1) and the linear form of equations.

import { describe, expect, test } from 'bun:test';
import { parseXmlPart, type Element } from '../../../store/index.ts';
import { isSupportedMath, readMathZone } from '../../../cascade/math.ts';
import type { RunProps } from '../../../cascade/index.ts';
import { layoutMath } from '../math-layout.ts';
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
