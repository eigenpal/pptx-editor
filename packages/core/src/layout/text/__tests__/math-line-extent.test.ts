// The ink of math glyphs, the extent an equation claims on its line, and script shifts.

import { describe, expect, test } from 'bun:test';
import type { RunProps } from '../../../cascade/index.ts';
import { readMathZone } from '../../../cascade/math.ts';
import { oMathParaOf } from '../../../automation/commands/equation-linear.ts';
import { parseXmlPart } from '../../../store/index.ts';
import { layoutMath } from '../math-layout.ts';
import { isInvisibleMathText, mathInkOf, mathTextInk } from '../math-ink.ts';
import { MATH_LEADING, mathLineExtent } from '../math-line-extent.ts';
import { scriptShifts } from '../math-scripts.ts';
import { createFixedMeasurer, type ResolvedFace, type TextMeasurer } from '../measurer.ts';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const A14 = 'http://schemas.microsoft.com/office/drawing/2010/main';
const EMU_PER_PT = 12_700;
const EM = 18 * EMU_PER_PT;

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

function layout(linear: string, measurer: TextMeasurer) {
  const xml = `<a14:m xmlns:a14="${A14}" xmlns:a="${A}">${oMathParaOf(linear)}</a14:m>`;
  const r = parseXmlPart('/z.xml', xml);
  if (!r.ok) throw new Error('parse');
  return layoutMath(
    readMathZone(r.part.root, () => props)!,
    measurer,
    1
  );
}

/** The fixed measurer with no outline for `cp`, drawn by a fallback face with a tall box. */
function withoutOutline(cp: number): { measurer: TextMeasurer; fallback: ResolvedFace } {
  const fixed = createFixedMeasurer();
  const own = fixed.resolve({ family: 'Fallback', bold: false, italic: false });
  const fallback: ResolvedFace = {
    ...own,
    id: 'fallback',
    metrics: { ...own.metrics, ascent: 1.2, descent: 0.4 },
  };
  const measurer: TextMeasurer = {
    ...fixed,
    fallbackFace: (_face, c) => (c === cp ? fallback : undefined),
    glyphOutline: (face, c) => (c === cp ? undefined : fixed.glyphOutline!(face, c)),
  };
  return { measurer, fallback };
}

describe('math ink', () => {
  test('reads the outline of a glyph, and gives spaces no ink', () => {
    const measurer = createFixedMeasurer();
    const face = measurer.resolve({ family: 'Cambria Math', bold: false, italic: false });
    expect(mathInkOf(measurer, face, 0x78)).toEqual({ lo: 0, hi: 0.7 });
    expect(mathInkOf(measurer, face, 32)).toEqual({ lo: 0, hi: 0 });
    expect(isInvisibleMathText(' ⁡')).toBe(true);
    expect(mathTextInk(measurer, face, ' x ')).toEqual({ lo: 0, hi: 0.7 });
  });

  test('a glyph with no outline gives the box of the face that draws it', () => {
    const { measurer } = withoutOutline(0x4e2d);
    const face = measurer.resolve({ family: 'Cambria Math', bold: false, italic: false });
    expect(mathInkOf(measurer, face, 0x4e2d)).toEqual({ lo: -0.4, hi: 1.2, fromFontBox: true });
    expect(mathTextInk(measurer, face, 'x中').fromFontBox).toBe(true);
  });

  test('faces that share an id keep their own boxes for a glyph with no outline', () => {
    const fixed = createFixedMeasurer();
    const measurer: TextMeasurer = { ...fixed, glyphOutline: () => undefined };
    const face = fixed.resolve({ family: 'Shared', bold: false, italic: false });
    const tall: ResolvedFace = { ...face, metrics: { ...face.metrics, ascent: 1.5 } };
    expect(mathInkOf(measurer, face, 0x78).hi).toBe(face.metrics.ascent);
    expect(mathInkOf(measurer, tall, 0x78).hi).toBe(1.5);
  });
});

describe('math line extent', () => {
  test('a glyph claims its ink and the math leading of the base size above it', () => {
    const measurer = createFixedMeasurer();
    const b = layout('x^2', measurer);
    const sup = b.glyphs[1]!;
    // The ink of the raised digit, 0.7 em of its size above its baseline, and 0.16 em of the
    // base size: more than the font box of the digit.
    const top = sup.y + 0.7 * sup.sizePt * EMU_PER_PT + MATH_LEADING * EM;
    expect(b.lineAscent).toBeCloseTo(top, 0);
    expect(b.lineAscent).toBeGreaterThan(sup.y + sup.face.metrics.ascent * sup.sizePt * EMU_PER_PT);
    expect(b.lineDescent).toBe(0);
  });

  test('normal text claims less leading than math text', () => {
    const measurer = createFixedMeasurer();
    const face = measurer.resolve({ family: 'Cambria Math', bold: false, italic: false });
    const glyph = { text: 'x', x: [0, 1], y: 0, face, sizePt: 18, props, level: 0 as const };
    const math = mathLineExtent([glyph], [], 18, measurer).ascent;
    const normal = mathLineExtent([{ ...glyph, normal: true }], [], 18, measurer).ascent;
    expect(math).toBeCloseTo(0.7 * EM + 0.16 * EM, 0);
    expect(normal).toBeCloseTo(0.7 * EM + 0.14 * EM, 0);
  });

  test('a glyph with no outline claims the box of the face that draws it', () => {
    const { measurer, fallback } = withoutOutline(0x4e2d);
    const b = layout('x+中', measurer);
    expect(b.lineAscent).toBeCloseTo(fallback.metrics.ascent * EM, 0);
    expect(b.lineDescent).toBeCloseTo(fallback.metrics.descent * EM, 0);
  });

  test('text that draws nothing claims no height, and marks claim their outline', () => {
    const measurer = createFixedMeasurer();
    const face = measurer.resolve({ family: 'Cambria Math', bold: false, italic: false });
    const glyph = { x: [0, 1], y: 0, face, sizePt: 18, props, level: 0 as const };
    const rule = {
      props,
      commands: [
        { op: 'M' as const, x: 0, y: -2 * EM },
        { op: 'L' as const, x: 1, y: 3 * EM },
        { op: 'Z' as const },
      ],
    };
    expect(mathLineExtent([{ ...glyph, text: ' ' }], [], 18, measurer)).toEqual({
      ascent: 0,
      descent: 0,
    });
    expect(mathLineExtent([], [rule], 18, measurer)).toEqual({ ascent: 3 * EM, descent: 2 * EM });
    // A mark that claims room above its outline: the rule of a radical.
    expect(mathLineExtent([], [{ ...rule, claimAbove: EM }], 18, measurer).ascent).toBe(4 * EM);
  });

  test('a radical claims the math leading above its rule', () => {
    const measurer = createFixedMeasurer();
    const b = layout('\\sqrt(x)', measurer);
    let top = -Infinity;
    for (const m of b.marks) for (const c of m.commands) if (c.op !== 'Z') top = Math.max(top, c.y);
    expect(b.lineAscent - top).toBeCloseTo(MATH_LEADING * EM, 0);
  });
});

describe('script shifts', () => {
  const part = (asc: number, desc = 0) => ({ asc: asc * EM, desc: desc * EM });

  test('a superscript rises at least 0.36 em, and over a tall base by its ink', () => {
    expect(scriptShifts(part(0.46), part(0.5), undefined, EM, false).u / EM).toBeCloseTo(0.36);
    expect(scriptShifts(part(0.46), part(0.5), undefined, EM, true).u / EM).toBeCloseTo(0.293);
    // A base 0.7 em tall, as a letter with an ascender, gives a cramped superscript 0.3 em.
    expect(scriptShifts(part(0.7), part(0.5), undefined, EM, true).u / EM).toBeCloseTo(0.3);
    expect(scriptShifts(part(0.9), part(0.5), undefined, EM, false).u / EM).toBeCloseTo(0.5);
  });

  test('a subscript drops 0.2 em, and lower when its ink is tall', () => {
    expect(scriptShifts(part(0.5), undefined, part(0.3), EM, false).v / EM).toBeCloseTo(0.2);
    expect(scriptShifts(part(0.5), undefined, part(0.7), EM, false).v / EM).toBeCloseTo(
      0.7 - 760 / 2048
    );
    // A font box that stands in for missing ink does not lower the subscript.
    const tofu = { ...part(1.2), fromFontBox: true as const };
    expect(scriptShifts(part(0.5), undefined, tofu, EM, false).v / EM).toBeCloseTo(0.2);
  });

  test('scripts on one base keep an ink gap: the superscript rises first, then the subscript drops', () => {
    const { u, v } = scriptShifts(part(0.5), part(0.5, 0.1), part(0.5), EM, false);
    // The ink gap between the bottom of the superscript and the top of the subscript.
    expect((u - 0.1 * EM - (0.5 * EM - v)) / EM).toBeCloseTo(300 / 2048);
    // The superscript bottom rises to its limit, 765/2048 em.
    expect((u - 0.1 * EM) / EM).toBeCloseTo(765 / 2048);
  });
});
