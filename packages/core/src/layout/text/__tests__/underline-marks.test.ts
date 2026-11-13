import { describe, expect, test } from 'bun:test';
import type { RunProps } from '../../../cascade/index.ts';
import type { GlyphRun } from '../text-layout.ts';
import {
  isDrawnUnderline,
  lineUnderlines,
  underlineRings,
  type UnderlineRing,
} from '../underline-marks.ts';

const PT = 12_700;
const round = (v: number): number => Math.round(v * 1000) / 1000;

/** The [top, bottom] of each ring, in points. */
const bands = (rings: readonly UnderlineRing[]): number[][] =>
  rings.map((r) => [
    round(Math.min(...r.map((p) => p[1])) / PT),
    round(Math.max(...r.map((p) => p[1])) / PT),
  ]);

/** The [left, right] of each ring, in points. */
const spans = (rings: readonly UnderlineRing[]): number[][] =>
  rings.map((r) => [
    round(Math.min(...r.map((p) => p[0])) / PT),
    round(Math.max(...r.map((p) => p[0])) / PT),
  ]);

/** A text run with Carlito's `post` underline: position -103 and thickness 194 of 2048. */
function glyphRun(text: string, x0: number, sizePt: number, underline: string): GlyphRun {
  return {
    kind: 'text',
    text,
    x: Array.from({ length: text.length + 1 }, (_, i) => (x0 + i * 10) * PT),
    sizePt,
    rise: 0,
    runIndex: 0,
    start: 0,
    face: { metrics: { underlinePosition: -103 / 2048, underlineThickness: 194 / 2048 } },
    props: { underline } as RunProps,
  } as unknown as GlyphRun;
}

describe('underlineRings', () => {
  // A 100 pt run on the baseline 0, its single bar 2 pt below it and 4 pt thick.
  const rings = (style: string, x0 = 0): UnderlineRing[] =>
    underlineRings(style, 0, x0 * PT, 100 * PT, 0, 2 * PT, 4 * PT);

  test('draws every underline type', () => {
    for (const style of [
      'words',
      'sng',
      'dbl',
      'heavy',
      'dotted',
      'dottedHeavy',
      'dash',
      'dashHeavy',
      'dashLong',
      'dashLongHeavy',
      'dotDash',
      'dotDashHeavy',
      'dotDotDash',
      'dotDotDashHeavy',
      'wavy',
      'wavyHeavy',
      'wavyDbl',
    ]) {
      // prettier-ignore
      expect(isDrawnUnderline(style)).toBe(true);
      expect(rings(style).length).toBeGreaterThan(0);
    }
    expect(isDrawnUnderline('none')).toBe(false);
  });

  test('single and words are one bar of one thickness, under the spaces too', () => {
    expect(bands(rings('sng'))).toEqual([[2, 6]]);
    expect(bands(rings('words'))).toEqual([[2, 6]]);
    expect(spans(rings('words'))).toEqual([[0, 100]]);
  });

  test('double is two half-thickness bars one thickness apart, from p - t to p + t', () => {
    expect(bands(rings('dbl'))).toEqual([
      [-2, 0],
      [4, 6],
    ]);
  });

  test('heavy is 1.5 thicknesses, from p - t to p + t / 2', () => {
    expect(bands(rings('heavy'))).toEqual([[-2, 4]]);
    expect(bands(rings('dashHeavy'))[0]).toEqual([-2, 4]);
  });

  test('dashes are multiples of the thickness from the origin, cut at the run end', () => {
    // 3 on and 2 off: 12 pt dashes every 20 pt.
    expect(spans(rings('dash'))).toEqual([
      [0, 12],
      [20, 32],
      [40, 52],
      [60, 72],
      [80, 92],
    ]);
    expect(spans(rings('dotDotDash')).slice(0, 4)).toEqual([
      [0, 12],
      [20, 24],
      [32, 36],
      [44, 56],
    ]);
    // A heavy style keeps the pattern of its plain style.
    expect(spans(rings('dashLongHeavy'))).toEqual([
      [0, 28],
      [44, 72],
      [88, 100],
    ]);
    // A run that starts after the origin keeps its phase.
    expect(spans(rings('dash', 30))[0]).toEqual([30, 32]);
  });

  test('a wave is 2 t high on the single bar, with 45° slopes and a 1 pt band', () => {
    const [wave] = rings('wavy');
    expect(bands([wave!])).toEqual([[0, 8]]);
    // The half period is the height less the band: 7 pt. The upper edge starts at the top.
    expect(wave!.slice(0, 3).map(([x, y]) => [round(x / PT), round(y / PT)])).toEqual([
      [0, 0],
      [7, 7],
      [14, 0],
    ]);
    // It ends at the first half period past the run end.
    expect(spans([wave!])).toEqual([[0, 105]]);
  });

  test('a heavy wave has a band of t / 2, at least 2 pt; a double wave has two of t', () => {
    const heavy = rings('wavyHeavy')[0]!;
    expect(round((heavy.at(-1)![1] - heavy[0]![1]) / PT)).toBe(2);
    expect(bands(rings('wavyDbl'))).toEqual([
      [-2, 2],
      [2, 6],
    ]);
  });
});

describe('lineUnderlines', () => {
  test('the thickness rounds to whole points, at least 1 pt', () => {
    // Carlito at 18 pt: 1.71 pt thick, 0.91 pt below the baseline.
    const run = glyphRun('Text', 0, 18, 'sng');
    const [top, bottom] = bands(lineUnderlines([run], 0).get(run)!)[0]!;
    expect(round(bottom - top)).toBe(2);
    expect(top).toBeCloseTo((103 / 2048) * 18, 3);
    const small = glyphRun('Text', 0, 6, 'sng');
    const [a, b] = bands(lineUnderlines([small], 0).get(small)!)[0]!;
    expect(round(b - a)).toBe(1);
  });

  test('runs of two sizes share the bar of the largest, and spaces between runs are underlined', () => {
    const large = glyphRun('Big ', 0, 36, 'sng');
    const small = glyphRun('small ', 40, 14, 'sng');
    const marks = lineUnderlines([large, small], 0);
    expect(bands(marks.get(large)!)).toEqual(bands(marks.get(small)!));
    expect(round(bands(marks.get(small)!)[0]![1] - bands(marks.get(small)!)[0]![0])).toBe(3);
    // The space that ends the first run is underlined; the one that ends the line is not.
    expect(spans(marks.get(large)!)).toEqual([[0, 40]]);
    expect(spans(marks.get(small)!)).toEqual([[40, 90]]);
  });

  test('a run of spaces after the last word has no underline', () => {
    const word = glyphRun('Word', 0, 18, 'sng');
    const spaces = glyphRun('  ', 40, 18, 'sng');
    const marks = lineUnderlines([word, spaces], 0);
    expect(marks.has(spaces)).toBe(false);
    expect(spans(marks.get(word)!)).toEqual([[0, 40]]);
  });

  test('adjacent runs of one dashed style keep one phase', () => {
    // Carlito at 36 pt: 3 pt thick, so dashes of 9 pt every 15 pt.
    const a = glyphRun('ab', 0, 36, 'dash');
    const b = glyphRun('cd', 20, 36, 'dash');
    const marks = lineUnderlines([a, b], 0);
    expect(spans(marks.get(a)!)).toEqual([
      [0, 9],
      [15, 20],
    ]);
    expect(spans(marks.get(b)!)).toEqual([
      [20, 24],
      [30, 39],
    ]);
  });

  test('runs without an underline draw none', () => {
    const run = glyphRun('Text', 0, 18, 'none');
    expect(lineUnderlines([run], 0).size).toBe(0);
  });
});
