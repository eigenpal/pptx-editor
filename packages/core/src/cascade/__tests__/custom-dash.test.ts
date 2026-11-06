// `a:custDash` (ECMA-376 Part 1, 20.1.8.21) is untrusted input: the parse keeps at most 64 dash
// and space pairs, makes a negative length 0, and draws a pattern with no length as solid. The
// SVG paint and the PDF export both read this result.

import { describe, expect, test } from 'bun:test';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '../../store/index.ts';
import { dashArray } from '../../layout/index.ts';
import { resolveSlide, type ResolvedLine } from '../index.ts';

function lineOf(pairs: string): ResolvedLine {
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          x: 0,
          y: 0,
          cx: 914_400,
          cy: 0,
          prst: 'line',
          line:
            '<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>' +
            `<a:custDash>${pairs}</a:custDash></a:ln>`,
        }),
      },
    ],
  });
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('open failed');
  const el = resolveSlide(r.pkg, 0)!.elements[0];
  if (el.kind !== 'shape' || !el.line) throw new Error('no line');
  return el.line;
}

const ds = (d: number, sp: number): string => `<a:ds d="${d}" sp="${sp}"/>`;

describe('custom dashes at the parse boundary', () => {
  test('a pattern keeps its pairs in line widths', () => {
    const line = lineOf(ds(300_000, 100_000) + ds(100_000, 100_000));
    expect(line.dash).toBe('custom');
    expect(line.dashes).toEqual([3, 1, 1, 1]);
  });

  test('a huge list keeps only the first 64 pairs', () => {
    const line = lineOf(ds(100_000, 100_000).repeat(10_000));
    expect(line.dashes?.length).toBe(128);
    expect(dashArray(line, 1)?.length).toBe(128);
  });

  test('an all-zero pattern draws solid in both outputs', () => {
    const line = lineOf(ds(0, 0).repeat(3));
    expect(line.dash).toBe('solid');
    expect(line.dashes).toBeUndefined();
    expect(dashArray(line, 12_700)).toBeUndefined();
  });

  test('a negative length is 0, and a pattern of only negative lengths draws solid', () => {
    const line = lineOf(ds(-300_000, 100_000) + ds(200_000, -50_000));
    expect(line.dashes).toEqual([0, 1, 2, 0]);
    expect(lineOf(ds(-300_000, -100_000)).dash).toBe('solid');
  });

  test('a period under a hundredth of a width draws solid', () => {
    expect(lineOf(ds(1, 1).repeat(64)).dash).toBe('solid');
  });
});
