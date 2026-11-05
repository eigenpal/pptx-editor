import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import { resolveSlide } from '../../cascade/index.ts';
import { createFixedMeasurer, layoutSlide } from '../../layout/index.ts';
import { readPptx } from '../../store/index.ts';
import { compoundBands } from '../compound-lines.ts';
import { paintSlide } from '../index.ts';
import {
  PATTERN_PIXEL,
  PRESET_PATTERNS,
  isMeasuredPattern,
  patternBitmap,
  patternDensity,
} from '../preset-patterns.ts';

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const PATTERN = (prst: string): string =>
  `<a:pattFill prst="${prst}"><a:fgClr><a:srgbClr val="2F5597"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>`;

/** The `<pattern>` element that paints the one pattern-filled shape of a slide. */
function patternOf(shapes: string): Element {
  const r = readPptx(buildPptx({ slides: [{ shapes, layout: 'blank' }] }));
  if (!r.ok) throw new Error('open failed');
  const layout = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer());
  const root = paintSlide(doc, layout, { images: () => undefined, idPrefix: 't' });
  return [...root.querySelectorAll('pattern')].find((p) => p.id.includes('patt'))!;
}

describe('preset patterns', () => {
  test('all 54 ST_PresetPatternVal names have a bitmap', () => {
    expect(PRESET_PATTERNS.length).toBe(54);
    for (const p of PRESET_PATTERNS) expect(patternBitmap(p)?.length).toBe(8);
    expect(patternBitmap('nope')).toBeUndefined();
    expect(patternBitmap('__proto__')).toBeUndefined();
    expect(isMeasuredPattern('__proto__')).toBe(false);
  });

  test('approximated percentage patterns have their ink density', () => {
    for (const p of [5, 10, 20, 25, 30, 40, 50, 60, 70, 75, 80, 90]) {
      if (isMeasuredPattern(`pct${p}`)) continue;
      expect(Math.abs(patternDensity(`pct${p}`)! - p / 100)).toBeLessThan(1 / 64 + 1e-9);
    }
  });

  test('measured presets keep the bitmaps of the reference PDF', () => {
    // fixtures/fills-lines.pdf: pct75 leaves 8 of 64 pixels, and dnDiag has half-tone neighbors.
    expect(patternDensity('pct25')).toBe(16 / 64);
    expect(patternDensity('pct75')).toBe(56 / 64);
    expect(patternBitmap('dnDiag')![0]).toEqual([1, 0.5, 0, 0, 0, 0, 0, 0.5]);
    expect(patternBitmap('cross')!.map((row) => row[0])).toEqual([1, 1, 1, 1, 1, 1, 1, 1]);
    expect(patternBitmap('cross')![0].every((v) => v === 1)).toBe(true);
  });

  test('a pattern pixel is 1 pt, and tiles sit on the slide grid', () => {
    expect(PATTERN_PIXEL).toBe(12700);
    // A box at (36 pt, 216 pt) is 4 pt right of a tile edge: the tile starts 4 pt before the box.
    const p = patternOf(
      shape({ id: 2, x: 457_200, y: 2_743_200, cx: 1_280_160, cy: 822_960, fill: PATTERN('cross') })
    );
    expect(Number(p.getAttribute('width'))).toBe(8 * 12700);
    expect(Number(p.getAttribute('x'))).toBe(-4 * 12700);
    expect(Number(p.getAttribute('y'))).toBe(0);
  });

  test('a rotated box keeps its tiles at its own corner', () => {
    const p = patternOf(
      shape({
        id: 2,
        x: 457_200,
        y: 2_743_200,
        cx: 1_280_160,
        cy: 822_960,
        rot: 1_800_000,
        fill: PATTERN('cross'),
      })
    );
    expect(Number(p.getAttribute('x'))).toBe(0);
  });

  test('a half-tone pixel draws the foreground at half opacity', () => {
    const p = patternOf(
      shape({ id: 2, x: 0, y: 0, cx: 914_400, cy: 914_400, fill: PATTERN('dnDiag') })
    );
    const half = [...p.querySelectorAll('rect')].filter(
      (r) => r.getAttribute('fill-opacity') === '0.5'
    );
    expect(half.length).toBe(16);
  });
});

describe('compound line bands', () => {
  const w = 60;
  test('dbl on a closed path: one band on each side', () => {
    const b = compoundBands('dbl', w, true)!;
    for (const side of [b.outer, b.inner]) {
      expect(side).toHaveLength(1);
      expect(side[0].a).toBeCloseTo(10);
      expect(side[0].b).toBeCloseTo(30);
    }
  });

  test('thickThin: the thick band outside, the thin band inside', () => {
    const b = compoundBands('thickThin', w, true)!;
    expect(b.outer).toHaveLength(1);
    expect(b.outer[0].b).toBeCloseTo(30);
    expect(b.outer[0].a).toBe(0);
    // The thick band crosses the center line by 0.1 × width.
    expect(b.inner[0].b).toBeCloseTo(6);
    expect(b.inner[1]).toMatchObject({ a: expect.closeTo(18), b: expect.closeTo(30) });
  });

  test('an open path draws asymmetric types with dbl bands; single lines are not compound', () => {
    expect(compoundBands('thinThick', w, false)).toEqual(compoundBands('dbl', w, false));
    expect(compoundBands('sng', w, true)).toBeNull();
  });

  test('tri: the middle band straddles the center line', () => {
    const b = compoundBands('tri', w, true)!;
    expect(b.outer.map((x) => [Math.round(x.a), Math.round(x.b)])).toEqual([
      [20, 30],
      [0, 10],
    ]);
  });
});
