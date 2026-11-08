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
  patternBitmap,
  patternDensity,
} from '../preset-patterns.ts';

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const PATTERN = (prst: string): string =>
  `<a:pattFill prst="${prst}"><a:fgClr><a:srgbClr val="2F5597"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>`;

/** The painted SVG of a one-slide deck with `shapes`. */
function paintShapes(shapes: string): Element {
  const r = readPptx(buildPptx({ slides: [{ shapes, layout: 'blank' }] }));
  if (!r.ok) throw new Error('open failed');
  const layout = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer());
  return paintSlide(doc, layout, { images: () => undefined, idPrefix: 't' });
}

/** The `<pattern>` element that paints the one pattern-filled shape of a slide. */
function patternOf(shapes: string): Element {
  return [...paintShapes(shapes).querySelectorAll('pattern')].find((p) => p.id.includes('patt'))!;
}

const parseMatrix = (t: string): number[] =>
  t
    .replace(/^matrix\(|\)$/g, '')
    .split(' ')
    .map(Number);

describe('preset patterns', () => {
  test('all 54 ST_PresetPatternVal names have a bitmap', () => {
    expect(PRESET_PATTERNS.length).toBe(54);
    for (const p of PRESET_PATTERNS) expect(patternBitmap(p)?.length).toBe(8);
    expect(patternBitmap('nope')).toBeUndefined();
    expect(patternBitmap('__proto__')).toBeUndefined();
    expect(patternBitmap('constructor')).toBeUndefined();
  });

  test('the bitmaps are the ones of the reference PDF', () => {
    // fixtures/follow-ups.pdf, slide 2: the percentage presets do not have their named density.
    expect(patternDensity('pct5')).toBe(2 / 64);
    expect(patternDensity('pct25')).toBe(16 / 64);
    expect(patternDensity('pct75')).toBe(56 / 64);
    expect(patternDensity('pct90')).toBe(62 / 64);
    expect(patternBitmap('dnDiag')![0]).toEqual([1, 0.5, 0, 0, 0, 0, 0, 0.5]);
    expect(patternBitmap('cross')).toEqual(patternBitmap('lgGrid'));
    expect(patternBitmap('horz')![0].every((v) => v === 1)).toBe(true);
    expect(
      patternBitmap('horz')!
        .slice(1)
        .flat()
        .every((v) => v === 0)
    ).toBe(true);
    expect(patternBitmap('narVert')!.map((row) => row.join(''))).toEqual(Array(8).fill('01010101'));
  });

  test('a pattern pixel is 1 pt, and tiles sit on the slide grid', () => {
    expect(PATTERN_PIXEL).toBe(12700);
    // A box at (36 pt, 216 pt): the pattern space is the slide's, so its origin is the slide's.
    const p = patternOf(
      shape({ id: 2, x: 457_200, y: 2_743_200, cx: 1_280_160, cy: 822_960, fill: PATTERN('cross') })
    );
    expect(Number(p.getAttribute('width'))).toBe(8 * 12700);
    expect(parseMatrix(p.getAttribute('patternTransform')!)).toEqual([
      1, 0, 0, 1, -457_200, -2_743_200,
    ]);
  });

  test('a rotated and flipped box fills in slide space, so its tiles stay square', () => {
    // fixtures/follow-ups.pdf, slide 3: turned and flipped shapes show unturned tiles.
    const root = paintShapes(
      shape({
        id: 2,
        x: 457_200,
        y: 2_743_200,
        cx: 1_280_160,
        cy: 822_960,
        rot: 1_800_000,
        flipH: true,
        fill: PATTERN('horz'),
      })
    );
    const pattern = [...root.querySelectorAll('pattern')].find((p) => p.id.includes('patt'))!;
    expect(pattern.getAttribute('patternTransform')).toBeNull();
    const filled = [...root.querySelectorAll('path')].find(
      (p) => p.getAttribute('fill') === `url(#${pattern.id})`
    )!;
    const wrap = filled.parentElement!;
    // The wrapper undoes the body's matrix: the two compose to the identity.
    const [a, b, c, d, e, f] = parseMatrix(wrap.parentElement!.getAttribute('transform')!);
    const [A, B, C, D, E, F] = parseMatrix(wrap.getAttribute('transform')!);
    const product = [
      a * A + c * B,
      b * A + d * B,
      a * C + c * D,
      b * C + d * D,
      a * E + c * F + e,
      b * E + d * F + f,
    ];
    product.forEach((v, i) => expect(v).toBeCloseTo([1, 0, 0, 1, 0, 0][i], 3));
    // The fill path is in slide space: its first point is a corner of the turned box.
    expect(filled.getAttribute('d')!.startsWith('M')).toBe(true);
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
  test('dbl: one band on each side', () => {
    const b = compoundBands('dbl', w)!;
    for (const side of [b.outer, b.inner]) {
      expect(side).toHaveLength(1);
      expect(side[0].a).toBeCloseTo(10);
      expect(side[0].b).toBeCloseTo(30);
    }
  });

  test('thickThin: the thick band outside, the thin band inside', () => {
    const b = compoundBands('thickThin', w)!;
    expect(b.outer).toHaveLength(1);
    expect(b.outer[0].b).toBeCloseTo(30);
    expect(b.outer[0].a).toBe(0);
    // The thick band crosses the center line by 0.1 × width.
    expect(b.inner[0].b).toBeCloseTo(6);
    expect(b.inner[1]).toMatchObject({ a: expect.closeTo(18), b: expect.closeTo(30) });
  });

  test('single lines and unknown types are not compound', () => {
    expect(compoundBands('sng', w)).toBeNull();
    expect(compoundBands('__proto__', w)).toBeNull();
  });

  test('tri: the middle band straddles the center line', () => {
    const b = compoundBands('tri', w)!;
    expect(b.outer.map((x) => [Math.round(x.a), Math.round(x.b)])).toEqual([
      [20, 30],
      [0, 10],
    ]);
  });
});

/** The band elements of the one compound outline of a slide. */
function bandsOf(shapes: string): Element[] {
  return [...paintShapes(shapes).querySelector('[data-compound]')!.children];
}

const LINE = (cmpd: string, extra = ''): string =>
  `<a:ln w="76200" cmpd="${cmpd}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>${extra}</a:ln>`;

describe('compound line paint', () => {
  test('thickThin on an open path: the thick band on the left of the direction', () => {
    // fixtures/fills-lines.pdf, slide 2: a left-to-right line has its thick band above.
    const bands = bandsOf(
      shape({ id: 2, x: 0, y: 914_400, cx: 914_400, cy: 0, prst: 'line', line: LINE('thickThin') })
    );
    expect(bands.map((b) => [b.getAttribute('d'), Number(b.getAttribute('stroke-width'))])).toEqual(
      [
        ['M0 -15240L914400 -15240', 45720],
        ['M0 30480L914400 30480', 15240],
      ]
    );
  });

  test('a solid line on a closed path: bands offset outward and inward', () => {
    const bands = bandsOf(
      shape({
        id: 2,
        x: 0,
        y: 0,
        cx: 914_400,
        cy: 914_400,
        line: LINE('thinThick', '<a:miter lim="800000"/>'),
      })
    );
    // The thin outer band is centered 30480 EMU outside the rectangle, the thick one 15240 inside.
    expect(bands.map((b) => b.getAttribute('d'))).toEqual([
      'M-30480 -30480L944880 -30480L944880 944880L-30480 944880Z',
      'M15240 15240L899160 15240L899160 899160L15240 899160Z',
    ]);
  });

  test('a round join rounds the outer corners of the outer bands', () => {
    const bands = bandsOf(
      shape({ id: 2, x: 0, y: 0, cx: 914_400, cy: 914_400, line: LINE('dbl', '<a:round/>') })
    );
    // An arc of 9 points about each corner on the outside, a sharp corner on the inside.
    expect(bands[0].getAttribute('d')!.split('L').length).toBe(4 * 10);
    expect(bands[1].getAttribute('d')!.split('L').length).toBe(4);
  });

  test('a dashed line on a closed path masks the bands, dashed from the closing segment', () => {
    // fixtures/fills-lines.pdf, slide 2: the dashes of a `tri` roundRect start where it closes.
    const root = paintShapes(
      shape({
        id: 2,
        x: 0,
        y: 0,
        cx: 914_400,
        cy: 914_400,
        line: LINE('tri', '<a:prstDash val="dash"/>'),
      })
    );
    const masked = [...root.querySelectorAll('mask path[stroke-dasharray]')];
    expect(masked.length).toBeGreaterThan(0);
    for (const p of masked) expect(p.getAttribute('d')).toMatch(/^M0 914400L0 0/);
  });
});
