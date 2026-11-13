import { describe, expect, test } from 'bun:test';
import type { ResolvedFill, Rgba } from '../../cascade/index.ts';
import type { LayoutItem, ThreeDPaint } from '../../layout/index.ts';
import {
  applyBlipColorSteps,
  blipColorSteps,
  LUMA_WEIGHTS,
  threeDFaceItem,
  tileGeometry,
} from '../painter-math.ts';
import { matrixAttr } from '../svg.ts';

// The color steps of picture effects and the tile geometry of tiled pictures: the rules that the
// SVG paint and the PDF export share.

describe('blipColorSteps', () => {
  test('maps each effect to the SVG filter primitives, in document order', () => {
    const steps = blipColorSteps([
      { kind: 'alphaModFix', amount: 0.5 },
      { kind: 'grayscale' },
      {
        kind: 'duotone',
        dark: { r: 0, g: 51, b: 102, a: 1 },
        light: { r: 255, g: 204, b: 153, a: 1 },
      },
      { kind: 'lum', bright: 0.1, contrast: 0 },
      { kind: 'biLevel', threshold: 0.25 },
      { kind: 'unsupported', name: 'clrChange' },
    ]);
    expect(steps.map((s) => s.op)).toEqual(['gray', 'gray', 'table', 'linear', 'gray', 'linear']);
    expect(steps[2]).toEqual({ op: 'table', from: [0, 0.2, 0.4], to: [1, 0.8, 0.6] });
    const lum = steps[3] as { slope: number; intercept: number };
    expect(lum.slope).toBe(1);
    expect(lum.intercept).toBeCloseTo(0.1, 12);
    expect(steps[5]).toEqual({ op: 'linear', slope: 10_000, intercept: 1 - 10_000 * 0.25 });
  });

  test('computes grayscale, duotone, luminance, and bi-level colors', () => {
    const [wr, wg, wb] = LUMA_WEIGHTS;
    const gray = applyBlipColorSteps(blipColorSteps([{ kind: 'grayscale' }]), [1, 0.5, 0]);
    expect(gray[0]).toBeCloseTo(wr + 0.5 * wg, 9);
    expect(gray[1]).toBe(gray[0]);
    const duotone = applyBlipColorSteps(
      blipColorSteps([
        {
          kind: 'duotone',
          dark: { r: 0, g: 0, b: 255, a: 1 },
          light: { r: 255, g: 255, b: 0, a: 1 },
        },
      ]),
      [0, 0, 1]
    );
    expect(duotone[0]).toBeCloseTo(wb, 9);
    expect(duotone[2]).toBeCloseTo(1 - wb, 9);
    // Brightness past white clamps.
    const bright = blipColorSteps([{ kind: 'lum', bright: 0.6, contrast: 0 }]);
    applyBlipColorSteps(bright, [0.5, 0.2, 0]).forEach((v, i) =>
      expect(v).toBeCloseTo([1, 0.8, 0.6][i]!, 12)
    );
    const bi = blipColorSteps([{ kind: 'biLevel', threshold: 0.5 }]);
    expect(applyBlipColorSteps(bi, [0.49, 0.49, 0.49])).toEqual([0, 0, 0]);
    expect(applyBlipColorSteps(bi, [0.51, 0.51, 0.51])).toEqual([1, 1, 1]);
  });
});

describe('tileGeometry', () => {
  const tile = { tx: 0, ty: 0, sx: 1, sy: 1, flip: 'none', algn: 'tl' };

  test('tiles at the natural size times the scale, from the alignment point plus the offset', () => {
    const g = tileGeometry(
      { ...tile, tx: 100, ty: -50, sx: 0.5, sy: 2, algn: 'br' },
      { cx: 40_000, cy: 20_000 },
      100_000,
      60_000
    );
    expect(g).toEqual({
      x: 100_000 - 20_000 + 100,
      y: 60_000 - 40_000 - 50,
      width: 20_000,
      height: 40_000,
      flipX: false,
      flipY: false,
    });
  });

  test('reads flip on each axis, and an unknown alignment as the top-left corner', () => {
    const g = tileGeometry(
      { ...tile, flip: 'xy', algn: 'toString' },
      { cx: 20_000, cy: 10_000 },
      9e4,
      9e4
    );
    expect([g.x, g.y, g.flipX, g.flipY]).toEqual([0, 0, true, true]);
    expect(tileGeometry({ ...tile, flip: 'y' }, undefined, 9e4, 9e4).flipX).toBe(false);
  });

  test('keeps a tile between one pixel and the largest slide side', () => {
    const tiny = tileGeometry({ ...tile, sx: 1e-9, sy: 1e-9 }, { cx: 10, cy: 10 }, 9e4, 9e4);
    expect([tiny.width, tiny.height]).toEqual([9525, 9525]);
    const huge = tileGeometry({ ...tile, sx: 1000, sy: 500 }, { cx: 1e6, cy: 1e6 }, 9e4, 9e4);
    expect(huge.width).toBe(51_206_400);
    expect(huge.height).toBe(51_206_400 / 2);
  });

  test('a tile is the box without a natural size', () => {
    const g = tileGeometry(tile, undefined, 70_000, 30_000);
    expect([g.width, g.height]).toEqual([70_000, 30_000]);
  });
});

describe('threeDFaceItem', () => {
  const item = (fill: ResolvedFill) =>
    ({ kind: 'shape', element: { fill } }) as unknown as Extract<LayoutItem, { kind: 'shape' }>;
  const paint = (faceColor?: Rgba): ThreeDPaint => ({
    front: [1, 0, 0, 1, 0, 0],
    showsFace: true,
    ...(faceColor ? { faceColor } : {}),
    behind: [],
    under: [],
    over: [],
  });
  const lit = { r: 75, g: 123, b: 208, a: 1 };

  test('a solid face takes the lit color of the paint', () => {
    const solid = item({ type: 'solid', color: { r: 68, g: 114, b: 196, a: 1 } });
    expect(threeDFaceItem(solid, paint(lit)).element.fill).toEqual({ type: 'solid', color: lit });
    // No paint, or a paint with no face color, keeps the item.
    expect(threeDFaceItem(solid, null)).toBe(solid);
    expect(threeDFaceItem(solid, paint())).toBe(solid);
  });

  test('other fills and a face color that is not finite keep the item', () => {
    const none = item({ type: 'none' });
    expect(threeDFaceItem(none, paint(lit))).toBe(none);
    const solid = item({ type: 'solid', color: { r: 68, g: 114, b: 196, a: 1 } });
    expect(threeDFaceItem(solid, paint({ ...lit, g: Number.NaN }))).toBe(solid);
  });
});

describe('matrixAttr', () => {
  test('keeps the linear part to 1e-6 and the offset to 0.01', () => {
    // A turned face: a coefficient rounded to 0.01 moves a point at 5,000,000 EMU by 25,000.
    expect(matrixAttr([0.898794, -0.135, 0.0000004, 0.951057, 616936.114, -0.004])).toBe(
      'matrix(0.898794 -0.135 0 0.951057 616936.11 0)'
    );
  });
});
