/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The glow, soft edge, inner shadow, and reflection rules of the Pro effects renderer
// (ECMA-376 Part 1, 20.1.8: CT_InnerShadowEffect, CT_GlowEffect, CT_SoftEdgesEffect,
// CT_ReflectionEffect). The constants are measured on the `effects` and `shape-effects-presets`
// reference PDFs.

import { describe, expect, test } from 'bun:test';
import type { ResolvedEffects } from '@pptx-editor.dev/core/cascade';
import type { EffectFilterPaint, EffectIds, EffectNode } from '@pptx-editor.dev/core/layout';
import {
  fadeLine,
  glowSpread,
  innerShadowShift,
  reflectionBlurSigma,
  reflectionFade,
  softEdgeSpread,
} from '../effects-math.ts';
import { bodyEffects, glowEffect } from '../effects-renderer.ts';

const PT = 12_700;
const black = { r: 0, g: 0, b: 0, a: 0.5 };
const bounds = { x: 0, y: 0, r: 2_000_000, b: 1_000_000 };

function ids(): EffectIds {
  let next = 0;
  return { id: (kind) => `t-${kind}${(next += 1)}` };
}

/** The `<filter>` that a paint applies. */
function filterOf(paint: EffectFilterPaint | null): EffectNode {
  if (!paint) throw new Error('no filter');
  const filter = paint.defs.find((d) => d.attrs.id === paint.filter);
  if (!filter) throw new Error('filter not in defs');
  return filter;
}

/** The primitives of the filter that a paint applies. */
function primitives(paint: EffectFilterPaint | null): EffectNode[] {
  return [...(filterOf(paint).children ?? [])];
}

const attr = (els: EffectNode[], tag: string, name: string): (string | undefined)[] =>
  els.filter((e) => e.tag === tag).map((e) => e.attrs[name]);

describe('reflection kernel', () => {
  test('a reflection keeps its unmeasured kernel', () => {
    expect(reflectionBlurSigma(12_700)).toBe(6_350);
    expect(reflectionBlurSigma(-1)).toBe(0);
  });

  test('the fade line spans the bounds along the fade direction', () => {
    const l = fadeLine({ x: 0, y: 0, r: 200, b: 100 }, 90);
    expect(l.x1).toBeCloseTo(100, 6);
    expect(l.y1).toBeCloseTo(0, 6);
    expect(l.x2).toBeCloseTo(100, 6);
    expect(l.y2).toBeCloseTo(100, 6);
  });

  test('the fade spans the reflected box, clamps its stops, and turns with the shape', () => {
    const ref = {
      blur: 0,
      dist: 0,
      dir: 90,
      sx: 1,
      sy: -1,
      kx: 0,
      ky: 0,
      align: 'bl' as const,
      rotWithShape: false,
      startAlpha: 0.5,
      startPos: -0.2,
      endAlpha: 0,
      endPos: 0.4,
      fadeDir: 90,
    };
    // A 200 × 100 shape at the origin, mirrored below its bottom edge.
    const flat = reflectionFade(ref, [1, 0, 0, 1, 0, 0], 200, 100);
    expect(flat.line.y1).toBeCloseTo(100, 6);
    expect(flat.line.y2).toBeCloseTo(200, 6);
    expect([flat.start, flat.end]).toEqual([0, 0.4]);
    // An end before the start is a hard step at the start.
    expect(
      reflectionFade({ ...ref, startPos: 0.5, endPos: 0.1 }, [1, 0, 0, 1, 0, 0], 200, 100).end
    ).toBe(0.5);
    // Turned a quarter: with the shape, the fade runs along the turned frame.
    const quarter = [0, 1, -1, 0, 0, 0] as const;
    const still = reflectionFade(ref, quarter, 200, 100).line;
    const turned = reflectionFade({ ...ref, rotWithShape: true }, quarter, 200, 100).line;
    expect(Math.abs(still.x2 - still.x1)).toBeCloseTo(0, 6);
    expect(Math.abs(turned.y2 - turned.y1)).toBeCloseTo(0, 6);
  });
});

describe('glow', () => {
  test('the outline grows by half the radius and blurs by the other half', () => {
    // Glow 18 pt: the half level at 9 pt outside the edge, the tail ending at 18 pt.
    const { spread, sigma } = glowSpread(18 * PT);
    expect(spread).toBe(9 * PT);
    expect(sigma).toBeCloseTo((9 * PT) / 3, 6);
    expect(spread + 3 * sigma).toBeCloseTo(18 * PT, 6);
    expect(glowSpread(-PT)).toEqual({ spread: 0, sigma: 0 });
  });

  test('the filter dilates the alpha, blurs, and keeps the content on top', () => {
    const fx: ResolvedEffects = { glow: { radius: 18 * PT, color: black }, unsupported: [] };
    const prims = primitives(glowEffect(fx, bounds, ids()));
    expect(attr(prims, 'feMorphology', 'operator')).toEqual(['dilate', 'dilate', 'dilate']);
    // The union is the largest alpha: `lighten` on the alpha moved into an opaque gray image,
    // so a fill of alpha 0.25 keeps a glow of strength 0.25.
    expect(attr(prims, 'feBlend', 'mode')).toEqual(['lighten', 'lighten']);
    expect(attr(prims, 'feColorMatrix', 'in')[0]).toBe('SourceAlpha');
    // Alpha into R, G, and B of an opaque image, and the red channel back into alpha.
    const matrices = attr(prims, 'feColorMatrix', 'values').map((v) => v!.split(/\s+/).map(Number));
    expect(matrices).toEqual([
      [0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 1],
      [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0],
    ]);
    // A one-pixel horizontal bar, a vertical bar, and a square of spread / √2.
    const radii = attr(prims, 'feMorphology', 'radius').map((r) => r!.split(' ').map(Number));
    const d = (9 * PT) / Math.SQRT2;
    expect(radii[0]).toEqual([9 * PT, 9525]);
    expect(radii[1]).toEqual([9525, 9 * PT]);
    expect(radii[2]![0]).toBeCloseTo(d, 0);
    expect(radii[2]![1]).toBeCloseTo(d, 0);
    expect(Number(attr(prims, 'feGaussianBlur', 'stdDeviation')[0])).toBeCloseTo(3 * PT, 1);
    // The region holds the whole glow: the radius and a pixel on each side.
    const filter = filterOf(glowEffect(fx, bounds, ids()));
    const m = 18 * PT + 9525;
    expect(Number(filter.attrs['x'])).toBeCloseTo(-m, 0);
    expect(Number(filter.attrs['y'])).toBeCloseTo(-m, 0);
    expect(Number(filter.attrs['width'])).toBeCloseTo(2_000_000 + 2 * m, 0);
    expect(Number(filter.attrs['height'])).toBeCloseTo(1_000_000 + 2 * m, 0);
  });

  test('a glow of radius 0 makes no filter', () => {
    const fx: ResolvedEffects = { glow: { radius: 0, color: black }, unsupported: [] };
    expect(glowEffect(fx, bounds, ids())).toBeNull();
  });
});

describe('soft edge', () => {
  test('the outline shrinks by 0.84 of the radius and blurs by the radius', () => {
    const { spread, sigma } = softEdgeSpread(10 * PT);
    expect(spread).toBeCloseTo(-8.4 * PT, 6);
    expect(sigma).toBeCloseTo((10 * PT) / 3, 6);
  });

  test('the mask erodes the alpha and keeps the shape transparency once', () => {
    const fx: ResolvedEffects = { softEdge: { radius: 10 * PT }, unsupported: [] };
    const prims = primitives(bodyEffects(fx, bounds, ids()));
    expect(attr(prims, 'feMorphology', 'operator')).toEqual(['erode', 'erode', 'erode']);
    // The intersection is the smallest alpha (`darken`), so inside a fill of alpha 0.25 the
    // mask is 0.25. The shape's colors, made opaque, take the mask: the fill alpha counts once.
    expect(attr(prims, 'feBlend', 'mode')).toEqual(['darken', 'darken']);
    const opaque = prims.find((e) => e.attrs.result === 'so')!;
    expect(opaque.attrs.in).toBe('SourceGraphic');
    expect(opaque.children?.[0]?.attrs.tableValues).toBe('1');
    const last = prims.at(-1)!;
    expect([last.attrs.in, last.attrs.in2]).toEqual(['sc', 'sb']);
    const masked = prims.find((e) => e.attrs.result === 'sc')!;
    expect([masked.attrs.in, masked.attrs.in2]).toEqual(['so', 'sm']);
  });

  test('the region holds the shrink of a large soft edge on a small shape', () => {
    const small = { x: 0, y: 0, r: 100_000, b: 100_000 };
    const filter = filterOf(
      bodyEffects({ softEdge: { radius: 100_000 }, unsupported: [] }, small, ids())
    );
    expect(Number(filter.attrs['x'])).toBeLessThanOrEqual(-84_000 - 9525);
  });

  test('a soft edge of radius 0 makes no filter', () => {
    expect(bodyEffects({ softEdge: { radius: 0 }, unsupported: [] }, bounds, ids())).toBeNull();
  });
});

describe('inner shadow', () => {
  test('the inverse moves opposite dir, so the shadow shows on the dir side', () => {
    // dir="0" shades the right edge: the inverse moves left.
    const right = innerShadowShift(4 * PT, 0);
    expect(right.dx).toBeCloseTo(-4 * PT, 6);
    expect(right.dy).toBeCloseTo(0, 6);
    // dir="5400000" (90°) shades the bottom edge: the inverse moves up.
    const bottom = innerShadowShift(4 * PT, 90);
    expect(bottom.dx).toBeCloseTo(0, 6);
    expect(bottom.dy).toBeCloseTo(-4 * PT, 6);
    // dir="13500000" (225°) shades the top and left edges: the inverse moves down and right.
    const topLeft = innerShadowShift(4 * PT, 225);
    expect(topLeft.dx).toBeCloseTo((4 * PT) / Math.SQRT2, 6);
    expect(topLeft.dy).toBeCloseTo((4 * PT) / Math.SQRT2, 6);
  });

  test('the filter offsets the inverse and leaves room for its blur', () => {
    const blur = 5 * PT;
    const fx: ResolvedEffects = {
      innerShadow: { blur, dist: 4 * PT, dir: 0, color: black },
      unsupported: [],
    };
    const paint = bodyEffects(fx, bounds, ids());
    const prims = primitives(paint);
    expect(Number(attr(prims, 'feOffset', 'dx')[0])).toBeCloseTo(-4 * PT, 0);
    expect(Number(attr(prims, 'feGaussianBlur', 'stdDeviation')[0])).toBeCloseTo(blur / 3, 1);
    // A tenth of the larger side, the offset, three σ, and a pixel.
    const filter = filterOf(paint);
    const margin = 200_000 + 4 * PT + blur + 9525;
    expect(Number(filter.attrs['x'])).toBeCloseTo(-margin, 0);
    expect(filter.attrs['filterUnits']).toBe('userSpaceOnUse');
  });
});
