// The blur, glow, soft edge, and inner shadow rules of the SVG paint (ECMA-376 Part 1, 20.1.8:
// CT_OuterShadowEffect, CT_InnerShadowEffect, CT_GlowEffect, CT_SoftEdgesEffect). The constants
// are measured on the `effects` and `shape-effects-presets` reference PDFs.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import type { ResolvedEffects } from '../../cascade/index.ts';
import {
  BLUR_SIGMA_PER_RADIUS,
  blurSigma,
  glowSpread,
  innerShadowShift,
  reflectionBlurSigma,
  softEdgeSpread,
} from '../effect-transforms.ts';
import { bodyFilter, glowFilter } from '../paint-effects.ts';
import { Defs } from '../svg.ts';

// A DOM for this file only: the engine test run registers no global DOM.
const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const PT = 12_700;
const black = { r: 0, g: 0, b: 0, a: 0.5 };
const bounds = { x: 0, y: 0, r: 2_000_000, b: 1_000_000 };

/** The primitives of the filter that `url(#id)` names. */
function primitives(defs: Defs, url: string | undefined): Element[] {
  if (!url) throw new Error('no filter');
  const filter = defs.find(url);
  if (!filter) throw new Error('filter not in defs');
  return [...filter.children];
}

const attr = (els: Element[], tag: string, name: string): (string | null)[] =>
  els.filter((e) => e.tagName.toLowerCase() === tag.toLowerCase()).map((e) => e.getAttribute(name));

describe('blur kernel', () => {
  test('a blur radius is three standard deviations', () => {
    expect(BLUR_SIGMA_PER_RADIUS).toBeCloseTo(1 / 3, 12);
    // Offset: center (shape-effects-presets slide 1): blurRad 63500 is 10 px, σ about 3.3 px.
    expect(blurSigma(63_500)).toBeCloseTo(21_166.67, 1);
    expect(blurSigma(-5)).toBe(0);
  });

  test('a reflection keeps its unmeasured kernel', () => {
    expect(reflectionBlurSigma(12_700)).toBe(6_350);
    expect(reflectionBlurSigma(-1)).toBe(0);
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
    const defs = new Defs(doc, 'g');
    const fx: ResolvedEffects = { glow: { radius: 18 * PT, color: black }, unsupported: [] };
    const prims = primitives(defs, glowFilter(defs, fx, bounds));
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
    const filter = defs.find(glowFilter(defs, fx, bounds)!)!;
    const m = 18 * PT + 9525;
    expect(Number(filter.getAttribute('x'))).toBeCloseTo(-m, 0);
    expect(Number(filter.getAttribute('y'))).toBeCloseTo(-m, 0);
    expect(Number(filter.getAttribute('width'))).toBeCloseTo(2_000_000 + 2 * m, 0);
    expect(Number(filter.getAttribute('height'))).toBeCloseTo(1_000_000 + 2 * m, 0);
  });

  test('a glow of radius 0 makes no filter', () => {
    const defs = new Defs(doc, 'z');
    const fx: ResolvedEffects = { glow: { radius: 0, color: black }, unsupported: [] };
    expect(glowFilter(defs, fx, bounds)).toBeUndefined();
  });
});

describe('soft edge', () => {
  test('the outline shrinks by 0.84 of the radius and blurs by the radius', () => {
    const { spread, sigma } = softEdgeSpread(10 * PT);
    expect(spread).toBeCloseTo(-8.4 * PT, 6);
    expect(sigma).toBeCloseTo((10 * PT) / 3, 6);
  });

  test('the mask erodes the alpha and keeps the shape transparency once', () => {
    const defs = new Defs(doc, 's');
    const fx: ResolvedEffects = { softEdge: { radius: 10 * PT }, unsupported: [] };
    const prims = primitives(defs, bodyFilter(defs, fx, bounds));
    expect(attr(prims, 'feMorphology', 'operator')).toEqual(['erode', 'erode', 'erode']);
    // The intersection is the smallest alpha (`darken`), so inside a fill of alpha 0.25 the
    // mask is 0.25. The shape's colors, made opaque, take the mask: the fill alpha counts once.
    expect(attr(prims, 'feBlend', 'mode')).toEqual(['darken', 'darken']);
    const opaque = prims.find((e) => e.getAttribute('result') === 'so')!;
    expect(opaque.getAttribute('in')).toBe('SourceGraphic');
    expect(opaque.firstElementChild?.getAttribute('tableValues')).toBe('1');
    const last = prims.at(-1)!;
    expect([last.getAttribute('in'), last.getAttribute('in2')]).toEqual(['sc', 'sb']);
    const masked = prims.find((e) => e.getAttribute('result') === 'sc')!;
    expect([masked.getAttribute('in'), masked.getAttribute('in2')]).toEqual(['so', 'sm']);
  });

  test('the region holds the shrink of a large soft edge on a small shape', () => {
    const defs = new Defs(doc, 'l');
    const small = { x: 0, y: 0, r: 100_000, b: 100_000 };
    const url = bodyFilter(defs, { softEdge: { radius: 100_000 }, unsupported: [] }, small);
    const filter = defs.find(url!)!;
    expect(Number(filter.getAttribute('x'))).toBeLessThanOrEqual(-84_000 - 9525);
  });

  test('a soft edge of radius 0 makes no filter', () => {
    const defs = new Defs(doc, 'z');
    expect(bodyFilter(defs, { softEdge: { radius: 0 }, unsupported: [] }, bounds)).toBeUndefined();
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
    const defs = new Defs(doc, 'i');
    const blur = 5 * PT;
    const fx: ResolvedEffects = {
      innerShadow: { blur, dist: 4 * PT, dir: 0, color: black },
      unsupported: [],
    };
    const url = bodyFilter(defs, fx, bounds);
    const prims = primitives(defs, url);
    expect(Number(attr(prims, 'feOffset', 'dx')[0])).toBeCloseTo(-4 * PT, 0);
    expect(Number(attr(prims, 'feGaussianBlur', 'stdDeviation')[0])).toBeCloseTo(blur / 3, 1);
    // A tenth of the larger side, the offset, three σ, and a pixel.
    const filter = defs.find(url!)!;
    const margin = 200_000 + 4 * PT + blur + 9525;
    expect(Number(filter.getAttribute('x'))).toBeCloseTo(-margin, 0);
    expect(filter.getAttribute('filterUnits')).toBe('userSpaceOnUse');
  });
});
