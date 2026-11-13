/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The rules of glow, inner shadow, reflection, and soft edge (design D2–D4 of `effects-fidelity`).
// Pure math: no DOM. The SVG renderer (`effects-renderer.ts`) and the PDF writer
// (`@pptx-editor.dev/pptx-to-pdf`) use the same copy. The blur kernel and the effect transform
// are the engine's (`blurSigma` and `effectMatrix` in `@pptx-editor.dev/core/output`), because
// the outer shadow uses them too.

import type { ResolvedEffects } from '@pptx-editor.dev/core/cascade';
import { blurSigma, effectMatrix } from '@pptx-editor.dev/core/output';
import { apply, multiply, type EffectBounds, type Matrix } from '@pptx-editor.dev/core/layout';

const DEG = Math.PI / 180;

/** One pixel at 96 dpi, in EMU: filter regions keep it so they never clip antialiased edges. */
export const PX = 9525;

/**
 * Gaussian σ per EMU of a reflection's `blurRad` (CT_ReflectionEffect). UNMEASURED: the fixture
 * reflections blur by 1 to 6 px, and at that size the reference raster of picture edges differs
 * from the browser's by more than the blur does (`picture-editing` slide 1). So reflections keep
 * σ = `blurRad` / 2 until a fixture with a larger reflection blur measures it.
 */
export const REFLECTION_BLUR_SIGMA_PER_RADIUS = 0.5;

/**
 * The Gaussian σ for a reflection's blur radius.
 *
 * @internal
 */
export function reflectionBlurSigma(blurRad: number): number {
  return Math.max(0, blurRad) * REFLECTION_BLUR_SIGMA_PER_RADIUS;
}

/**
 * How far the outline moves and how much it then blurs, in EMU. A positive `spread` grows the
 * outline (a dilation), a negative one shrinks it (an erosion).
 *
 * @internal
 */
export interface SpreadBlur {
  /** How far the outline moves, in EMU: outward when positive. */
  readonly spread: number;
  /** The Gaussian σ of the blur after the move, in EMU. */
  readonly sigma: number;
}

/**
 * A glow of radius `rad` (ECMA-376 Part 1, 20.1.8, CT_GlowEffect: "a color blurred outline is
 * added outside the edges of the object"). The reference PDFs put the glow's half level at half
 * the radius outside the edge, and its tail ends at the radius. So the outline grows by `rad / 2`
 * and blurs with a radius of `rad / 2`. The glow is at its full alpha on the edge.
 *
 * @internal
 */
export function glowSpread(rad: number): SpreadBlur {
  const half = Math.max(0, rad) / 2;
  return { spread: half, sigma: blurSigma(half) };
}

/**
 * Where a soft edge (ECMA-376 Part 1, 20.1.8, CT_SoftEdgesEffect) puts its half level, as a
 * fraction of `rad` inside the edge. The specification says only that the edges are blurred by
 * the radius. The reference PDFs put the half level at 0.82 to 0.88 of the radius inside the edge
 * (soft edges of 5 and 20 px, that is 2.5 and 10 pt, in `shape-effects-presets` slide 2 and
 * `effects` slide 1), and the shape is clear on its outline. Larger radii are not measured: the
 * rule is applied to them unchanged.
 */
export const SOFT_EDGE_INSET_PER_RADIUS = 0.84;

/**
 * A soft edge of radius `rad`: the outline shrinks by 0.84 `rad` and blurs with a radius of
 * `rad`, so the alpha rises from 0 on the outline to 1 at about 1.84 `rad` inside it.
 *
 * @internal
 */
export function softEdgeSpread(rad: number): SpreadBlur {
  const r = Math.max(0, rad);
  return { spread: -r * SOFT_EDGE_INSET_PER_RADIUS, sigma: blurSigma(r) };
}

/**
 * How far an inner shadow moves the shape's inverse, in slide space (ECMA-376 Part 1, 20.1.8,
 * CT_InnerShadowEffect: `dist` at `dir`). The shadow shows on the side that `dir` points to: in
 * the reference PDFs, `dir="0"` shades the right edge, `5400000` the bottom edge, and `13500000`
 * the top and left edges (`shape-effects-presets` slide 1, `effects` slide 1). So the inverse,
 * which casts the shadow, moves the opposite way.
 *
 * @internal
 */
export function innerShadowShift(dist: number, dir: number): { dx: number; dy: number } {
  const r = dir * DEG;
  return { dx: -Math.cos(r) * dist, dy: -Math.sin(r) * dist };
}

/** Bounds of points. */
export function boundsOf(points: readonly { x: number; y: number }[]): EffectBounds {
  let x = Infinity;
  let y = Infinity;
  let r = -Infinity;
  let b = -Infinity;
  for (const p of points) {
    x = Math.min(x, p.x);
    y = Math.min(y, p.y);
    r = Math.max(r, p.x);
    b = Math.max(b, p.y);
  }
  return { x, y, r, b };
}

/** The corners of a `w`×`h` local box under a matrix. */
export function cornersUnder(m: Matrix, w: number, h: number): { x: number; y: number }[] {
  return [apply(m, 0, 0), apply(m, w, 0), apply(m, w, h), apply(m, 0, h)];
}

/** Bounds grown by a margin on every side. */
export function grow(b: EffectBounds, margin: number): EffectBounds {
  return { x: b.x - margin, y: b.y - margin, r: b.r + margin, b: b.b + margin };
}

/**
 * The gradient line for a reflection fade: from the bounds' first point along `fadeDir` to its
 * last, so offsets 0 and 1 are the bounds' extremes in that direction.
 *
 * @internal
 */
export function fadeLine(
  bounds: EffectBounds,
  fadeDir: number
): { x1: number; y1: number; x2: number; y2: number } {
  const dx = Math.cos(fadeDir * DEG);
  const dy = Math.sin(fadeDir * DEG);
  const corners = [
    { x: bounds.x, y: bounds.y },
    { x: bounds.r, y: bounds.y },
    { x: bounds.r, y: bounds.b },
    { x: bounds.x, y: bounds.b },
  ];
  const proj = corners.map((p) => p.x * dx + p.y * dy);
  const lo = proj.reduce((a, v) => Math.min(a, v), Infinity);
  const hi = proj.reduce((a, v) => Math.max(a, v), -Infinity);
  const cx = (bounds.x + bounds.r) / 2;
  const cy = (bounds.y + bounds.b) / 2;
  const mid = cx * dx + cy * dy;
  return {
    x1: cx + (lo - mid) * dx,
    y1: cy + (lo - mid) * dy,
    x2: cx + (hi - mid) * dx,
    y2: cy + (hi - mid) * dy,
  };
}

/**
 * Where a reflection paints and how it fades. The SVG renderer and the PDF writer
 * (`@pptx-editor.dev/pptx-to-pdf`) both paint from it, so their reflections agree.
 *
 * @internal
 */
export interface ReflectionFade {
  /** The slide-space matrix that moves the shape's content to its reflection. */
  readonly matrix: Matrix;
  /** The fade's gradient line: offset 0 at (`x1`, `y1`), offset 1 at (`x2`, `y2`). */
  readonly line: {
    readonly x1: number;
    readonly y1: number;
    readonly x2: number;
    readonly y2: number;
  };
  /** The offset of the first fade stop, in [0, 1]. */
  readonly start: number;
  /** The offset of the last fade stop, in [`start`, 1]. */
  readonly end: number;
}

const unit = (v: number): number => Math.max(0, Math.min(1, v));

/**
 * The placement and fade of a reflection (ECMA-376 Part 1, 20.1.8, CT_ReflectionEffect, as the
 * cascade resolves it) of a `w` × `h` shape whose local box maps to the slide
 * through `shape`. The fade runs across the reflected shape box, in the slide's frame, or in
 * the shape's when `rotWithShape` is set. The last stop is at the later of the two positions.
 *
 * @internal
 */
export function reflectionFade(
  ref: NonNullable<ResolvedEffects['reflection']>,
  shape: Matrix,
  w: number,
  h: number
): ReflectionFade {
  const matrix = effectMatrix(ref, shape, w, h);
  const box = boundsOf(cornersUnder(multiply(matrix, shape), w, h));
  const turn = ref.rotWithShape ? (Math.atan2(shape[1], shape[0]) * 180) / Math.PI : 0;
  const start = unit(ref.startPos);
  return {
    matrix,
    line: fadeLine(box, ref.fadeDir + turn),
    start,
    end: Math.max(start, unit(ref.endPos)),
  };
}
