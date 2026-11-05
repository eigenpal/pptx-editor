// The geometry of outer shadows and reflections (design D2–D4 of `effects-fidelity`). Pure math
// over layout matrices: no DOM. Paint uses the results to transform and filter a copy of a
// shape's content.

import type { EffectAlign, EffectPlacement } from '../cascade/index.ts';
import { apply, invert, multiply, type Matrix } from '../layout/index.ts';

const DEG = Math.PI / 180;

/**
 * The linear transfer of a luminance effect (`a:lum`) on each sRGB channel in [0, 1]: brightness
 * shifts, contrast scales around mid-gray (ECMA-376 Part 1, 20.1.8.42). The SVG paint and the
 * Corrections gallery previews use this one copy. @internal
 */
export function lumTransfer(
  bright: number,
  contrast: number
): { readonly slope: number; readonly intercept: number } {
  const slope = contrast >= 0 ? 1 / Math.max(0.0001, 1 - contrast) : 1 + contrast;
  return { slope, intercept: bright + 0.5 - 0.5 * slope };
}

/** One pixel at 96 dpi, in EMU: filter regions keep it so they never clip antialiased edges. */
export const PX = 9525;

/**
 * Gaussian σ per EMU of a blur radius. ECMA-376 Part 1, 20.1.8 gives `blurRad` (CT_OuterShadowEffect,
 * CT_InnerShadowEffect) and `rad` (CT_GlowEffect, CT_SoftEdgesEffect) as a radius, but it does
 * not give the kernel. Reflections use `REFLECTION_BLUR_SIGMA_PER_RADIUS`. The reference PDFs show a Gaussian whose tail ends at
 * the radius, so σ is a third of it (3σ = radius). Measured as edge profiles at 2 px per point:
 * the `effects` fixture (slide 1) and `shape-effects-presets` (slides 1 and 2), with outer and
 * inner shadows of 10 and 18 px, glows of 10 to 36 px, and soft edges of 5 and 20 px.
 */
export const BLUR_SIGMA_PER_RADIUS = 1 / 3;

/** The Gaussian σ for a blur radius. */
export function blurSigma(blurRad: number): number {
  return Math.max(0, blurRad) * BLUR_SIGMA_PER_RADIUS;
}

/**
 * Gaussian σ per EMU of a reflection's `blurRad` (CT_ReflectionEffect). UNMEASURED: the fixture
 * reflections blur by 1 to 6 px, and at that size the reference raster of picture edges differs
 * from the browser's by more than the blur does (`picture-editing` slide 1). So reflections keep
 * σ = `blurRad` / 2 until a fixture with a larger reflection blur measures it.
 */
export const REFLECTION_BLUR_SIGMA_PER_RADIUS = 0.5;

/** The Gaussian σ for a reflection's blur radius. */
export function reflectionBlurSigma(blurRad: number): number {
  return Math.max(0, blurRad) * REFLECTION_BLUR_SIGMA_PER_RADIUS;
}

/**
 * How far the outline moves and how much it then blurs, in EMU. A positive `spread` grows the
 * outline (a dilation), a negative one shrinks it (an erosion).
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
 */
export function innerShadowShift(dist: number, dir: number): { dx: number; dy: number } {
  const r = dir * DEG;
  return { dx: -Math.cos(r) * dist, dy: -Math.sin(r) * dist };
}

/** Where `algn` sits on a box, as fractions of its width and height. */
export const ALIGN_POINTS: Readonly<Record<EffectAlign, readonly [number, number]>> = {
  tl: [0, 0],
  t: [0.5, 0],
  tr: [1, 0],
  l: [0, 0.5],
  ctr: [0.5, 0.5],
  r: [1, 0.5],
  bl: [0, 1],
  b: [0.5, 1],
  br: [1, 1],
};

/** Axis-aligned bounds. */
export interface Bounds {
  readonly x: number;
  readonly y: number;
  readonly r: number;
  readonly b: number;
}

/** Bounds of points. */
export function boundsOf(points: readonly { x: number; y: number }[]): Bounds {
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

/** Scale, then skew, about point `a`, then offset by `dist` at `dir`. */
function placementMatrix(e: EffectPlacement, a: { x: number; y: number }, dir: number): Matrix {
  const skewScale: Matrix = [
    e.sx,
    Math.tan(e.ky * DEG) * e.sx,
    Math.tan(e.kx * DEG) * e.sy,
    e.sy,
    0,
    0,
  ];
  const ox = Math.cos(dir * DEG) * e.dist;
  const oy = Math.sin(dir * DEG) * e.dist;
  return multiply([1, 0, 0, 1, a.x + ox, a.y + oy], multiply(skewScale, [1, 0, 0, 1, -a.x, -a.y]));
}

/**
 * The slide-space matrix that moves a shape's content to its shadow or reflection. `shape` maps
 * the shape's `w`×`h` local box to the slide.
 */
export function effectMatrix(e: EffectPlacement, shape: Matrix, w: number, h: number): Matrix {
  const [fx, fy] = ALIGN_POINTS[e.align];
  if (e.rotWithShape) {
    // In the shape's own frame: the effect turns and flips with the shape.
    const local = placementMatrix(e, { x: fx * w, y: fy * h }, e.dir);
    return multiply(shape, multiply(local, invert(shape)));
  }
  const box = boundsOf(cornersUnder(shape, w, h));
  const anchor = { x: box.x + fx * (box.r - box.x), y: box.y + fy * (box.b - box.y) };
  return placementMatrix(e, anchor, e.dir);
}

/** Bounds grown by a margin on every side. */
export function grow(b: Bounds, margin: number): Bounds {
  return { x: b.x - margin, y: b.y - margin, r: b.r + margin, b: b.b + margin };
}

/**
 * The gradient line for a reflection fade: from the bounds' first point along `fadeDir` to its
 * last, so offsets 0 and 1 are the bounds' extremes in that direction.
 */
export function fadeLine(
  bounds: Bounds,
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
