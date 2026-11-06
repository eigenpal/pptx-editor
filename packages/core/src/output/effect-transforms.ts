// The geometry of outer shadows (design D2–D4 of `effects-fidelity`). Pure math over layout
// matrices: no DOM. Paint uses the results to transform and filter a copy of a shape's content.
// The rules of glow, inner shadow, reflection, and soft edge are in `@pptx-editor.dev/pro/drawing`
// (`packages/pro/src/drawing/effects-math.ts`), which reuses `effectMatrix` and `blurSigma`.

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
 * not give the kernel. The reference PDFs show a Gaussian whose tail ends at the radius, so σ
 * is a third of it (3σ = radius). Measured as edge profiles at 2 px per point: the `effects`
 * fixture (slide 1) and `shape-effects-presets` (slides 1 and 2), with outer and inner shadows
 * of 10 and 18 px, glows of 10 to 36 px, and soft edges of 5 and 20 px.
 */
export const BLUR_SIGMA_PER_RADIUS = 1 / 3;

/** The Gaussian σ for a blur radius. */
export function blurSigma(blurRad: number): number {
  return Math.max(0, blurRad) * BLUR_SIGMA_PER_RADIUS;
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
