// Pure math that a second painter of the same layout (`@pptx-editor.dev/pptx-to-pdf`) shares with
// this SVG paint: these rules have one copy, the gradient geometry and stop colors included.
// Internal glue between packages, not a public API.

import type { EffectPlacement, GradientStop, ResolvedFill, Rgba } from '../cascade/index.ts';
import type { GeometryPath, Matrix } from '../layout/index.ts';
import {
  blurSigma as sigmaFor,
  effectMatrix as placementOf,
  glowSpread as glowFor,
  innerShadowShift as innerShiftFor,
  lumTransfer,
  type SpreadBlur,
} from './effect-transforms.ts';
import { adjustForPathMode as adjusted } from './paint-fill.ts';
import { circleGeometry, linearEnds, paintStops } from './paint-gradient.ts';

export type { SpreadBlur };
export { lumTransfer };

/**
 * Lighten or darken a color for a preset path fill mode (`a:path/@fill`), as the SVG paint does.
 * @internal
 */
export function adjustForPathMode(color: Rgba, mode: GeometryPath['fill']): Rgba {
  return adjusted(color, mode);
}

/**
 * The Gaussian σ that the SVG paint uses for an effect's blur radius, in EMU.
 * @internal
 */
export function blurSigma(blurRad: number): number {
  return sigmaFor(blurRad);
}

/**
 * How far the SVG paint grows a glow's outline and how much it blurs it, in EMU.
 * @internal
 */
export function glowSpread(rad: number): SpreadBlur {
  return glowFor(rad);
}

/**
 * How far the SVG paint moves an inner shadow's inverse, in slide space: opposite `dir`, so the
 * shadow shows on the side that `dir` points to.
 * @internal
 */
export function innerShadowShift(dist: number, dir: number): { dx: number; dy: number } {
  return innerShiftFor(dist, dir);
}

/**
 * The slide-space matrix that moves a shape's content to its shadow or reflection. `shape` maps
 * the shape's `w` × `h` local box to the slide.
 * @internal
 */
export function effectMatrix(e: EffectPlacement, shape: Matrix, w: number, h: number): Matrix {
  return placementOf(e, shape, w, h);
}

/**
 * The end points of a linear gradient over a `w` × `h` box at the origin, as the SVG paint draws
 * them (`a:lin`, with `scaled`, as reference PDFs show).
 * @internal
 */
export function linearGradientEnds(
  angle: number,
  scaled: boolean,
  w: number,
  h: number
): { x1: number; y1: number; x2: number; y2: number } {
  return linearEnds(angle, scaled, { x: 0, y: 0, w, h });
}

/**
 * The circle and focal point of a circular path gradient over a `w` × `h` box, as the SVG paint
 * draws it, as reference PDFs show: the first stop at (`fx`, `fy`), the last on the circle.
 * @internal
 */
export function circleGradientGeometry(
  fill: Extract<ResolvedFill, { type: 'path' }>,
  w: number,
  h: number
): { cx: number; cy: number; r: number; fx: number; fy: number } {
  return circleGeometry(fill, w, h);
}

/**
 * The stops that the SVG paint draws: a two-stop gradient bends its colors, as reference PDFs show.
 * @internal
 */
export function gradientPaintStops(stops: readonly GradientStop[]): GradientStop[] {
  return paintStops(stops);
}
