// Pure math that a second painter of the same layout (`@pptx-editor.dev/pptx-to-pdf`) and the
// effects renderer (`@pptx-editor.dev/pro/drawing`) share with this SVG paint: these rules have
// one copy, the gradient geometry and stop colors included. Internal glue between packages, not
// a public API.

import type { EffectPlacement, GradientStop, ResolvedFill, Rgba } from '../cascade/index.ts';
import type { GeometryPath, LayoutItem, Matrix, ThreeDFacet } from '../layout/index.ts';
import {
  applyBlipColorSteps,
  blipColorSteps,
  blurSigma as sigmaFor,
  effectMatrix as placementOf,
  LUMA_WEIGHTS,
  lumTransfer,
  type BlipColorStep,
} from './effect-transforms.ts';
import { boxCorners } from './item-corners.ts';
import { adjustForPathMode as adjusted, tileGeometry, type TileGeometry } from './paint-fill.ts';
import {
  circleGeometry,
  linearEnds,
  paintStops,
  rectGeometry,
  shapeGradientBands as bandsOf,
  ShapeBandBudget,
} from './paint-gradient.ts';

export { applyBlipColorSteps, blipColorSteps, LUMA_WEIGHTS, lumTransfer, tileGeometry };
export type { BlipColorStep, TileGeometry };

/**
 * Lighten or darken a color for a preset path fill mode (`a:path/@fill`), as the SVG paint does.
 * @internal
 */
export function adjustForPathMode(color: Rgba, mode: GeometryPath['fill']): Rgba {
  return adjusted(color, mode);
}

/**
 * The Gaussian σ that the paint uses for an effect's blur radius, in EMU.
 * @internal
 */
export function blurSigma(blurRad: number): number {
  return sigmaFor(blurRad);
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

/**
 * The color that a bevel and an extrusion shade: the solid fill, the middle stop of a gradient,
 * the pattern's foreground, else mid gray (a picture fill, or no fill).
 * @internal
 */
export function threeDColor(fill: ResolvedFill): Rgba {
  if (fill.type === 'solid') return fill.color;
  if ((fill.type === 'linear' || fill.type === 'path') && fill.stops.length > 0)
    return fill.stops[Math.floor(fill.stops.length / 2)]!.color;
  if (fill.type === 'pattern') return fill.fg;
  return { r: 128, g: 128, b: 128, a: 1 };
}

/**
 * Whether every coordinate and color channel of a 3-D facet is a finite number. The SVG and PDF
 * paints skip a facet that is not, so a degenerate shape never writes `NaN`.
 * @internal
 */
export function finiteFacet(f: ThreeDFacet): boolean {
  const c = f.color;
  return (
    [c.r, c.g, c.b, c.a].every(Number.isFinite) &&
    f.polygons.every((poly) => poly.every((p) => Number.isFinite(p.x) && Number.isFinite(p.y)))
  );
}

/**
 * The bands of a `path="shape"` gradient over a `w` × `h` box, as the SVG paint draws them:
 * outline copies scaled about (`cx`, `cy`), outside in, each in one color. `stops` are the fill's
 * stops after the path fill mode.
 * @internal
 */
export function shapeGradientBands(
  fill: Extract<ResolvedFill, { type: 'path' }>,
  stops: readonly GradientStop[],
  w: number,
  h: number
): { cx: number; cy: number; bands: { scale: number; color: Rgba }[] } {
  return bandsOf(fill, stops, w, h);
}

/**
 * The shape gradient band work of one slide, with the SVG paint's limits: a shape draws outline
 * bands while `allow` returns true, and the rectangle construction after.
 * @internal
 */
export function shapeBandBudget(): { allow(paths: number, commands: number): boolean } {
  return new ShapeBandBudget();
}

/**
 * The construction of a `path="rect"` gradient over a `w` × `h` box, as the SVG paint draws it:
 * the first stop on the focus rectangle (`l`, `t`, `r`, `b`), the last on the outer rectangle
 * (`x0`, `y0`, `x1`, `y1`), four trapezoids between them.
 * @internal
 */
export function rectGradientGeometry(
  fill: Extract<ResolvedFill, { type: 'path' }>,
  w: number,
  h: number
): { l: number; t: number; r: number; b: number; x0: number; y0: number; x1: number; y1: number } {
  return rectGeometry(fill, w, h);
}

/**
 * The slide-space corners of an item's box, geometry, and outline: the bounds that the SVG
 * paint's body effects (soft edge, inner shadow) read.
 * @internal
 */
export function effectBoxCorners(
  item: Extract<LayoutItem, { kind: 'shape' } | { kind: 'picture' }>
): { x: number; y: number }[] {
  return boxCorners(item);
}
