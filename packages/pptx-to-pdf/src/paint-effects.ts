/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Reflections, outer shadows, and glow as PDF paint, under the shape, and inner shadows over its
// geometry. The placement is the editor's (`effectMatrix` and `blurSigma`, shared with
// `packages/core/src/output/`; `glowSpread`, `innerShadowShift`, `reflectionFade`, and
// `reflectionBlurSigma`, shared with `@pptx-editor.dev/pro/drawing`). Glow, inner shadow, and
// reflection paint only while the Pro drawing renderers are installed, as in the SVG paint;
// opening an export session always installs them. The blur is an APPROXIMATION: the writer has
// no blur filter, so a blurred silhouette is drawn as a few layers, each the silhouette grown by a ring and painted as
// one transparency group, at alphas that add up to the Gaussian edge (an approximation).
//
// The silhouette has parts: filled paths, outlines of unfilled paths, and the text's glyph
// outlines (when the shape has no fill). Each part is one form, traced once in slide space and
// stroked with the line width it inherits, so a layer is a line width and a form per part. Wide
// parts (filled paths) take the Gaussian edge of a half-plane; outlines and text take the
// Gaussian of a band of their own width, so a thin feature does not cast a shadow as dark as a
// wide one. An inner shadow is the inverse of the filled paths (a large box minus them), moved
// and grown the same way and clipped to them.

import {
  concatTransformationMatrix,
  drawObject,
  LineCapStyle,
  LineJoinStyle,
  PDFNumber,
  PDFOperator,
  PDFOperatorNames,
  popGraphicsState,
  pushGraphicsState,
  setDashPattern,
  setFillingGrayscaleColor,
  setFillingRgbColor,
  setLineCap,
  setLineJoin,
  setLineWidth,
  setStrokingGrayscaleColor,
  setStrokingRgbColor,
  stroke,
  type PDFName,
  type PDFRef,
} from 'pdf-lib';
import type { ResolvedEffects, ResolvedFill, Rgba, Xfrm } from '@pptx-editor.dev/core/cascade';
import {
  apply,
  ARROW_MITER_LIMIT,
  dashArray,
  drawingRenderers,
  hasCharacters,
  invert,
  multiply,
  planStroke,
  strokeWidth,
  type GeometryPath,
  type LayoutItem,
  type Matrix,
  type PathCommand,
  type StrokePlan,
} from '@pptx-editor.dev/core/layout';
import {
  blurSigma,
  effectBoxCorners,
  effectMatrix,
  threeDColor,
} from '@pptx-editor.dev/core/output';
import {
  glowSpread,
  innerShadowShift,
  MAX_MORPH_SPREAD,
  reflectionBlurSigma,
  reflectionFade,
  softEdgeSpread,
} from '@pptx-editor.dev/pro/drawing';
import { flatColor, ownFill, type FillScope } from './paint-fill.ts';
import { fillTextOutlines } from './paint-text.ts';
import { outerEdges } from './outer-edges.ts';
import {
  clipEvenOdd,
  EMU_PER_PT,
  shadingFill,
  endPath,
  fillEvenOdd,
  fillEvenOddAndStroke,
  matrixScale,
  rgb,
  tracePath,
  type PageOps,
} from './page-ops.ts';

export type Drawable = Extract<LayoutItem, { kind: 'shape' } | { kind: 'picture' }>;

/** A shape, picture, or group whose effects place copies of its content by its box. */
type EffectHost = {
  readonly element: { readonly effects?: ResolvedEffects; readonly xfrm: Xfrm };
  readonly matrix: Matrix;
};

/** A blur below half a point is drawn as a hard edge: it is under one output pixel. */
export const MIN_SIGMA = 6_350;

/** A typical stem of a glyph, as a fraction of the font size: the width of a text silhouette. */
const STEM_PER_EM = 0.1;

/** Whether the Pro drawing renderers are installed, so glow, inner shadow, and reflection paint. */
export function proEffects(): boolean {
  return drawingRenderers().effects !== undefined;
}

/** Whether the item's fill, outline, or text has any transparency. */
function translucent(item: Drawable, scope: FillScope): boolean {
  const el = item.element;
  const fill = ownFill(el.fill, scope).fill;
  const see = (f: typeof fill): boolean =>
    (f.type === 'solid' && f.color.a < 1) ||
    ((f.type === 'linear' || f.type === 'path') && f.stops.some((s) => s.color.a < 1));
  if (see(fill) || (el.line && see(ownFill(el.line.fill, scope).fill))) return true;
  if (item.kind !== 'shape' || !item.text) return false;
  return item.text.lines.some((l) => l.runs.some((r) => see(ownFill(r.props.fill, scope).fill)));
}

/** Report the effects this writer draws approximately or not at all, when they draw something. */
export function effectDiagnostics(item: Drawable, p: PageOps, scope: FillScope): void {
  const el = item.element;
  const effects = el.effects;
  if (!effects) return;
  const fills = item.kind === 'picture' || ownFill(el.fill, scope).fill.type !== 'none';
  // The parts of the silhouette (`silhouette`): filled paths, a drawn outline, and text.
  const filled = fills ? item.geometry.paths.filter((path) => path.fill !== 'none') : [];
  const lineFill = el.line?.fill.type;
  const outlined = lineFill === 'solid' || lineFill === 'linear' || lineFill === 'path';
  const hasText = item.kind === 'shape' && item.text !== undefined && hasCharacters(item.text);
  const visible = filled.length > 0 || outlined || hasText;
  const pro = proEffects();
  const { outerShadow, glow, innerShadow, softEdge, reflection } = effects;
  const shadowDrawn = visible && outerShadow !== undefined && outerShadow.color.a > 0;
  const glowDrawn = pro && visible && glow !== undefined && glow.radius > 0 && glow.color.a > 0;
  if (shadowDrawn && blurSigma(outerShadow.blur) >= MIN_SIGMA)
    p.diag.add('effect.shadow-blur', 'approximation', 'Shadow blur is drawn as layered rings.');
  if (glowDrawn) p.diag.add('effect.glow', 'approximation', 'Glow is drawn as layered rings.');
  if ((shadowDrawn || glowDrawn) && (item.kind === 'picture' || translucent(item, scope)))
    p.diag.add(
      'effect.silhouette',
      'approximation',
      'Shadows and glow use an opaque silhouette: picture and fill transparency do not shape them.'
    );
  const m = item.matrix;
  const inner = pro && innerShadow !== undefined && innerShadow.color.a > 0;
  // `drawInnerShadow` shades the filled paths of a shape with an area.
  const innerDrawn = inner && filled.length > 0 && Math.abs(m[0] * m[3] - m[1] * m[2]) >= 1e-9;
  if (innerDrawn && blurSigma(innerShadow.blur) >= MIN_SIGMA)
    p.diag.add(
      'effect.inner-shadow-blur',
      'approximation',
      'Inner shadow blur is drawn as layered rings.'
    );
  // The editor shades from the edge of everything the shape paints; the writer starts at the
  // fill. They differ for an outline wider than one point, an unfilled outline, and paths that
  // overlap (the even-odd inverse makes them holes).
  const wide = outlined && (el.line?.width ?? 0) > EMU_PER_PT;
  if (inner && ((outlined && (!fills || wide)) || filled.length > 1))
    p.diag.add(
      'effect.inner-shadow',
      'approximation',
      'Inner shadows start at the fill: outlines and overlapping paths are shaded differently.'
    );
  if (pro && softEdge && softEdge.radius > 0)
    p.diag.add(
      'effect.soft-edge',
      'approximation',
      'Soft edges fade by the distance to the painted edge: corners round.'
    );
  if (pro && effects.blur && effects.blur.radius > 0)
    p.diag.add(
      'effect.blur',
      'approximation',
      'Blur is drawn as an edge fade and layered rings: text and detail inside the shape stay sharp.'
    );
  if (
    pro &&
    reflection &&
    (reflection.startAlpha > 0 || reflection.endAlpha > 0) &&
    reflectionBlurSigma(reflection.blur) >= MIN_SIGMA
  )
    p.diag.add('effect.reflection-blur', 'approximation', 'Reflections are drawn without blur.');
}

/** The standard normal distribution function (Abramowitz and Stegun 7.1.26 for erf). */
export function phi(x: number): number {
  const z = Math.abs(x) / Math.SQRT2;
  const t = 1 / (1 + 0.3275911 * z);
  const erf =
    1 -
    ((((1.061405429 * t - 1.453152027) * t + 1.421413741) * t - 0.284496736) * t + 0.254829592) *
      t *
      Math.exp(-z * z);
  return x >= 0 ? (1 + erf) / 2 : (1 - erf) / 2;
}

/** One point: a ring narrower than this is under one output pixel. */
const MIN_RING = 12_700;
/** At most this many rings, so a large blur stays a bounded number of layers. */
const MAX_RINGS = 12;

/** Ring widths from 0 to `extent`, about `step` apart, at least a point apart: at most 13 layers. */
export function rings(extent: number, step: number): number[] {
  const count = Math.max(1, Math.min(MAX_RINGS, Math.round(extent / Math.max(step, MIN_RING))));
  return Array.from({ length: count + 1 }, (_, i) => (extent * i) / count);
}

/**
 * Layers that approximate an alpha profile across the silhouette edge. Layer `k` covers the
 * silhouette grown by `rings[k]` (the first ring is 0). Painted in order, the layers give the
 * alpha `profile(d)` at the middle of each band between rings, and `inside` within the edge.
 */
export function layers(
  rings: readonly number[],
  inside: number,
  profile: (d: number) => number
): { grow: number; alpha: number }[] {
  const n = rings.length;
  // Wanted coverage of each band: band 0 is inside the edge, band j lies between rings j-1 and j.
  const want = rings.map((r, j) =>
    j === 0 ? inside : Math.max(0, Math.min(1, profile((rings[j - 1] + r) / 2)))
  );
  const alpha = new Array<number>(n).fill(0);
  let outer = 0;
  for (let j = n - 1; j >= 0; j -= 1) {
    const covered = Math.max(want[j], outer);
    alpha[j] = outer >= 1 ? 0 : 1 - (1 - covered) / (1 - outer);
    outer = covered;
  }
  return rings.map((grow, j) => ({ grow, alpha: alpha[j] })).filter((l) => l.alpha > 0.002);
}

/**
 * The layers of a blurred silhouette with alpha `a`, Gaussian σ `sigma`, and part width `width`
 * (Infinity for a filled area). Outside the edge, the profile is the blurred band's; inside, the
 * level is the profile at half the shadow's `offset`, where the part that shows beside the shape
 * lies (the centre of a band).
 */
export function blurLayers(
  a: number,
  sigma: number,
  width: number,
  offset: number
): { grow: number; alpha: number }[] {
  if (sigma < MIN_SIGMA) return [{ grow: 0, alpha: a }];
  const wide = !Number.isFinite(width);
  // Distance d outside the edge: a half-plane gives a(1 − Φ(d/σ)); a band of width W gives
  // a(Φ((W + d)/σ) − Φ(d/σ)).
  const profile = (d: number): number =>
    wide ? a * (1 - phi(d / sigma)) : a * (phi((width + d) / sigma) - phi(d / sigma));
  const inside = wide
    ? a * phi(Math.min(offset, 4 * sigma) / (2 * sigma))
    : a * (2 * phi(width / (2 * sigma)) - 1);
  return layers(rings(2.5 * sigma, sigma / 4), inside, profile);
}

/** Path commands under a matrix: affine maps keep lines lines and Béziers Béziers. */
function pathUnder(commands: readonly PathCommand[], m: Matrix): PathCommand[] {
  const at = (x: number, y: number) => apply(m, x, y);
  return commands.map((c) => {
    if (c.op === 'Z') return c;
    const end = at(c.x, c.y);
    if (c.op === 'C') {
      const [a, b] = [at(c.x1, c.y1), at(c.x2, c.y2)];
      return { ...c, x1: a.x, y1: a.y, x2: b.x, y2: b.y, x: end.x, y: end.y };
    }
    if (c.op === 'Q') {
      const a = at(c.x1, c.y1);
      return { ...c, x1: a.x, y1: a.y, x: end.x, y: end.y };
    }
    return { ...c, x: end.x, y: end.y };
  });
}

/** The bounds of every form here: the slide and its surroundings, in slide EMU. */
export function formBox(p: PageOps): number[] {
  const { width, height } = p.page.getSize();
  const [w, h] = [width * EMU_PER_PT, height * EMU_PER_PT];
  return [-w, -h, 2 * w, 2 * h];
}

/** One form of ready operators, in slide EMU, drawn in the caller's color and line width. */
async function form(p: PageOps, draw: () => void): Promise<PDFName> {
  const ops = await p.capture(async () => {
    p.push(setLineJoin(LineJoinStyle.Round), setLineCap(LineCapStyle.Round));
    draw();
  });
  return p.name('XObject', p.form(ops, formBox(p)));
}

/**
 * One part of a silhouette: `grown` is filled and stroked with the inherited line width,
 * `flat` (when present) only filled, for a layer that does not grow. `width(grow)` is the line
 * width that grows the part by `grow` slide EMU.
 */
export interface Part {
  readonly grown: PDFName;
  readonly flat?: PDFName;
  readonly width: (grow: number) => number;
}

/** A class of silhouette parts with one blur profile: wide areas, or bands of one width. */
export interface PartClass {
  readonly parts: readonly Part[];
  /** Infinity for areas; the band width in slide EMU for outlines and text. */
  readonly bandWidth: number;
}

/**
 * The silhouette of an item under `m` (the effect transform), as classes of parts. Filled paths
 * are one class (with their outline when it is drawn); outlines of unfilled paths are another,
 * a band of the line width; the text is a third, a band of a glyph stem, when the shape has no
 * fill to cover it.
 */
export async function silhouette(
  item: Drawable,
  m: Matrix,
  p: PageOps,
  scope: FillScope
): Promise<PartClass[]> {
  const el = item.element;
  const full = multiply(m, item.matrix);
  // A line's width as the shape draws it, under the effect transform.
  const lineScale = matrixScale(full) / matrixScale(item.matrix);
  const fills = item.kind === 'picture' || ownFill(el.fill, scope).fill.type !== 'none';
  const paintLine = item.paintLine ?? el.line;
  const line = paintLine && flatColor(paintLine.fill, scope, p, 'lines') ? paintLine : undefined;
  const lineWidth = line ? (line.width || 9525) * lineScale : 0;
  const trace = (paths: readonly GeometryPath[], close: PDFOperator) => () => {
    for (const path of paths) {
      tracePath(pathUnder(path.commands, full), p);
      p.push(close);
    }
  };
  // A stroke as the shape draws it (`planStroke`): an open path stops short under its line ends.
  const planOf = (path: GeometryPath): StrokePlan | undefined =>
    line && item.kind === 'shape'
      ? planStroke(path.commands, line, matrixScale(item.matrix))
      : undefined;
  // The line ends cast their own shadow: filled heads as areas, open arrows as strokes.
  const endParts = async (plans: readonly (StrokePlan | undefined)[]): Promise<Part[]> => {
    const all = plans.flatMap((plan) => plan?.ends ?? []);
    const heads = all.filter((e) => e.kind === 'fill');
    const arrows = all.filter((e) => e.kind === 'stroke');
    const ends = (close: PDFOperator) => (): void => {
      for (const head of heads) {
        tracePath(pathUnder(head.commands, full), p);
        p.push(close);
      }
    };
    const parts: Part[] = [];
    if (heads.length)
      parts.push({
        grown: await form(p, ends(fillEvenOddAndStroke())),
        flat: await form(p, ends(fillEvenOdd())),
        width: (grow) => 2 * grow,
      });
    if (arrows.length)
      parts.push({
        grown: await form(p, () => {
          p.push(
            setLineJoin(LineJoinStyle.Miter),
            PDFOperator.of(PDFOperatorNames.SetLineMiterLimit, [PDFNumber.of(ARROW_MITER_LIMIT)])
          );
          for (const arrow of arrows) tracePath(pathUnder(arrow.commands, full), p);
          p.push(stroke());
        }),
        width: (grow) => lineWidth + 2 * grow,
      });
    return parts;
  };
  const classes: PartClass[] = [];
  const filled = fills ? item.geometry.paths.filter((path) => path.fill !== 'none') : [];
  const outlined = filled.filter((path) => path.stroke && line);
  const bare = filled.filter((path) => !(path.stroke && line));
  const area: Part[] = [];
  if (outlined.length)
    area.push(
      {
        grown: await form(p, trace(outlined, fillEvenOddAndStroke())),
        width: (grow) => lineWidth + 2 * grow,
      },
      ...(await endParts(outlined.map(planOf)))
    );
  if (bare.length)
    area.push({
      grown: await form(p, trace(bare, fillEvenOddAndStroke())),
      flat: await form(p, trace(bare, fillEvenOdd())),
      width: (grow) => 2 * grow,
    });
  if (area.length) classes.push({ parts: area, bandWidth: Infinity });
  const lines = line
    ? item.geometry.paths.filter((path) => path.stroke && !filled.includes(path))
    : [];
  if (line && lines.length) {
    // Open lines stop short under their line ends, and dashes keep their gaps. The dashes take
    // round caps, as every part of the silhouette does, so each dash keeps its drawn length.
    const plans = lines.map(planOf);
    const dashes = dashArray({ ...line, cap: 'rnd' }, lineWidth);
    const strokes = (): void => {
      if (dashes) p.push(setDashPattern([...dashes], 0));
      lines.forEach((path, i) => {
        tracePath(pathUnder(plans[i]?.commands ?? path.commands, full), p);
        p.push(stroke());
      });
    };
    classes.push({
      parts: [
        { grown: await form(p, strokes), width: (g) => lineWidth + 2 * g },
        ...(await endParts(plans)),
      ],
      bandWidth: lineWidth,
    });
  }
  if (!fills && item.kind === 'shape' && item.text && item.textMatrix && hasCharacters(item.text)) {
    const { text, textMatrix: t } = item;
    const inner = multiply(m, t);
    const k = matrixScale(inner);
    const inText = (draw: () => void) => (): void => {
      p.push(
        concatTransformationMatrix(inner[0], inner[1], inner[2], inner[3], inner[4], inner[5])
      );
      draw();
    };
    let size = 0;
    for (const l of text.lines) for (const r of l.runs) size = Math.max(size, r.sizePt);
    classes.push({
      parts: [
        {
          // The text form strokes in the text frame's units, under the effect transform.
          grown: await form(
            p,
            inText(() => fillTextOutlines(text, p, scope, true))
          ),
          flat: await form(
            p,
            inText(() => fillTextOutlines(text, p, scope))
          ),
          width: (grow) => (2 * grow) / k,
        },
      ],
      bandWidth: STEM_PER_EM * size * EMU_PER_PT * k,
    });
  }
  return classes;
}

/** Paint the layers of one part class, each as one transparency group, so overlaps never darken. */
async function drawLayers(
  parts: readonly Part[],
  color: Rgba,
  list: readonly { grow: number; alpha: number }[],
  p: PageOps
): Promise<void> {
  for (const layer of list) {
    const ops = await p.capture(async () => {
      p.push(setFillingRgbColor(...rgb(color)), setStrokingRgbColor(...rgb(color)));
      for (const part of parts) {
        if (layer.grow === 0 && part.flat) p.push(drawObject(part.flat));
        else p.push(setLineWidth(part.width(layer.grow)), drawObject(part.grown));
      }
    });
    const group = p.form(ops, formBox(p), { S: 'Transparency' });
    p.push(pushGraphicsState());
    p.alpha(Math.round(layer.alpha * 1000) / 1000, 1);
    p.push(drawObject(p.name('XObject', group)), popGraphicsState());
  }
}

/**
 * Cast the outer shadows of `effects` (a preset shadow's second copy under its first). `matrix`
 * maps the `w`x`h` box of the shape or group to the slide; `classesUnder` gives the silhouette
 * under an effect transform.
 */
export async function castShadows(
  effects: ResolvedEffects,
  matrix: Matrix,
  w: number,
  h: number,
  face: Matrix,
  classesUnder: (m: Matrix) => Promise<PartClass[]>,
  p: PageOps
): Promise<void> {
  for (const shadow of [effects.secondShadow, effects.outerShadow]) {
    if (!shadow || shadow.color.a <= 0) continue;
    const m = multiply(face, effectMatrix(shadow, matrix, w, h));
    const sigma = blurSigma(shadow.blur);
    for (const cls of await classesUnder(m))
      await drawLayers(
        cls.parts,
        shadow.color,
        blurLayers(shadow.color.a, sigma, cls.bandWidth, shadow.dist),
        p
      );
  }
}

/** Cast a glow around the silhouette `classes`, as the SVG glow (`glowSpread`). */
export async function castGlow(
  glow: ResolvedEffects['glow'],
  classes: () => Promise<PartClass[]>,
  p: PageOps
): Promise<void> {
  if (!glow || glow.radius <= 0 || glow.color.a <= 0 || !proEffects()) return;
  // The silhouette grown by `spread` and blurred by sigma. The inside level is the profile at
  // the edge, a * phi(spread / sigma).
  const { spread, sigma } = glowSpread(glow.radius);
  const a = glow.color.a;
  const list = layers(
    rings(spread + 3 * sigma, sigma / 2),
    a * phi(spread / sigma),
    (d) => a * (1 - phi((d - spread) / sigma))
  );
  for (const cls of await classes()) await drawLayers(cls.parts, glow.color, list, p);
}

/**
 * The outer shadows of a shape or picture, painted before its content. `face` is the camera's
 * turn of the face (`ThreeDPaint.front`): the shadow is placed on the face, then turns with it,
 * as in the SVG paint.
 */
export async function drawShadows(
  item: Drawable,
  p: PageOps,
  scope: FillScope,
  face: Matrix
): Promise<void> {
  const effects = item.element.effects;
  if (!effects) return;
  const { cx: w, cy: h } = item.element.xfrm;
  await castShadows(effects, item.matrix, w, h, face, (m) => silhouette(item, m, p, scope), p);
}

/**
 * The glow of a shape or picture: after its shadows, before its extrusion and its content, as in
 * the SVG paint. `face` is the camera's turn of the face.
 */
export async function drawGlow(
  item: Drawable,
  p: PageOps,
  scope: FillScope,
  face: Matrix
): Promise<void> {
  await castGlow(item.element.effects?.glow, () => silhouette(item, face, p, scope), p);
}

/**
 * The inner shadow's offset in the shape's local space. The SVG filter moves the inverse in slide
 * space by `innerShadowShift` (opposite `dir`), so the local offset is m⁻¹ · T · m.
 */
export function innerShadowOffset(m: Matrix, dist: number, dir: number): Matrix {
  const { dx, dy } = innerShadowShift(dist, dir);
  const shift: Matrix = [1, 0, 0, 1, dx, dy];
  return multiply(invert(m), multiply(shift, m));
}

/**
 * The inner shadow, over the geometry and under the text, as the SVG `bodyFilter`: the inverse of
 * the filled paths, offset by `dist` at `dir` in slide space, blurred, in the shadow color,
 * clipped to the filled paths. Everything is traced in slide space; the caller has no matrix.
 */
export async function drawInnerShadow(item: Drawable, p: PageOps, scope: FillScope): Promise<void> {
  const inner = item.element.effects?.innerShadow;
  // The SVG filter reads the shape's alpha: an unfilled shape has no inside to shade.
  const fills = item.kind === 'picture' || ownFill(item.element.fill, scope).fill.type !== 'none';
  const paths = fills ? item.geometry.paths.filter((path) => path.fill !== 'none') : [];
  if (!inner || inner.color.a <= 0 || paths.length === 0 || !proEffects()) return;
  const m = item.matrix;
  // A shape with no area has no inside.
  if (Math.abs(m[0] * m[3] - m[1] * m[2]) < 1e-9) return;
  const shifted = multiply(m, innerShadowOffset(m, inner.dist, inner.dir));
  const sigma = blurSigma(inner.blur);
  const list = blurLayers(inner.color.a, sigma, Infinity, inner.dist);
  // The inverse is a box around the shape, its offset, and its blur.
  const { cx: w, cy: h } = item.element.xfrm;
  const corners = [apply(m, 0, 0), apply(m, w, 0), apply(m, w, h), apply(m, 0, h)];
  const margin = inner.dist + 3 * sigma + Math.max(w, h) * matrixScale(m);
  const xs = corners.map((c) => c.x);
  const ys = corners.map((c) => c.y);
  const [x0, y0] = [Math.min(...xs) - margin, Math.min(...ys) - margin];
  const [x1, y1] = [Math.max(...xs) + margin, Math.max(...ys) + margin];
  const inverse = (close: PDFOperator) => (): void => {
    tracePath(
      [
        { op: 'M', x: x0, y: y0 },
        { op: 'L', x: x1, y: y0 },
        { op: 'L', x: x1, y: y1 },
        { op: 'L', x: x0, y: y1 },
        { op: 'Z' },
      ],
      p
    );
    for (const path of paths) tracePath(pathUnder(path.commands, shifted), p);
    p.push(close);
  };
  const part: Part = {
    grown: await form(p, inverse(fillEvenOddAndStroke())),
    flat: await form(p, inverse(fillEvenOdd())),
    width: (grow) => 2 * grow,
  };
  p.push(pushGraphicsState());
  for (const path of paths) tracePath(pathUnder(path.commands, m), p);
  p.push(clipEvenOdd(), endPath());
  await drawLayers([part], inner.color, list, p);
  p.push(popGraphicsState());
}

/**
 * The reflection of a shape, picture, or group, painted before its shadow and content, as the SVG
 * renderer paints it: `drawContent` (the geometry, inner shadow, and text, in slide space) under
 * the reflection's effect transform, as one transparency group, through a luminosity soft mask
 * that fades from `startAlpha` at `startPos` to `endAlpha` at `endPos` along `fadeDir`. The blur
 * is left out (`effectDiagnostics` reports it).
 */
export async function drawReflection(
  item: EffectHost,
  p: PageOps,
  drawContent: () => Promise<void>
): Promise<void> {
  const ref = item.element.effects?.reflection;
  if (!ref || (ref.startAlpha <= 0 && ref.endAlpha <= 0) || !proEffects()) return;
  const { cx: w, cy: h } = item.element.xfrm;
  // The SVG renderer's placement and fade, so the two reflections agree.
  const { matrix: m, line, start, end: last } = reflectionFade(ref, item.matrix, w, h);
  // A transform with no area paints nothing.
  if (Math.abs(m[0] * m[3] - m[1] * m[2]) < 1e-12) return;
  const content = await p.capture(async () => {
    p.push(pushGraphicsState(), concatTransformationMatrix(m[0], m[1], m[2], m[3], m[4], m[5]));
    await drawContent();
    p.push(popGraphicsState());
  });
  const group = p.form(content, formBox(p), { S: 'Transparency' });
  // Equal positions are a hard step; the shading needs a line of some length.
  const end = Math.max(start + 1e-6, last);
  const at = (t: number): [number, number] => [
    line.x1 + (line.x2 - line.x1) * t,
    line.y1 + (line.y2 - line.y1) * t,
  ];
  const clamp = (a: number): number => Math.max(0, Math.min(1, a));
  const ctx = p.res.doc.context;
  const shading = ctx.obj({
    ShadingType: 2,
    ColorSpace: 'DeviceGray',
    Coords: [...at(start), ...at(end)],
    Function: {
      FunctionType: 2,
      Domain: [0, 1],
      C0: [clamp(ref.startAlpha)],
      C1: [clamp(ref.endAlpha)],
      N: 1,
    },
    Extend: [true, true],
  });
  p.res.work.reserveContent(shading.sizeInBytes());
  const mask = p.form([shadingFill(p.name('Shading', ctx.register(shading)))], formBox(p), {
    S: 'Transparency',
    CS: 'DeviceGray',
  });
  p.push(pushGraphicsState());
  p.state(`smask:${mask.toString()}`, () =>
    ctx.obj({ Type: 'ExtGState', SMask: { Type: 'Mask', S: 'Luminosity', G: mask } })
  );
  p.push(drawObject(p.name('XObject', group)), popGraphicsState());
}

/** At most this many bands across a soft edge: each is one stroke of the mask. */
const MAX_SOFT_EDGE_BANDS = 48;
/** The narrowest band of a soft edge, half a point: a narrower one is under one output pixel. */
const MIN_SOFT_EDGE_BAND = 6_350;

/**
 * The alpha of a soft edge as bands of the distance inside the outline (`softEdgeSpread`): the
 * outline shrunk by `inset`, blurred with Gaussian σ `sigma`. Band `k` reaches `reach[k]` inside
 * the outline and takes `alpha[k]`, the profile at its middle; beyond the last band the alpha is 1.
 */
export function softEdgeBands(inset: number, sigma: number): { reach: number[]; alpha: number[] } {
  const extent = inset + 3 * sigma;
  if (!(extent > 0)) return { reach: [], alpha: [] };
  const count = Math.max(
    1,
    Math.min(MAX_SOFT_EDGE_BANDS, Math.ceil(extent / Math.max(sigma / 6, MIN_SOFT_EDGE_BAND)))
  );
  const profile = (d: number): number =>
    sigma < MIN_SIGMA ? (d >= inset ? 1 : 0) : phi((d - inset) / sigma);
  const reach = Array.from({ length: count }, (_, k) => (extent * (k + 1)) / count);
  const alpha = reach.map((r, k) => profile((r + (k ? reach[k - 1] : 0)) / 2));
  return { reach, alpha };
}

/**
 * What a soft edge fades, as the SVG renderer's body filter reads it (`SourceAlpha`): the filled
 * paths, and the outline at `width` in slide space (a picture's border lies outside the picture:
 * twice its width, clipped outside). `rimmed` are the filled paths whose outline lies over their
 * edge, so their own edge is `width / 2` inside the painted edge.
 */
export interface SoftEdgeSilhouette {
  readonly filled: readonly PathCommand[][];
  readonly rimmed: readonly boolean[];
  readonly outlines: readonly PathCommand[][];
  readonly width: number;
}

export function softEdgeSilhouette(
  item: Drawable,
  scope: FillScope,
  groupScale: number
): SoftEdgeSilhouette {
  const el = item.element;
  const m = item.matrix;
  const picture = item.kind === 'picture';
  const fills = picture || ownFill(el.fill, scope).fill.type !== 'none';
  const line = item.paintLine ?? el.line;
  const lineFill = line ? ownFill(line.fill, scope).fill.type : 'none';
  const drawn = !!line && (lineFill === 'solid' || lineFill === 'linear' || lineFill === 'path');
  const { cx: w, cy: h } = el.xfrm;
  // A picture with no geometry shows in its box, as `drawGeometry` draws it.
  const box: GeometryPath = {
    commands: [
      { op: 'M', x: 0, y: 0 },
      { op: 'L', x: w, y: 0 },
      { op: 'L', x: w, y: h },
      { op: 'L', x: 0, y: h },
      { op: 'Z' },
    ],
    fill: 'norm',
    stroke: true,
  } as GeometryPath;
  const paths = picture && item.geometry.paths.length === 0 ? [box] : item.geometry.paths;
  const filledPaths = fills ? paths.filter((path) => path.fill !== 'none') : [];
  const stroked = drawn ? paths.filter((path) => picture || path.stroke) : [];
  const width = drawn ? strokeWidth(line, groupScale) * matrixScale(m) * (picture ? 2 : 1) : 0;
  return {
    filled: filledPaths.map((path) => pathUnder(path.commands, m)),
    rimmed: filledPaths.map((path) => stroked.includes(path)),
    outlines: stroked.map((path) => pathUnder(path.commands, m)),
    width,
  };
}

/**
 * Paint `drawBody` (the geometry, its bevel, and the inner shadow, in slide space) through the
 * item's soft edge, as the SVG renderer's body filter: an edge fade (`drawEdgeFade`) of
 * everything the body paints, shrunk by 0.84 times the radius and blurred by the radius
 * (`softEdgeSpread`). Without a soft edge, `drawBody` paints as it is.
 */
export async function drawSoftEdge(
  item: Drawable,
  p: PageOps,
  scope: FillScope,
  groupScale: number,
  drawBody: () => Promise<void>
): Promise<void> {
  const soft = item.element.effects?.softEdge;
  const m = item.matrix;
  if (!soft || soft.radius <= 0 || !proEffects() || Math.abs(m[0] * m[3] - m[1] * m[2]) < 1e-9)
    return drawBody();
  const shape = softEdgeSilhouette(item, scope, groupScale);
  if (shape.filled.length === 0 && shape.outlines.length === 0) return drawBody();
  const { spread, sigma } = softEdgeSpread(soft.radius);
  // A shrink by more than half the narrower side of the bounds leaves nothing, as in the SVG
  // renderer, which compares the spread before its cap: the shape paints nothing.
  const corners = effectBoxCorners(item);
  const xs = corners.map((c) => c.x);
  const ys = corners.map((c) => c.y);
  const narrow = Math.min(Math.max(...xs) - Math.min(...xs), Math.max(...ys) - Math.min(...ys));
  if (Math.abs(spread) > narrow / 2) return;
  const inset = Math.min(Math.abs(spread), MAX_MORPH_SPREAD);
  await drawEdgeFade(shape, p, inset, sigma, corners, drawBody);
}

/**
 * Paint `drawBody` as one transparency group under a luminosity soft mask whose alpha follows
 * the painted edge of `shape` (the fill and the outline) shrunk by `inset` and blurred by
 * `sigma`. The mask holds the alpha by the distance to the painted edge, in bands: inside the
 * filled paths, strokes of the outer edge of their union (`outerEdges`) from the widest band in;
 * on the outline, outside the filled paths, strokes of the outline from the full width in. At a corner a band rounds where the blur
 * takes the product of the two edges. `corners` bound the shape in slide space.
 */
async function drawEdgeFade(
  shape: SoftEdgeSilhouette,
  p: PageOps,
  inset: number,
  sigma: number,
  corners: readonly { x: number; y: number }[],
  drawBody: () => Promise<void>
): Promise<void> {
  const body = p.form(await p.capture(drawBody), formBox(p), { S: 'Transparency' });
  drawThroughMask(p, body, await edgeFadeOps(shape, p, inset, sigma, corners));
}

/** Draw the transparency group `body` under a luminosity soft mask of the gray `maskOps`. */
export function drawThroughMask(p: PageOps, body: PDFRef, maskOps: PDFOperator[]): void {
  const mask = p.form(maskOps, formBox(p), { S: 'Transparency', CS: 'DeviceGray' });
  p.push(pushGraphicsState());
  p.state(`smask:${mask.toString()}`, () =>
    p.res.doc.context.obj({ Type: 'ExtGState', SMask: { Type: 'Mask', S: 'Luminosity', G: mask } })
  );
  p.push(drawObject(p.name('XObject', body)), popGraphicsState());
}

/** The gray operators of the edge fade of `shape` (`drawEdgeFade`), white where it is opaque. */
export async function edgeFadeOps(
  shape: SoftEdgeSilhouette,
  p: PageOps,
  inset: number,
  sigma: number,
  corners: readonly { x: number; y: number }[]
): Promise<PDFOperator[]> {
  const xs = corners.map((c) => c.x);
  const ys = corners.map((c) => c.y);
  const narrow = Math.min(Math.max(...xs) - Math.min(...xs), Math.max(...ys) - Math.min(...ys));
  const { reach, alpha } = softEdgeBands(inset, sigma);
  const extent = reach[reach.length - 1] ?? 0;
  const trace = (paths: readonly PathCommand[][]): void => {
    for (const path of paths) tracePath(path, p);
  };
  const W = shape.width;
  // The bands fade only the outer edge of the union of the filled paths: an edge inside another
  // filled path stays white (`outerEdges`).
  const edges = outerEdges(shape.filled);
  const ops = await p.capture(async () => {
    p.push(setLineJoin(LineJoinStyle.Round), setLineCap(LineCapStyle.Round));
    // Each filled path fills white on its own, so where paths overlap the mask stays white.
    p.push(setFillingGrayscaleColor(1));
    for (const path of shape.filled) {
      tracePath(path, p);
      p.push(fillEvenOdd());
    }
    // Inside a filled path: its own edge lies `off` inside the painted edge. Widest band first,
    // over the paths with and without an outline over their edge, so a point keeps the gray of
    // the narrowest band that reaches it.
    for (let k = reach.length - 1; k >= 0; k -= 1) {
      for (const rim of [false, true]) {
        const off = rim ? W / 2 : 0;
        if (reach[k] <= off) continue;
        const paths = edges.filter((path, i) => shape.rimmed[i] === rim && path.length > 0);
        if (paths.length === 0) continue;
        p.push(setStrokingGrayscaleColor(alpha[k]), setLineWidth(2 * (reach[k] - off)));
        trace(paths);
        p.push(stroke());
      }
    }
    // On the outline, outside the filled paths: the painted edge is the outline's outer edge.
    // From the full width in, each stroke covers the points at least its band inside that edge.
    if (shape.outlines.length && W > 0) {
      p.push(pushGraphicsState());
      if (shape.filled.length) {
        const r =
          4 * Math.max(narrow, W, extent) + Math.max(...xs.map(Math.abs), ...ys.map(Math.abs));
        clipOutside(shape.filled, r, p);
      }
      for (let j = 0; j < reach.length; j += 1) {
        const from = j ? reach[j - 1] : 0;
        if (from >= W / 2) break;
        p.push(setStrokingGrayscaleColor(alpha[j]), setLineWidth(W - 2 * from));
        trace([...shape.outlines]);
        p.push(stroke());
      }
      if (extent < W / 2) {
        p.push(setStrokingGrayscaleColor(1), setLineWidth(W - 2 * extent));
        trace([...shape.outlines]);
        p.push(stroke());
      }
      p.push(popGraphicsState());
    }
  });
  return ops;
}

/**
 * Clip to the outside of every path of `filled`: one clip per path, each a box of half-side `r`
 * minus the path by the even-odd rule. The clips intersect, so overlapping paths clip as their
 * union.
 */
function clipOutside(filled: readonly PathCommand[][], r: number, p: PageOps): void {
  for (const path of filled) {
    tracePath(
      [
        { op: 'M', x: -r, y: -r },
        { op: 'L', x: r, y: -r },
        { op: 'L', x: r, y: r },
        { op: 'L', x: -r, y: r },
        { op: 'Z' },
      ],
      p
    );
    tracePath(path, p);
    p.push(clipEvenOdd(), endPath());
  }
}

/** The color of a painted edge: the representative color of a fill (`threeDColor`). */
function edgeColor(fill: ResolvedFill, scope: FillScope): Rgba {
  return threeDColor(ownFill(fill, scope).fill);
}

/**
 * Paint `drawBody` through the item's blur (`a:blur`, ECMA-376 Part 1, 20.1.8.15), as the SVG
 * renderer's blur filter with the kernel of the shadows (`blurSigma`): inside the painted edge,
 * the body fades to half its alpha at the edge (an edge fade with no inset). With `grow`, a halo
 * in the color of the painted edge (the outline, else the fill) lies outside the filled paths,
 * as layered rings of the blurred profile. Text and detail inside the shape stay sharp.
 */
export async function drawBlur(
  item: Drawable,
  p: PageOps,
  scope: FillScope,
  groupScale: number,
  drawBody: () => Promise<void>
): Promise<void> {
  const blur = item.element.effects?.blur;
  const m = item.matrix;
  const sigma = blur ? blurSigma(blur.radius) : 0;
  if (!blur || sigma < MIN_SIGMA || !proEffects() || Math.abs(m[0] * m[3] - m[1] * m[2]) < 1e-9)
    return drawBody();
  const shape = softEdgeSilhouette(item, scope, groupScale);
  if (shape.filled.length === 0 && shape.outlines.length === 0) return drawBody();
  const corners = effectBoxCorners(item);
  if (blur.grow) {
    const el = item.element;
    const line = item.paintLine ?? el.line;
    const color = shape.width > 0 && line ? edgeColor(line.fill, scope) : edgeColor(el.fill, scope);
    p.push(pushGraphicsState());
    // Outside the filled paths only: inside, the faded body gives the profile.
    if (shape.filled.length) {
      const xs = corners.map((c) => Math.abs(c.x));
      const ys = corners.map((c) => Math.abs(c.y));
      const r = 4 * (Math.max(...xs, ...ys) + 3 * sigma + shape.width);
      clipOutside(shape.filled, r, p);
    }
    const identity: Matrix = [1, 0, 0, 1, 0, 0];
    for (const cls of await silhouette(item, identity, p, scope))
      await drawLayers(cls.parts, color, blurLayers(color.a, sigma, cls.bandWidth, 0), p);
    p.push(popGraphicsState());
  }
  await drawEdgeFade(shape, p, 0, sigma, corners, drawBody);
}
