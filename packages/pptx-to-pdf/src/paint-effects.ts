/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pptx-to-pdf/LICENSE.md.
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
  setFillingRgbColor,
  setLineCap,
  setLineJoin,
  setLineWidth,
  setStrokingRgbColor,
  stroke,
  type PDFName,
} from 'pdf-lib';
import type { Rgba } from '@pptx-editor.dev/core/cascade';
import {
  apply,
  ARROW_MITER_LIMIT,
  dashArray,
  drawingRenderers,
  hasCharacters,
  invert,
  multiply,
  planStroke,
  type GeometryPath,
  type LayoutItem,
  type Matrix,
  type PathCommand,
} from '@pptx-editor.dev/core/layout';
import { blurSigma, effectMatrix } from '@pptx-editor.dev/core/output';
import {
  glowSpread,
  innerShadowShift,
  reflectionBlurSigma,
  reflectionFade,
} from '@pptx-editor.dev/pro/drawing';
import { flatColor, ownFill, type FillScope } from './paint-fill.ts';
import { fillTextOutlines } from './paint-text.ts';
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

type Drawable = Extract<LayoutItem, { kind: 'shape' } | { kind: 'picture' }>;

/** A blur below half a point is drawn as a hard edge: it is under one output pixel. */
const MIN_SIGMA = 6_350;

/** A typical stem of a glyph, as a fraction of the font size: the width of a text silhouette. */
const STEM_PER_EM = 0.1;

/** Whether the Pro drawing renderers are installed, so glow, inner shadow, and reflection paint. */
function proEffects(): boolean {
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
    p.diag.add('effect.soft-edge', 'approximation', 'Soft edges are drawn hard.');
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
function formBox(p: PageOps): number[] {
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
interface Part {
  readonly grown: PDFName;
  readonly flat?: PDFName;
  readonly width: (grow: number) => number;
}

/** A class of silhouette parts with one blur profile: wide areas, or bands of one width. */
interface PartClass {
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
async function silhouette(
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
  const line = el.line && flatColor(el.line.fill, scope, p, 'lines') ? el.line : undefined;
  const lineWidth = line ? (line.width || 9525) * lineScale : 0;
  const trace = (paths: readonly GeometryPath[], close: PDFOperator) => () => {
    for (const path of paths) {
      tracePath(pathUnder(path.commands, full), p);
      p.push(close);
    }
  };
  const classes: PartClass[] = [];
  const filled = fills ? item.geometry.paths.filter((path) => path.fill !== 'none') : [];
  const outlined = filled.filter((path) => path.stroke && line);
  const bare = filled.filter((path) => !(path.stroke && line));
  const area: Part[] = [];
  if (outlined.length)
    area.push({
      grown: await form(p, trace(outlined, fillEvenOddAndStroke())),
      width: (grow) => lineWidth + 2 * grow,
    });
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
    // As the shape draws them (`planStroke`): open lines stop short under their line ends,
    // dashes keep their gaps, and the ends cast their own shadow. The dashes take round caps,
    // as every part of the silhouette does, so each dash keeps its drawn length.
    const plans = lines.map((path) =>
      item.kind === 'shape' ? planStroke(path.commands, line, matrixScale(item.matrix)) : undefined
    );
    const dashes = dashArray({ ...line, cap: 'rnd' }, lineWidth);
    const arrows = plans.flatMap((plan) => plan?.ends.filter((e) => e.kind === 'stroke') ?? []);
    const heads = plans.flatMap((plan) => plan?.ends.filter((e) => e.kind === 'fill') ?? []);
    const strokes = (): void => {
      if (dashes) p.push(setDashPattern([...dashes], 0));
      lines.forEach((path, i) => {
        tracePath(pathUnder(plans[i]?.commands ?? path.commands, full), p);
        p.push(stroke());
      });
      if (!arrows.length) return;
      p.push(
        setDashPattern([], 0),
        setLineJoin(LineJoinStyle.Miter),
        PDFOperator.of(PDFOperatorNames.SetLineMiterLimit, [PDFNumber.of(ARROW_MITER_LIMIT)])
      );
      for (const arrow of arrows) tracePath(pathUnder(arrow.commands, full), p);
      p.push(stroke());
    };
    const ends = (close: PDFOperator) => (): void => {
      for (const head of heads) {
        tracePath(pathUnder(head.commands, full), p);
        p.push(close);
      }
    };
    const parts: Part[] = [{ grown: await form(p, strokes), width: (g) => lineWidth + 2 * g }];
    if (heads.length)
      parts.push({
        grown: await form(p, ends(fillEvenOddAndStroke())),
        flat: await form(p, ends(fillEvenOdd())),
        width: (grow) => 2 * grow,
      });
    classes.push({ parts, bandWidth: lineWidth });
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

/** The outer shadow and the glow of a shape or picture, painted before its content. */
export async function drawEffectsBehind(
  item: Drawable,
  p: PageOps,
  scope: FillScope
): Promise<void> {
  const effects = item.element.effects;
  if (!effects) return;
  const { cx: w, cy: h } = item.element.xfrm;
  const shadow = effects.outerShadow;
  if (shadow && shadow.color.a > 0) {
    const m = effectMatrix(shadow, item.matrix, w, h);
    const sigma = blurSigma(shadow.blur);
    for (const cls of await silhouette(item, m, p, scope))
      await drawLayers(
        cls.parts,
        shadow.color,
        blurLayers(shadow.color.a, sigma, cls.bandWidth, shadow.dist),
        p
      );
  }
  const glow = effects.glow;
  if (glow && glow.radius > 0 && glow.color.a > 0 && proEffects()) {
    // As the SVG glow (`glowSpread`): the silhouette grown by `spread` and blurred by σ. The
    // inside level is the profile at the edge, a·Φ(spread / σ).
    const { spread, sigma } = glowSpread(glow.radius);
    const a = glow.color.a;
    const list = layers(
      rings(spread + 3 * sigma, sigma / 2),
      a * phi(spread / sigma),
      (d) => a * (1 - phi((d - spread) / sigma))
    );
    for (const cls of await silhouette(item, [1, 0, 0, 1, 0, 0], p, scope))
      await drawLayers(cls.parts, glow.color, list, p);
  }
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
 * The reflection of a shape or picture, painted before its shadow and content, as the SVG
 * renderer paints it: `drawContent` (the geometry, inner shadow, and text, in slide space) under
 * the reflection's effect transform, as one transparency group, through a luminosity soft mask
 * that fades from `startAlpha` at `startPos` to `endAlpha` at `endPos` along `fadeDir`. The blur
 * is left out (`effectDiagnostics` reports it).
 */
export async function drawReflection(
  item: Drawable,
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
