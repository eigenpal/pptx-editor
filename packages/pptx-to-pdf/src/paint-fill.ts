/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Fills as PDF paint, from the same resolved fills the editor's SVG paint reads
// (`packages/core/src/output/paint-fill.ts`): solid colors, linear and path gradients as PDF
// shadings (a `path="shape"` gradient as the SVG paint's outline bands), and pictures cropped by
// `srcRect` and placed by `fillRect`. A fill other than a solid color clips to its path and
// paints the shape's local box, as an SVG paint server does.

import {
  closePath,
  concatTransformationMatrix,
  drawObject,
  lineTo,
  moveTo,
  popGraphicsState,
  pushGraphicsState,
  setFillingGrayscaleColor,
  setFillingRgbColor,
  setDashPattern,
  setLineCap,
  setLineJoin,
  setLineWidth,
  setStrokingRgbColor,
  LineCapStyle,
  LineJoinStyle,
  PDFContentStream,
  PDFName,
  PDFNumber,
  PDFOperator,
  PDFOperatorNames,
  type PDFRef,
} from 'pdf-lib';
import type { GradientStop, ResolvedFill, ResolvedLine, Rgba } from '@pptx-editor.dev/core/cascade';
import {
  dashArray,
  outsideDashArray,
  strokeCap,
  strokeWidth,
  type GeometryPath,
} from '@pptx-editor.dev/core/layout';
import {
  adjustForPathMode,
  circleGradientGeometry,
  gradientPaintStops,
  linearGradientEnds,
  rectGradientGeometry,
  shapeGradientBands,
  patternBitmap,
  PATTERN_PIXEL,
} from '@pptx-editor.dev/core/output';
import {
  around,
  clipEvenOdd,
  EMU_PER_PT,
  endPath,
  fillEvenOdd,
  rgb,
  shadingFill,
  traceRect,
  type PageOps,
  type StopFunctions,
} from './page-ops.ts';
import { drawPicture } from './paint-picture.ts';

/** Where a fill's pictures resolve, and the enclosing group's fill for `a:grpFill`. */
export interface FillScope {
  /** The part whose relationships resolve this element's `r:embed` ids. */
  readonly sourcePart: string;
  readonly groupFill?: ResolvedFill;
  readonly groupSourcePart?: string;
}

type PathGradient = Extract<ResolvedFill, { type: 'path' }>;

/** A fill with `a:grpFill` replaced by the enclosing group's fill, as the SVG paint does. */
export function ownFill(
  fill: ResolvedFill,
  scope: FillScope
): { fill: ResolvedFill; sourcePart: string } {
  if (fill.type !== 'group') return { fill, sourcePart: scope.sourcePart };
  const group = scope.groupFill;
  return group && group.type !== 'group'
    ? { fill: group, sourcePart: scope.groupSourcePart ?? scope.sourcePart }
    : { fill: { type: 'none' }, sourcePart: scope.sourcePart };
}

/**
 * Fill the path that `trace` draws, over a `w` × `h` local box. A solid color fills the path;
 * any other fill clips to it (even-odd, as the SVG `fill-rule`) and paints the box. `mode` is
 * the preset path fill mode (`a:path/@fill`). With `banded`, a `path="shape"` gradient draws as
 * bands of the path that `trace` draws (`shapeBandBudget` decides, as in the SVG paint); without
 * it, as a rectangle gradient. Returns false when the fill paints nothing.
 */
export async function paintFill(
  fill: ResolvedFill,
  w: number,
  h: number,
  trace: () => void,
  p: PageOps,
  scope: FillScope,
  what: string,
  mode: GeometryPath['fill'] = 'norm',
  banded = false
): Promise<boolean> {
  const own = ownFill(fill, scope);
  const f = own.fill;
  if (f.type === 'none' || f.type === 'group') return false;
  if (f.type === 'pattern') {
    // The preset path fill modes lighten and darken both colors, as in the SVG paint.
    const colors = { fg: adjustForPathMode(f.fg, mode), bg: adjustForPathMode(f.bg, mode) };
    p.push(pushGraphicsState(), ...patternColor({ ...f, ...colors }, p));
    trace();
    p.push(fillEvenOdd(), popGraphicsState());
    return true;
  }
  p.push(pushGraphicsState());
  if (f.type === 'solid') {
    const c = adjustForPathMode(f.color, mode);
    p.alpha(c.a, 1);
    p.push(setFillingRgbColor(...rgb(c)));
    trace();
    p.push(fillEvenOdd());
  } else {
    trace();
    p.push(clipEvenOdd(), endPath());
    if (f.type === 'picture') await drawPicture(own.sourcePart, f, w, h, p);
    else gradientBox(f, w, h, p, mode, banded ? trace : undefined);
  }
  p.push(popGraphicsState());
  return true;
}

/**
 * Set the fill color to a preset pattern, as the SVG paint draws it (`core/output`,
 * `preset-patterns.ts`): an 8 × 8 tile of 1 pt pixels in the background and foreground colors,
 * a half-tone pixel as the foreground at half opacity. A pattern maps to the space of the content
 * that uses it (ISO 32000-1, 8.7.3.1): the page's default space in points, or the slide EMU of a
 * form XObject (a reflection copy). Each space takes its own matrix to the slide grid, so the tiles
 * sit on that grid and stay square to the slide whatever the shape's transform.
 */
function patternColor(f: Extract<ResolvedFill, { type: 'pattern' }>, p: PageOps): PDFOperator[] {
  const height = p.page.getHeight();
  // Pattern space: one unit per pattern pixel, y down from the slide's top-left corner.
  const s = PATTERN_PIXEL / EMU_PER_PT;
  const matrix = p.inForm ? [PATTERN_PIXEL, 0, 0, PATTERN_PIXEL, 0, 0] : [s, 0, 0, -s, 0, height];
  const key = `${f.preset}|${f.fg.r},${f.fg.g},${f.fg.b},${f.fg.a}|${f.bg.r},${f.bg.g},${f.bg.b},${f.bg.a}|${matrix.join(',')}`;
  let name = p.patterns.get(key);
  if (!name) {
    name = p.name('Pattern', patternRef(f, matrix, key, p));
    p.patterns.set(key, name);
  }
  return [
    PDFOperator.of(PDFOperatorNames.NonStrokingColorspace, [PDFName.of('Pattern')]),
    PDFOperator.of(PDFOperatorNames.NonStrokingColorN, [name]),
  ];
}

/**
 * The tiling pattern of a preset in two colors with the pattern `matrix`, by `key`, written once
 * per document: a pattern of a master or a layout repeats on every slide.
 */
function patternRef(
  f: Extract<ResolvedFill, { type: 'pattern' }>,
  matrix: number[],
  key: string,
  p: PageOps
): PDFRef {
  let ref = p.res.patterns.get(key);
  if (!ref) {
    const ctx = p.res.doc.context;
    const bits = patternBitmap(f.preset) ?? patternBitmap('pct50')!;
    const states: Record<string, PDFRef> = {};
    const ops: PDFOperator[] = [];
    const paint = (c: Rgba, alpha: number): void => {
      const a = Math.max(0, Math.min(1, alpha));
      if (a < 1) {
        const id = `A${Object.keys(states).length}`;
        states[id] = ctx.register(ctx.obj({ Type: 'ExtGState', ca: a }));
        ops.push(PDFOperator.of(PDFOperatorNames.SetGraphicsStateParams, [PDFName.of(id)]));
      }
      ops.push(setFillingRgbColor(...rgb(c)));
    };
    const rect = (x: number, y: number, w: number, h: number): PDFOperator[] => [
      moveTo(x, y),
      lineTo(x + w, y),
      lineTo(x + w, y + h),
      lineTo(x, y + h),
      closePath(),
    ];
    ops.push(pushGraphicsState());
    paint(f.bg, f.bg.a);
    ops.push(...rect(0, 0, 8, 8), fillEvenOdd(), popGraphicsState());
    for (const coverage of [1, 0.5]) {
      const runs: PDFOperator[] = [];
      bits.forEach((row, y) => {
        let x = 0;
        while (x < 8) {
          if (row[x] !== coverage) {
            x += 1;
            continue;
          }
          let end = x;
          while (end < 8 && row[end] === coverage) end += 1;
          runs.push(...rect(x, y, end - x, 1));
          x = end;
        }
      });
      if (!runs.length) continue;
      ops.push(pushGraphicsState());
      paint(f.fg, f.fg.a * coverage);
      ops.push(...runs, fillEvenOdd(), popGraphicsState());
    }
    const dict = ctx.obj({
      Type: 'Pattern',
      PatternType: 1,
      PaintType: 1,
      TilingType: 1,
      BBox: [0, 0, 8, 8],
      XStep: 8,
      YStep: 8,
      Matrix: matrix,
      Resources: ctx.obj({ ExtGState: ctx.obj(states) }),
    });
    let bytes = dict.sizeInBytes();
    for (const op of ops) bytes += op.sizeInBytes() + 1;
    p.res.work.count(ops.length);
    p.res.work.reserveContent(bytes);
    ref = ctx.register(PDFContentStream.of(dict, ops, false));
    p.res.patterns.set(key, ref);
  }
  return ref;
}

/**
 * Paint a gradient over a `w` × `h` box at the origin, inside the current clip. `outline` traces
 * the filled path, for the bands of a `path="shape"` gradient; without it, that gradient paints
 * as a rectangle gradient, as the SVG paint does without a fill path.
 */
export function gradientBox(
  f: Extract<ResolvedFill, { type: 'linear' | 'path' }>,
  w: number,
  h: number,
  p: PageOps,
  mode: GeometryPath['fill'] = 'norm',
  outline?: () => void
): void {
  const adjusted = f.stops.map((s) => ({ ...s, color: adjustForPathMode(s.color, mode) }));
  if (f.type === 'path' && f.path === 'shape' && outline) {
    shapeGradient(f, adjusted, w, h, p, outline);
    return;
  }
  // The SVG paint's stop colors, so a two-stop gradient bends.
  const stops = gradientPaintStops(adjusted);
  if (f.type === 'linear') linearGradient(f.angle, f.scaled, stops, w, h, p);
  else if (f.path === 'circle') radialGradient(f, stops, w, h, p);
  else rectGradient(f, stops, w, h, p);
}

/** The color a fill shows when only one color can be drawn, such as text or a line. */
export function flatColor(
  fill: ResolvedFill,
  scope: FillScope,
  p: PageOps,
  what: string
): Rgba | null {
  const f = ownFill(fill, scope).fill;
  if (f.type === 'solid') return f.color;
  if ((f.type === 'linear' || f.type === 'path') && f.stops.length > 0) {
    // Only lines come here with a gradient: gradient text and shapes are drawn as shadings.
    p.diag.add('line.gradient', 'approximation', 'Gradient lines are drawn in their middle color.');
    return middleColor(f.stops);
  }
  if (f.type === 'picture' || f.type === 'pattern')
    p.diag.add(
      `fill.${f.type}`,
      'unsupported',
      `${f.type === 'picture' ? 'Picture' : 'Pattern'} fills on ${what} are not drawn yet.`
    );
  return null;
}

/** The gradient's color halfway along it. */
function middleColor(stops: readonly GradientStop[]): Rgba {
  const s = normalizedStops(stops);
  for (let i = 0; i + 1 < s.length; i += 1) {
    const [a, b] = [s[i], s[i + 1]];
    if (b.pos >= 0.5 && b.pos > a.pos) {
      const t = (0.5 - a.pos) / (b.pos - a.pos);
      const mix = (x: number, y: number): number => x + (y - x) * Math.max(0, Math.min(1, t));
      return {
        r: mix(a.color.r, b.color.r),
        g: mix(a.color.g, b.color.g),
        b: mix(a.color.b, b.color.b),
        a: mix(a.color.a, b.color.a),
      };
    }
  }
  return s[s.length - 1].color;
}

/**
 * Stops as the SVG gradient reads them: offsets clamped to 0–1 and never below the previous
 * one, padded with the first and last colors so the stops span 0 to 1.
 */
function normalizedStops(stops: readonly GradientStop[]): GradientStop[] {
  const out: GradientStop[] = [];
  let last = 0;
  for (const s of stops) {
    last = Math.max(last, Math.max(0, Math.min(1, s.pos)));
    out.push({ pos: last, color: s.color });
  }
  if (out[0].pos > 0) out.unshift({ pos: 0, color: out[0].color });
  if (out[out.length - 1].pos < 1) out.push({ pos: 1, color: out[out.length - 1].color });
  return out;
}

/**
 * At most this many stops are drawn. A common gradient editor allows 10. A stop costs about
 * 2 KB of memory in the PDF objects until the save, so a file cannot raise it. The cascade already
 * keeps at most 64 stops of a file (`cascade/paint-props.ts`), so this is a second bound.
 */
export const MAX_GRADIENT_STOPS = 64;

/** The stops, at most `MAX_GRADIENT_STOPS` of them, evenly chosen, first and last kept. */
function boundedStops(stops: readonly GradientStop[], p: PageOps): readonly GradientStop[] {
  if (stops.length <= MAX_GRADIENT_STOPS) return stops;
  p.diag.add(
    'fill.gradient-stops',
    'approximation',
    `Gradients with more than ${MAX_GRADIENT_STOPS} stops are drawn with ${MAX_GRADIENT_STOPS}.`
  );
  const last = stops.length - 1;
  return Array.from(
    { length: MAX_GRADIENT_STOPS },
    (_, i) => stops[Math.round((i * last) / (MAX_GRADIENT_STOPS - 1))]!
  );
}

/**
 * A PDF function over 0–1 through the stops: one exponential (type 2) function per segment,
 * stitched (type 3). Segments of zero length are hard steps, so they are left out.
 */
function stopFunction(stops: readonly GradientStop[], values: (c: Rgba) => number[]) {
  const segments: [GradientStop, GradientStop][] = [];
  for (let i = 0; i + 1 < stops.length; i += 1)
    if (stops[i + 1].pos - stops[i].pos > 1e-6) segments.push([stops[i], stops[i + 1]]);
  const part = ([a, b]: [GradientStop, GradientStop]) => ({
    FunctionType: 2,
    Domain: [0, 1],
    C0: values(a.color),
    C1: values(b.color),
    N: 1,
  });
  if (segments.length === 1) return part(segments[0]);
  return {
    FunctionType: 3,
    Domain: [0, 1],
    Functions: segments.map(part),
    Bounds: segments.slice(1).map(([a]) => a.pos),
    Encode: segments.flatMap(() => [0, 1]),
  };
}

/**
 * The stop functions of a fill, written the first time the document uses these stops; `null`
 * when it has no stops.
 */
function stopFunctions(stops: readonly GradientStop[], p: PageOps): StopFunctions | null {
  if (stops.length === 0) return null;
  const s = normalizedStops(boundedStops(stops, p));
  const key = s.map(({ pos, color: c }) => `${pos} ${c.r} ${c.g} ${c.b} ${c.a}`).join(';');
  const known = p.res.gradients.get(key);
  if (known) return known;
  const ctx = p.res.doc.context;
  const write = (values: (c: Rgba) => number[]): PDFRef => {
    const fn = ctx.obj(stopFunction(s, values));
    // The file decides how many segments there are: each is work, and its bytes are content.
    p.res.work.count(s.length);
    p.res.work.reserveContent(fn.sizeInBytes());
    return ctx.register(fn);
  };
  const alphas = s.map((x) => x.color.a);
  const flat = alphas.every((a) => a === alphas[0]);
  const fns = {
    color: write(rgb),
    alpha: flat ? null : write((c) => [c.a]),
    flatAlpha: alphas[0]!,
  };
  p.res.gradients.set(key, fns);
  return fns;
}

/** The fill operator of a type 2 (axial) or 3 (radial) shading with `coords` in local space. */
function shadingOp(
  type: 2 | 3,
  coords: number[],
  space: 'DeviceRGB' | 'DeviceGray',
  fn: PDFRef,
  p: PageOps
): PDFOperator {
  const ctx = p.res.doc.context;
  const dict = ctx.obj({
    ShadingType: type,
    ColorSpace: space,
    Coords: coords,
    Function: fn,
    Extend: [true, true],
  });
  p.res.work.reserveContent(dict.sizeInBytes());
  return shadingFill(p.name('Shading', ctx.register(dict)));
}

/**
 * Paint a shading over the current clip: type 2 (axial) or 3 (radial) with `coords` in local
 * space. Stops with different alphas paint through a luminosity soft mask of the same shading.
 * `bbox` bounds the mask in local space.
 */
function shade(
  type: 2 | 3,
  coords: number[],
  fns: StopFunctions,
  bbox: [number, number, number, number],
  p: PageOps
): void {
  withStopAlpha(
    fns.alpha ? null : fns.flatAlpha,
    bbox,
    p,
    (gray) =>
      p.push(
        gray
          ? shadingOp(type, coords, 'DeviceGray', fns.alpha!, p)
          : shadingOp(type, coords, 'DeviceRGB', fns.color, p)
      ),
    false
  );
}

/**
 * Paint `draw` with the stop alpha of a gradient, as the SVG paint's luminance masks: one
 * `flatAlpha` for every stop, or, when it is null, a luminosity soft mask that `draw(true)` paints
 * in gray (the stop alpha as luminance). `draw(false)` paints the colors opaque. The alpha of a
 * graphics state applies to each paint operator, so a construction of overlapping parts
 * (`grouped`) paints into a transparency group first, and the alpha applies once, to the group.
 * `bbox` bounds the mask and the group in local space.
 */
function withStopAlpha(
  flatAlpha: number | null,
  bbox: [number, number, number, number],
  p: PageOps,
  draw: (gray: boolean) => void,
  grouped = true
): void {
  if (flatAlpha === null) {
    const ops = p.captureNow(() => draw(true));
    const form = p.form(ops, bbox, { S: 'Transparency', CS: 'DeviceGray' });
    p.state(`smask:${form.toString()}`, () =>
      p.res.doc.context.obj({
        Type: 'ExtGState',
        SMask: { Type: 'Mask', S: 'Luminosity', G: form },
      })
    );
  } else p.alpha(flatAlpha, 1);
  if (!grouped || flatAlpha === 1) {
    draw(false);
    return;
  }
  const group = p.form(
    p.captureNow(() => draw(false)),
    bbox,
    { S: 'Transparency' }
  );
  p.push(drawObject(p.name('XObject', group)));
}

/** Set the fill color of a construction part: the color, or its alpha as gray for a mask. */
const partColor = (c: Rgba, gray: boolean): PDFOperator =>
  gray ? setFillingGrayscaleColor(Math.max(0, Math.min(1, c.a))) : setFillingRgbColor(...rgb(c));

/** `a:lin`: the SVG paint's gradient line (`linearGradientEnds`, ECMA-376 §20.1.8.41). */
function linearGradient(
  angleDeg: number,
  scaled: boolean,
  stops: readonly GradientStop[],
  w: number,
  h: number,
  p: PageOps
): void {
  const { x1, y1, x2, y2 } = linearGradientEnds(angleDeg, scaled, w, h);
  const fns = stopFunctions(stops, p);
  if (fns) shade(2, [x1, y1, x2, y2], fns, around(w, h), p);
}

/**
 * `a:path path="circle"`: the SVG paint's focal radial gradient (`circleGradientGeometry`),
 * from the focal point (the first stop) to the circle (the last stop).
 */
function radialGradient(
  fill: PathGradient,
  stops: readonly GradientStop[],
  w: number,
  h: number,
  p: PageOps
): void {
  const { cx, cy, r, fx, fy } = circleGradientGeometry(fill, w, h);
  const fns = stopFunctions(stops, p);
  if (fns) shade(3, [fx, fy, 0, cx, cy, r], fns, around(w, h), p);
}

/**
 * `a:path path="rect"`, and `shape` without bands: the SVG paint's construction
 * (`rectGradientGeometry`), concentric rectangles from `fillToRect` (the first stop) to the edge
 * of the tile rectangle (the last stop), as four trapezoids with linear gradients over a base of
 * the last color. Stop alpha masks the whole construction.
 */
function rectGradient(
  fill: PathGradient,
  stops: readonly GradientStop[],
  w: number,
  h: number,
  p: PageOps
): void {
  const fns = stopFunctions(stops, p);
  if (!fns) return;
  const { l, t, r, b, x0, y0, x1, y1 } = rectGradientGeometry(fill, w, h);
  withStopAlpha(fns.alpha ? null : fns.flatAlpha, around(w, h), p, (gray) => {
    const solid = (c: Rgba, x: number, y: number, rw: number, rh: number): void => {
      p.push(partColor(c, gray));
      traceRect(x, y, rw, rh, p);
      p.push(fillEvenOdd());
    };
    // The base hides antialiasing seams between the trapezoids.
    solid(stops[stops.length - 1].color, 0, 0, w, h);
    const trapezoid = (pts: [number, number][], coords: number[]): void => {
      p.push(pushGraphicsState(), moveTo(pts[0][0], pts[0][1]));
      for (const [x, y] of pts.slice(1)) p.push(lineTo(x, y));
      p.push(closePath(), clipEvenOdd(), endPath());
      p.push(
        gray
          ? shadingOp(2, coords, 'DeviceGray', fns.alpha!, p)
          : shadingOp(2, coords, 'DeviceRGB', fns.color, p)
      );
      p.push(popGraphicsState());
    };
    trapezoid(
      [
        [x0, y0],
        [x1, y0],
        [r, t],
        [l, t],
      ],
      [0, t, 0, y0]
    );
    trapezoid(
      [
        [l, b],
        [r, b],
        [x1, y1],
        [x0, y1],
      ],
      [0, b, 0, y1]
    );
    trapezoid(
      [
        [x0, y0],
        [l, t],
        [l, b],
        [x0, y1],
      ],
      [l, 0, x0, 0]
    );
    trapezoid(
      [
        [r, t],
        [x1, y0],
        [x1, y1],
        [r, b],
      ],
      [r, 0, x1, 0]
    );
    if (r > l && b > t) solid(stops[0].color, l, t, r - l, b - t);
  });
}

/**
 * `a:path path="shape"`: the SVG paint's bands (`shapeGradientBands`), copies of the outline that
 * `outline` traces, scaled about the focus center from 1 down, each filled in one color, outside
 * in. `stops` are the fill's stops after the path fill mode. Stop alpha masks all the bands.
 */
function shapeGradient(
  fill: PathGradient,
  stops: readonly GradientStop[],
  w: number,
  h: number,
  p: PageOps,
  outline: () => void
): void {
  if (stops.length === 0) return;
  const { cx, cy, bands } = shapeGradientBands(fill, stops, w, h);
  const alphas = stops.map((s) => s.color.a);
  const flat = alphas.every((a) => a === alphas[0]);
  withStopAlpha(flat ? alphas[0]! : null, around(w, h), p, (gray) => {
    for (const { scale: k, color } of bands) {
      p.push(
        pushGraphicsState(),
        concatTransformationMatrix(k, 0, 0, k, cx - k * cx, cy - k * cy),
        partColor(color, gray)
      );
      outline();
      p.push(fillEvenOdd(), popGraphicsState());
    }
  });
}

/** The PDF line cap of an SVG `stroke-linecap` value. */
export function lineCapStyle(cap: 'butt' | 'round' | 'square'): LineCapStyle {
  return cap === 'round'
    ? LineCapStyle.Round
    : cap === 'square'
      ? LineCapStyle.Projecting
      : LineCapStyle.Butt;
}

/**
 * Set the stroke for an outline, as the SVG `strokeAttrs`: a zero width is a hairline of one
 * pixel at 96 dpi, and `scale` undoes the enclosing group's scale. The width, the dash pattern,
 * and the caps come from the engine's shared rules (`core/layout`, `line-strokes.ts`). A
 * picture's border (`outside`) takes the dashes of a border outside the picture. `bands` is true
 * when the caller draws the bands of a compound line. Returns the stroke color, or null when the
 * outline draws nothing.
 */
export function strokeStyle(
  line: ResolvedLine,
  p: PageOps,
  scope: FillScope,
  scale = 1,
  outside = false,
  bands = false
): Rgba | null {
  const color = flatColor(line.fill, scope, p, 'lines');
  if (!color) return null;
  if (line.compound !== 'sng' && !bands)
    p.diag.add('line.compound', 'approximation', 'Compound lines are drawn as one line.');
  const width = strokeWidth(line, scale);
  const dashes = outside ? outsideDashArray(line, width) : dashArray(line, width);
  p.push(
    setStrokingRgbColor(...rgb(color)),
    setLineWidth(width),
    setLineCap(lineCapStyle(dashes && outside ? 'butt' : strokeCap(line))),
    setLineJoin(
      line.join === 'miter'
        ? LineJoinStyle.Miter
        : line.join === 'bevel'
          ? LineJoinStyle.Bevel
          : LineJoinStyle.Round
    )
  );
  if (dashes) p.push(setDashPattern(dashes, 0));
  // The SVG paint's `stroke-miterlimit`; PDF's default of 10 would differ.
  if (line.join === 'miter')
    p.push(
      PDFOperator.of(PDFOperatorNames.SetLineMiterLimit, [
        PDFNumber.of(Math.max(1, line.miterLimit)),
      ])
    );
  return color;
}
