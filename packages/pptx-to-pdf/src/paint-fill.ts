/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Fills as PDF paint, from the same resolved fills the editor's SVG paint reads
// (`packages/core/src/output/paint-fill.ts`): solid colors, linear and path gradients as PDF
// shadings, and pictures cropped by `srcRect` and placed by `fillRect`. A fill other than a
// solid color clips to its path and paints the shape's local box, as an SVG paint server does.

import {
  closePath,
  concatTransformationMatrix,
  drawObject,
  lineTo,
  moveTo,
  popGraphicsState,
  pushGraphicsState,
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
  type PDFImage,
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
  metafilePng,
  patternBitmap,
  PATTERN_PIXEL,
} from '@pptx-editor.dev/core/output';
import { binaryPartOf, imageSize, metafileFormat, relTarget } from '@pptx-editor.dev/core/store';
import {
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
import { fontEmbeddingDecision } from './pdf-font-embedding.ts';

/**
 * The largest picture this writer decodes. `pdf-lib` decodes every PNG pixel, so a small file
 * that declares a huge size could exhaust memory; the header is read first and checked here.
 */
export const MAX_IMAGE_PIXELS = 25_000_000;
const MAX_IMAGE_SIDE = 20_000;

/** Where a fill's pictures resolve, and the enclosing group's fill for `a:grpFill`. */
export interface FillScope {
  /** The part whose relationships resolve this element's `r:embed` ids. */
  readonly sourcePart: string;
  readonly groupFill?: ResolvedFill;
  readonly groupSourcePart?: string;
}

type PictureFill = Extract<ResolvedFill, { type: 'picture' }>;
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
 * the preset path fill mode (`a:path/@fill`). Returns false when the fill paints nothing.
 */
export async function paintFill(
  fill: ResolvedFill,
  w: number,
  h: number,
  trace: () => void,
  p: PageOps,
  scope: FillScope,
  what: string,
  mode: GeometryPath['fill'] = 'norm'
): Promise<boolean> {
  const own = ownFill(fill, scope);
  const f = own.fill;
  if (f.type === 'none' || f.type === 'group') return false;
  if (f.type === 'pattern') {
    p.push(pushGraphicsState(), ...patternColor(f, p));
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
    else gradientBox(f, w, h, p, mode);
  }
  p.push(popGraphicsState());
  return true;
}

/**
 * Set the fill color to a preset pattern, as the SVG paint draws it (`core/output`,
 * `preset-patterns.ts`): an 8 × 8 tile of 1 pt pixels in the background and foreground colors,
 * a half-tone pixel as the foreground at half opacity. The pattern space is the page's, so the
 * tiles sit on the slide grid and stay square to the slide whatever the shape's transform. In a
 * form XObject (a shadow or a reflection copy), the tiles follow the form's space instead.
 */
function patternColor(f: Extract<ResolvedFill, { type: 'pattern' }>, p: PageOps): PDFOperator[] {
  const key = `${f.preset}|${f.fg.r},${f.fg.g},${f.fg.b},${f.fg.a}|${f.bg.r},${f.bg.g},${f.bg.b},${f.bg.a}`;
  let name = p.patterns.get(key);
  if (!name) {
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
    // Pattern space: one unit per pattern pixel, y down from the slide's top-left corner.
    const s = PATTERN_PIXEL / EMU_PER_PT;
    const dict = ctx.obj({
      Type: 'Pattern',
      PatternType: 1,
      PaintType: 1,
      TilingType: 1,
      BBox: [0, 0, 8, 8],
      XStep: 8,
      YStep: 8,
      Matrix: [s, 0, 0, -s, 0, p.page.getHeight()],
      Resources: ctx.obj({ ExtGState: ctx.obj(states) }),
    });
    let bytes = dict.sizeInBytes();
    for (const op of ops) bytes += op.sizeInBytes() + 1;
    p.res.work.count(ops.length);
    p.res.work.reserveContent(bytes);
    name = p.name('Pattern', ctx.register(PDFContentStream.of(dict, ops, false)));
    p.patterns.set(key, name);
  }
  return [
    PDFOperator.of(PDFOperatorNames.NonStrokingColorspace, [PDFName.of('Pattern')]),
    PDFOperator.of(PDFOperatorNames.NonStrokingColorN, [name]),
  ];
}

/** Paint a gradient over a `w` × `h` box at the origin, inside the current clip. */
export function gradientBox(
  f: Extract<ResolvedFill, { type: 'linear' | 'path' }>,
  w: number,
  h: number,
  p: PageOps,
  mode: GeometryPath['fill'] = 'norm'
): void {
  // The SVG paint's stop colors, so a two-stop gradient bends.
  const stops = gradientPaintStops(
    f.stops.map((s) => ({ ...s, color: adjustForPathMode(s.color, mode) }))
  );
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
  const ctx = p.res.doc.context;
  const shading = (space: 'DeviceRGB' | 'DeviceGray', fn: PDFRef): PDFName => {
    const dict = ctx.obj({
      ShadingType: type,
      ColorSpace: space,
      Coords: coords,
      Function: fn,
      Extend: [true, true],
    });
    p.res.work.reserveContent(dict.sizeInBytes());
    return p.name('Shading', ctx.register(dict));
  };
  if (fns.alpha) {
    const form = p.form([shadingFill(shading('DeviceGray', fns.alpha))], bbox, {
      S: 'Transparency',
      CS: 'DeviceGray',
    });
    p.state(`smask:${form.toString()}`, () =>
      ctx.obj({ Type: 'ExtGState', SMask: { Type: 'Mask', S: 'Luminosity', G: form } })
    );
  } else p.alpha(fns.flatAlpha, 1);
  p.push(shadingFill(shading('DeviceRGB', fns.color)));
}

/** The local box grown by its own size on every side, for soft-mask bounds. */
const around = (w: number, h: number): [number, number, number, number] => [
  -Math.abs(w),
  -Math.abs(h),
  2 * Math.abs(w),
  2 * Math.abs(h),
];

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
 * `a:path path="rect"` (and `shape`, which the SVG paint draws the same way): concentric
 * rectangles from `fillToRect` (the first stop) to the box edge (the last stop), as four
 * trapezoids with linear gradients over a base of the last color.
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
  const f = fill.fillToRect;
  const l = Math.min(f.l * w, w);
  const t = Math.min(f.t * h, h);
  const r = Math.max(l, w - f.r * w);
  const b = Math.max(t, h - f.b * h);
  const solid = (c: Rgba, x: number, y: number, rw: number, rh: number): void => {
    // The SVG paint draws the base and the center opaque.
    p.push(setFillingRgbColor(...rgb(c)));
    traceRect(x, y, rw, rh, p);
    p.push(fillEvenOdd());
  };
  solid(stops[stops.length - 1].color, 0, 0, w, h);
  const trapezoid = (pts: [number, number][], coords: number[]): void => {
    p.push(pushGraphicsState(), moveTo(pts[0][0], pts[0][1]));
    for (const [x, y] of pts.slice(1)) p.push(lineTo(x, y));
    p.push(closePath(), clipEvenOdd(), endPath());
    shade(2, coords, fns, around(w, h), p);
    p.push(popGraphicsState());
  };
  const [tl, tr, br, bl]: [number, number][] = [
    [0, 0],
    [w, 0],
    [w, h],
    [0, h],
  ];
  trapezoid([tl, tr, [r, t], [l, t]], [0, t, 0, 0]);
  trapezoid([[l, b], [r, b], br, bl], [0, b, 0, h]);
  trapezoid([tl, [l, t], [l, b], bl], [l, 0, 0, 0]);
  trapezoid([[r, t], tr, br, [r, b]], [r, 0, w, 0]);
  if (r > l && b > t) solid(stops[0].color, l, t, r - l, b - t);
}

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
const isPng = (b: Uint8Array): boolean => PNG_SIGNATURE.every((v, i) => b[i] === v);
const isJpeg = (b: Uint8Array): boolean => b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff;

/** Embed a picture part once per document: the format from its signature, its size checked first. */
async function imageFor(target: string, p: PageOps): Promise<PDFImage | null> {
  const cached = p.res.images.get(target);
  if (cached !== undefined) return cached;
  let image: PDFImage | null = null;
  let bytes = binaryPartOf(p.res.pkg, target);
  // An EMF or WMF picture draws as its PNG, converted as in the editor.
  const metafile = bytes !== undefined && metafileFormat(bytes) !== undefined;
  if (bytes && metafile) {
    // Metafile text draws only with fonts that their embedding permission lets the PDF carry.
    const fonts = (p.res.metafileFonts ??= [...p.res.faces.values()]
      .filter((f) => {
        const face = {
          identity: f.id,
          bytes: f.bytes,
          faceIndex: 0,
          request: { family: f.family },
        };
        return fontEmbeddingDecision(face).kind !== 'refuse';
      })
      .map((f) => f.bytes));
    let png = await metafilePng(bytes, fonts, p.res.work.signal);
    // When the export's signal or deadline stopped the conversion, the export fails.
    p.res.work.check();
    // A picture that the converter's own time limit stopped tries once more (a busy machine),
    // then reports `picture.metafile`.
    if (png === undefined) png = await metafilePng(bytes, fonts, p.res.work.signal);
    p.res.work.check();
    bytes = png ?? undefined;
  }
  const png = bytes ? isPng(bytes) : false;
  if (!bytes && metafile)
    p.diag.add('picture.metafile', 'unsupported', 'An EMF or WMF picture does not convert.');
  else if (!bytes) p.diag.add('picture.missing', 'unsupported', 'A picture part is missing.');
  else if (!png && !isJpeg(bytes))
    p.diag.add('picture.format', 'unsupported', 'A picture is not PNG, JPEG, EMF, or WMF.');
  else {
    const size = imageSize(bytes);
    if (!size || size.width <= 0 || size.height <= 0)
      p.diag.add('picture.unreadable', 'unsupported', 'A picture header cannot be read.');
    else if (
      size.width > MAX_IMAGE_SIDE ||
      size.height > MAX_IMAGE_SIDE ||
      size.width * size.height > MAX_IMAGE_PIXELS
    )
      p.diag.add(
        'picture.too-large',
        'unsupported',
        `A ${size.width} × ${size.height} picture exceeds the ${MAX_IMAGE_PIXELS} pixel limit.`
      );
    else {
      // The picture's file bytes become PDF content once per document. pdf-lib keeps a PNG
      // decoded until the save, and only the pixel limit above bounds each one. DOCX Editor
      // passes a PNG through without decoding it.
      p.res.work.reserveContent(bytes.length);
      try {
        image = png ? await p.res.doc.embedPng(bytes) : await p.res.doc.embedJpg(bytes);
      } catch {
        p.diag.add('picture.unreadable', 'unsupported', 'A picture cannot be decoded.');
      }
    }
  }
  p.res.images.set(target, image);
  return image;
}

/** Whether a picture fill has an image this writer draws (the image is embedded once). */
export async function canDrawPicture(
  sourcePart: string,
  fill: PictureFill,
  p: PageOps
): Promise<boolean> {
  const target = relTarget(p.res.pkg, sourcePart, fill.rId);
  return target ? (await imageFor(target, p)) !== null : false;
}

/**
 * Draw a picture fill over a `w` × `h` box, as the SVG `pictureImage`: `fillRect` places the
 * picture in the box, and `srcRect` crops it, so the picture's visible part fills that place.
 * The caller clips to the shape.
 */
export async function drawPicture(
  sourcePart: string,
  fill: PictureFill,
  w: number,
  h: number,
  p: PageOps
): Promise<void> {
  const target = relTarget(p.res.pkg, sourcePart, fill.rId);
  if (!target) {
    p.diag.add('picture.missing', 'unsupported', 'A picture part is missing or external.');
    return;
  }
  if (fill.mode === 'tile')
    p.diag.add('picture.tile', 'approximation', 'Tiled pictures are drawn stretched.');
  // The fallback is the deck's own raster picture of the SVG, so it is reported, not refused.
  if (fill.svgRId)
    p.diag.add(
      'picture.svg-fallback',
      'information',
      'SVG pictures are drawn from their PNG fallback.'
    );
  if (fill.effects.some((e) => e.kind !== 'alphaModFix'))
    p.diag.add(
      'picture.color-effects',
      'approximation',
      'Picture color effects (grayscale, duotone, brightness, black and white) are not drawn.'
    );
  const image = await imageFor(target, p);
  if (!image) return;
  const fr = fill.fillRect;
  const bx = fr.l * w;
  const by = fr.t * h;
  const bw = w * (1 - fr.l - fr.r);
  const bh = h * (1 - fr.t - fr.b);
  // A crop that leaves nothing visible on an axis (two concurrent crops merged edge by edge)
  // draws uncropped on that axis, as the SVG paint and the crop commands read it.
  const raw = fill.srcRect;
  const okW = 1 - raw.l - raw.r >= 0.001;
  const okH = 1 - raw.t - raw.b >= 0.001;
  const sr = {
    l: okW ? raw.l : 0,
    r: okW ? raw.r : 0,
    t: okH ? raw.t : 0,
    b: okH ? raw.b : 0,
  };
  const iw = bw / (1 - sr.l - sr.r);
  const ih = bh / (1 - sr.t - sr.b);
  const alpha = fill.effects.reduce((a, e) => (e.kind === 'alphaModFix' ? a * e.amount : a), 1);
  const name = p.name('XObject', image.ref);
  p.push(pushGraphicsState());
  p.alpha(alpha, 1);
  // The image space is the unit square with y up; the picture is iw × ih EMU with y down.
  p.push(concatTransformationMatrix(iw, 0, 0, -ih, bx - sr.l * iw, by - sr.t * ih + ih));
  p.push(drawObject(name), popGraphicsState());
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
