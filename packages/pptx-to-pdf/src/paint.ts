/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Paint one laid-out slide into a PDF page. Like DOCX Editor's `docx-to-pdf`
// (`packages/docx-to-pdf/src/paint.ts` at adf0e178), this reads the engine's layout, the same
// `layoutSlide` output the editor's SVG paint reads; it is not a second layout.
//
// The page content works in slide EMU with y down: one page matrix maps EMU to PDF points and
// flips y, so geometry paths and layout matrices apply unchanged. Text draws one character at a
// time at the x position layout measured, so line breaks and positions are the engine's.

import {
  closePath,
  concatTransformationMatrix,
  drawObject,
  fill,
  lineTo,
  moveTo,
  popGraphicsState,
  pushGraphicsState,
  setDashPattern,
  setFillingGrayscaleColor,
  setFillingRgbColor,
  setLineCap,
  setLineJoin,
  setLineWidth,
  setStrokingGrayscaleColor,
  stroke,
  LineJoinStyle,
  PDFNumber,
  PDFOperator,
  PDFOperatorNames,
} from 'pdf-lib';
import type { BlendMode, ResolvedFill, ResolvedLine, Rgba } from '@pptx-editor.dev/core/cascade';
import {
  ARROW_MITER_LIMIT,
  dashArray,
  hasCharacters,
  planStroke,
  strokeCap,
  strokeWidth,
  type LayoutItem,
  type Matrix,
  type PathCommand,
  type SlideLayout,
  type StrokePlan,
} from '@pptx-editor.dev/core/layout';
import {
  compoundPlan,
  MAX_OVERLAY_PATHS,
  threeDFaceItem,
  type CompoundBand,
  type CompoundBandLine,
  type CompoundPlan,
} from '@pptx-editor.dev/core/output';
import {
  drawGlow,
  drawShadows,
  drawInnerShadow,
  drawReflection,
  drawBlur,
  drawSoftEdge,
  effectDiagnostics,
} from './paint-effects.ts';
import { canDrawPicture, drawPicture } from './paint-picture.ts';
import { lineCapStyle, ownFill, paintFill, strokeStyle, type FillScope } from './paint-fill.ts';
import { drawTable } from './paint-table.ts';
import { faceMatrix, fillFacets, threeDOf, turnsFace } from './paint-three-d.ts';
import { drawText } from './paint-text.ts';
import {
  around,
  clip,
  clipEvenOdd,
  clipNonZero,
  matrixScale,
  EMU_PER_PT,
  endPath,
  PageOps,
  rgb,
  tracePath,
  traceRect,
  type PaintResources,
  type SlideDiagnostics,
} from './page-ops.ts';

/** One pixel at 96 DPI (9525 EMU, 0.75 pt): less than this past the box is rounding in the presets. */
const PAST_BOX_EMU = 9525;

/** The point at `s` (0–1) along a curve that starts at (x0, y0). */
function curvePoint(
  c: Extract<PathCommand, { op: 'C' | 'Q' }>,
  x0: number,
  y0: number,
  s: number
): [number, number] {
  const u = 1 - s;
  if (c.op === 'Q')
    return [
      u * u * x0 + 2 * u * s * c.x1 + s * s * c.x,
      u * u * y0 + 2 * u * s * c.y1 + s * s * c.y,
    ];
  return [
    u * u * u * x0 + 3 * u * u * s * c.x1 + 3 * u * s * s * c.x2 + s * s * s * c.x,
    u * u * u * y0 + 3 * u * u * s * c.y1 + 3 * u * s * s * c.y2 + s * s * s * c.y,
  ];
}

/** Whether a path reaches past the `w` × `h` box by more than one point, curves included. */
function pastBox(commands: readonly PathCommand[], w: number, h: number): boolean {
  const t = PAST_BOX_EMU;
  const out = (x: number, y: number): boolean => x < -t || y < -t || x > w + t || y > h + t;
  let [x0, y0] = [0, 0];
  for (const c of commands) {
    if (c.op === 'Z') continue;
    if (out(c.x, c.y)) return true;
    // A curve can bulge past the box between its ends (`heart`): test points along it.
    if (c.op === 'C' || c.op === 'Q')
      for (let i = 1; i < 8; i += 1) if (out(...curvePoint(c, x0, y0, i / 8))) return true;
    [x0, y0] = [c.x, c.y];
  }
  return false;
}

/**
 * The line ends and square dash caps of a stroke, as the SVG paint draws them (`output/paint.ts`,
 * `strokeExtras`): filled heads, open arrows stroked as wide as the line with round caps and a
 * mitered tip, and solid caps with butt ends. The stroke color and alpha are already set.
 */
function drawStrokeExtras(plan: StrokePlan, color: Rgba, width: number, p: PageOps): void {
  if (!plan.ends.length && !plan.caps.length) return;
  p.push(setDashPattern([], 0));
  for (const end of plan.ends) {
    if (end.kind === 'fill') {
      p.alpha(color.a, color.a);
      p.push(setFillingRgbColor(...rgb(color)));
      tracePath(end.commands, p);
      p.push(fill());
      continue;
    }
    p.push(
      setLineWidth(end.width),
      setLineCap(lineCapStyle('round')),
      setLineJoin(LineJoinStyle.Miter),
      PDFOperator.of(PDFOperatorNames.SetLineMiterLimit, [PDFNumber.of(ARROW_MITER_LIMIT)])
    );
    tracePath(end.commands, p);
    p.push(stroke());
  }
  if (!plan.caps.length) return;
  p.push(setLineWidth(width), setLineCap(lineCapStyle('butt')));
  for (const cap of plan.caps) tracePath(cap, p);
  p.push(stroke());
}

type Drawable = Extract<LayoutItem, { kind: 'shape' } | { kind: 'picture' }>;

/**
 * Clip to the outside of a picture's path and double the line width, so the stroke lies
 * outside the picture, as the
 * editor's paint does (`output/paint.ts`, `outsideOutline`).
 */
function outsideOf(
  commands: readonly PathCommand[],
  width: number,
  cx: number,
  cy: number,
  p: PageOps
): void {
  const pad = 2 * width + Math.max(cx, cy);
  p.push(
    moveTo(-pad, -pad),
    lineTo(cx + pad, -pad),
    lineTo(cx + pad, cy + pad),
    lineTo(-pad, cy + pad),
    closePath()
  );
  tracePath(commands, p);
  p.push(clipEvenOdd(), endPath(), setLineWidth(2 * width));
}

/** The PDF blend mode of each `ST_BlendMode`, as the SVG paint's `mix-blend-mode`. */
const PDF_BLEND: Readonly<Record<BlendMode, string>> = {
  over: 'Normal',
  mult: 'Multiply',
  screen: 'Screen',
  darken: 'Darken',
  lighten: 'Lighten',
};

/**
 * Paint a fill and its `a:fillOverlay`, as the SVG paint does: `base` paints the fill, then the
 * overlay fill blends over the path that `trace` draws. Both paint into an isolated transparency
 * group, so the blend sees only this fill, not the slide behind it. Returns false, with nothing
 * painted, when the item has no overlay to draw; the caller then paints `base` itself. A shape
 * with more than `MAX_OVERLAY_PATHS` paths draws no overlay, as in the SVG paint.
 */
async function withOverlay(
  item: Drawable,
  w: number,
  h: number,
  trace: () => void,
  p: PageOps,
  scope: FillScope,
  base: () => Promise<boolean>
): Promise<boolean> {
  const overlay = item.element.effects?.fillOverlay;
  if (!overlay) return false;
  if (item.geometry.paths.length > MAX_OVERLAY_PATHS) {
    p.diag.add(
      'effect.fill-overlay',
      'unsupported',
      `A fill overlay on a shape with more than ${MAX_OVERLAY_PATHS} paths is not drawn.`
    );
    return false;
  }
  const ops = await p.capture(async () => {
    if (!(await base())) return;
    p.push(pushGraphicsState());
    p.state(`blend:${overlay.blend}`, () =>
      p.res.doc.context.obj({ Type: 'ExtGState', BM: PDF_BLEND[overlay.blend] })
    );
    // A `path="shape"` overlay draws as bands of the outline, as the SVG paint draws it.
    const banded = bandsShapeGradient(item, overlay.fill, p);
    await paintFill(overlay.fill, w, h, trace, p, scope, 'fill overlays', 'norm', banded);
    p.push(popGraphicsState());
  });
  const group = p.form(ops, around(w, h), { S: 'Transparency', I: true });
  p.push(drawObject(p.name('XObject', group)));
  return true;
}

/** A shape's or picture's geometry: fills, the picture clipped to the geometry, outlines. */
async function drawGeometry(
  item: Drawable,
  p: PageOps,
  scope: FillScope,
  groupScale: number
): Promise<void> {
  const el = item.element;
  const { cx: w, cy: h } = el.xfrm;
  const paths = item.geometry.paths;
  if (item.kind === 'picture') {
    const own = ownFill(el.fill, scope);
    if (own.fill.type === 'picture') {
      // The picture shows inside its geometry (nonzero, as the SVG clip path).
      p.push(pushGraphicsState());
      if (paths.length) for (const path of paths) tracePath(path.commands, p);
      else traceRect(0, 0, w, h, p);
      p.push(clipNonZero(), endPath());
      const pictureFill = own.fill;
      const picture = async (): Promise<boolean> => {
        await drawPicture(own.sourcePart, pictureFill, w, h, p);
        return true;
      };
      const traceAll = (): void => {
        if (paths.length) for (const path of paths) tracePath(path.commands, p);
        else traceRect(0, 0, w, h, p);
      };
      if (!(await withOverlay(item, w, h, traceAll, p, scope, picture))) await picture();
      p.push(popGraphicsState());
    }
  }
  // Picture fills and rectangular path gradients paint the shape's box only.
  const boxed = ownFill(el.fill, scope).fill;
  // Only a shape's fill: the editor draws a `p:pic` the same way, its image clipped to the
  // geometry, so a picture's geometry past its box differs from nothing.
  const boxOnly =
    item.kind === 'shape' &&
    (boxed.type === 'picture' || (boxed.type === 'path' && boxed.path !== 'circle'));
  // Only an outline that draws: a picture with `a:ln/a:noFill` has none to split.
  const line = item.paintLine ?? el.line;
  if (item.kind === 'picture' && line && line.fill.type !== 'none' && paths.length > 1)
    p.diag.add(
      'picture.outline',
      'approximation',
      'The border of a picture cropped to a shape with several paths is drawn path by path.'
    );
  const banded = item.kind === 'shape' && bandsShapeGradient(item, ownFill(el.fill, scope).fill, p);
  // Path by path, fill then outline, in the SVG paint's order.
  for (const path of paths) {
    if (boxOnly && path.fill !== 'none' && pastBox(path.commands, w, h))
      p.diag.add(
        'fill.outside-box',
        'approximation',
        "Picture fills and rectangular path gradients paint only the shape's box."
      );
    if (item.kind === 'shape' && path.fill !== 'none') {
      const trace = (): void => tracePath(path.commands, p);
      const own = (): Promise<boolean> =>
        paintFill(el.fill, w, h, trace, p, scope, 'shapes', path.fill, banded);
      if (!(await withOverlay(item, w, h, trace, p, scope, own))) await own();
    }
    // A picture's outline follows all of its geometry, as in the SVG paint.
    if (!line || (!path.stroke && item.kind !== 'picture')) continue;
    p.push(pushGraphicsState());
    const picture = item.kind === 'picture';
    const width = strokeWidth(line, groupScale);
    // An open path stops short under its line ends, and a dashed one can get square caps, as
    // in the SVG paint (`core/layout`, `planStroke`).
    const plan = picture ? undefined : planStroke(path.commands, line, groupScale);
    const commands = plan ? plan.commands : path.commands;
    const bands = picture ? null : compoundOf(line, width, commands);
    const color = strokeStyle(line, p, scope, groupScale, picture, !!bands);
    if (color) {
      p.alpha(1, color.a);
      // A picture's border lies outside the picture.
      if (picture) outsideOf(path.commands, width, w, h, p);
      if (bands?.kind === 'offset')
        drawBands(bands.bands, line.join === 'miter' ? line.miterLimit : 4, p);
      else if (bands) drawMaskedBands(bands, line, width, commands, w, h, p);
      else {
        tracePath(commands, p);
        p.push(stroke());
      }
      if (plan) drawStrokeExtras(plan, color, width, p);
    }
    p.push(popGraphicsState());
  }
}

/**
 * Whether a shape draws its `path="shape"` gradient as bands of its paths, as the SVG paint
 * decides (`shapeBandBudget`): only shapes with such a fill spend the page's band work, once each.
 */
function bandsShapeGradient(item: Drawable, fill: ResolvedFill, p: PageOps): boolean {
  if (fill.type !== 'path' || fill.path !== 'shape') return false;
  let banded = p.bandedShapes.get(item);
  if (banded === undefined) {
    const paths = item.geometry.paths;
    const commands = paths.reduce((n, path) => n + path.commands.length, 0);
    banded = p.shapeBands.allow(paths.length, commands);
    p.bandedShapes.set(item, banded);
  }
  return banded;
}

/**
 * How a compound line draws, as in the SVG paint (`core/output`, `compoundPlan`): offset bands,
 * or bands cut from the line's own stroke by a mask. Null for a single line.
 */
function compoundOf(
  line: ResolvedLine,
  width: number,
  commands: readonly PathCommand[]
): CompoundPlan | null {
  if (line.compound === 'sng') return null;
  const join = line.join === 'miter' ? 'miter' : line.join === 'bevel' ? 'bevel' : 'round';
  const closed = commands[commands.length - 1]?.op === 'Z';
  const dashed = !!dashArray(line, width);
  return compoundPlan(line.compound, width, closed, dashed, join, miterLimitOf(line), commands);
}

/** The miter limit of a line as the SVG paint writes it: 4 when the join is not a miter. */
const miterLimitOf = (line: ResolvedLine): number =>
  line.join === 'miter' ? Math.max(1, line.miterLimit) : 4;

/**
 * Draw the masked bands of a compound line, as the SVG paint's `compoundStroke`: the line's own
 * stroke (color, width, caps, joins, and dashes already set) under a luminosity soft mask of a
 * white stroke `2b` wide and a black stroke `2a` wide for each band. On a closed path, the outer
 * bands lose the inside of the path, and the inner bands keep only the inside.
 */
function drawMaskedBands(
  plan: Extract<CompoundPlan, { kind: 'masked' }>,
  line: ResolvedLine,
  width: number,
  given: readonly PathCommand[],
  w: number,
  h: number,
  p: PageOps
): void {
  const commands = plan.commands ?? given;
  const grow = width + Math.max(w, h);
  const bbox = [-grow, -grow, w + 2 * grow, h + 2 * grow];
  const dashes = dashArray(line, width);
  const join =
    line.join === 'miter'
      ? LineJoinStyle.Miter
      : line.join === 'bevel'
        ? LineJoinStyle.Bevel
        : LineJoinStyle.Round;
  /** Paint the line through a mask of `bands`; `cutInside` removes the inside of the path. */
  const painted = (bands: readonly CompoundBand[], cutInside: boolean): void => {
    const ops = p.captureNow(() => {
      // A soft mask starts from the default graphics state: set the line's own stroke.
      p.push(
        setLineCap(lineCapStyle(strokeCap(line))),
        setLineJoin(join),
        PDFOperator.of(PDFOperatorNames.SetLineMiterLimit, [PDFNumber.of(miterLimitOf(line))])
      );
      if (dashes) p.push(setDashPattern(dashes, 0));
      // Widest first, so each black stroke cuts only its own band.
      for (const band of [...bands].sort((x, y) => y.b - x.b)) {
        p.push(setStrokingGrayscaleColor(1), setLineWidth(2 * band.b));
        tracePath(commands, p);
        p.push(stroke());
        if (band.a > 0) {
          p.push(setStrokingGrayscaleColor(0), setLineWidth(2 * band.a));
          tracePath(commands, p);
          p.push(stroke());
        }
      }
      if (cutInside) {
        p.push(setFillingGrayscaleColor(0));
        tracePath(commands, p);
        p.push(fill());
      }
    });
    const form = p.form(ops, bbox, { S: 'Transparency', CS: 'DeviceGray' });
    p.push(pushGraphicsState());
    p.state(`smask:${form.toString()}`, () =>
      p.res.doc.context.obj({
        Type: 'ExtGState',
        SMask: { Type: 'Mask', S: 'Luminosity', G: form },
      })
    );
    tracePath(commands, p);
    p.push(stroke(), popGraphicsState());
  };
  if (!plan.closed) {
    painted(plan.outer, false);
    return;
  }
  if (plan.outer.length) painted(plan.outer, true);
  if (plan.inner.length) {
    // The inside of the path, nonzero, as the SVG paint's mask of the filled path.
    p.push(pushGraphicsState());
    tracePath(commands, p);
    p.push(clipNonZero(), endPath());
    painted(plan.inner, false);
    p.push(popGraphicsState());
  }
}

/** Stroke the bands of a compound line in the current stroke color, caps, and dashes. */
function drawBands(bands: readonly CompoundBandLine[], miterLimit: number, p: PageOps): void {
  // The corners are in the offset geometry: a miter join keeps them as built.
  p.push(
    setLineJoin(LineJoinStyle.Miter),
    PDFOperator.of(PDFOperatorNames.SetLineMiterLimit, [PDFNumber.of(Math.max(miterLimit, 4))])
  );
  for (const band of bands) {
    p.push(setLineWidth(band.width));
    for (const { points, closed } of band.lines) {
      points.forEach((q, i) => p.push(i ? lineTo(q.x, q.y) : moveTo(q.x, q.y)));
      if (closed) p.push(closePath());
    }
    p.push(stroke());
  }
}

/** Why a frame is not drawn. A diagram names why it shows its placeholder. */
function frameMessage(
  kind: string,
  fallback: 'no-drawing' | 'unsupported' | 'limit' | undefined
): string {
  if (kind !== 'diagram' || !fallback) return `${kind} frames are not drawn yet.`;
  if (fallback === 'limit')
    return 'A diagram is not drawn: its saved drawing is past the shape, nesting, or slide limit.';
  if (fallback === 'unsupported')
    return 'A diagram is not drawn: its saved drawing holds content that is not drawn yet.';
  return 'A diagram is not drawn: it has no saved drawing.';
}

async function drawItem(
  item: LayoutItem,
  p: PageOps,
  scope: FillScope,
  groupScale: number
): Promise<void> {
  if (item.element.hidden) return;
  await p.step();
  const local: FillScope = { ...scope, sourcePart: item.element.partName };
  if (item.kind === 'group') {
    const inner: FillScope = {
      ...local,
      groupFill: item.element.fill,
      groupSourcePart: item.element.partName,
    };
    for (const child of item.children) await drawItem(child, p, inner, matrixScale(item.matrix));
    return;
  }
  if (item.kind === 'frame') {
    // An OLE object shows its picture, as in the SVG paint.
    if (item.poster) return drawItem(item.poster, p, local, groupScale);
    // A diagram draws its saved drawing, as in the SVG paint.
    if (item.drawing) {
      for (const child of item.drawing) await drawItem(child, p, scope, matrixScale(item.matrix));
      return;
    }
    // A chart drawn by an installed renderer, as in the SVG paint.
    if (item.chart) {
      for (const child of item.chart) await drawItem(child, p, local, groupScale);
      return;
    }
    if (!item.table) {
      const kind = item.element.frameKind;
      p.diag.add(`frame.${kind}`, 'unsupported', frameMessage(kind, item.element.drawingFallback));
      return;
    }
    p.push(pushGraphicsState());
    p.matrix(item.matrix);
    await drawTable(item.table, item.element.table?.background ?? { type: 'none' }, p, local);
    p.push(popGraphicsState());
    return;
  }
  // A picture fill whose image is not drawn (a `p:pic`, or a shape's picture fill) casts no
  // shadow and has no inside to shade.
  const own = ownFill(item.element.fill, local);
  const shown = own.fill.type !== 'picture' || (await canDrawPicture(own.sourcePart, own.fill, p));
  if (shown) effectDiagnostics(item, p, local);
  // Bevels, contour, extrusion, and the camera's turn of the face, as the SVG paint draws them.
  const threeD = threeDOf(item);
  const turned = turnsFace(threeD);
  const drawContent = async (): Promise<void> => {
    if (turned) p.push(pushGraphicsState(), concatTransformationMatrix(...threeD.front));
    // The soft edge fades the body (the geometry, its bevel, and the inner shadow), not the text,
    // as in the SVG paint. The blur fades the body after the soft edge.
    const drawBody = (): Promise<void> =>
      drawSoftEdge(item, p, local, groupScale, async () => {
        p.push(pushGraphicsState());
        p.matrix(item.matrix);
        if (threeD) fillFacets(threeD.under, p);
        await drawGeometry(threeDFaceItem(item, threeD), p, local, groupScale);
        if (threeD) fillFacets(threeD.over, p);
        p.push(popGraphicsState());
        // Over the geometry and under the text, in slide space.
        if (shown) await drawInnerShadow(item, p, local);
      });
    if (shown) await drawBlur(item, p, local, groupScale, drawBody);
    else await drawBody();
    // An empty body's caret line draws nothing: no operators for it.
    if (item.kind === 'shape' && item.text && item.textMatrix && hasCharacters(item.text)) {
      p.push(pushGraphicsState());
      p.matrix(item.textMatrix);
      await drawText(item.text, p, local);
      p.push(popGraphicsState());
    }
    if (turned) p.push(popGraphicsState());
  };
  // A face turned edge-on paints only its extrusion: it has no area to paint or to copy.
  if (threeD && !threeD.showsFace) {
    fillFacets(threeD.behind, p);
    return;
  }
  // The reflection, shadows, glow, and extrusion paint under the content, in the SVG paint's
  // order; the shadows and the glow take the camera's turn of the face.
  const face = faceMatrix(threeD);
  // A turned extrusion casts its shadows and its glow from its back, as in the SVG paint.
  const back = threeD?.backOffset;
  const cast: Matrix = back
    ? [face[0], face[1], face[2], face[3], face[4] + back.x, face[5] + back.y]
    : face;
  if (shown) await drawReflection(item, p, drawContent);
  if (shown) await drawShadows(item, p, local, cast);
  if (shown) await drawGlow(item, p, local, cast);
  if (threeD) fillFacets(threeD.behind, p);
  await drawContent();
}

/** Map layout's fidelity notes to PDF diagnostics, with fixed codes. */
function layoutDiagnostics(layout: SlideLayout, diag: SlideDiagnostics): void {
  const after = (note: string, prefix: string): string => clip(note.slice(prefix.length));
  for (const note of layout.diagnostics) {
    if (note.startsWith('approximation: '))
      diag.add(
        'layout.approximation',
        'information',
        `Engine approximation: ${after(note, 'approximation: ')}.`
      );
    else if (note.startsWith('font substituted: '))
      diag.add(
        'font.substituted',
        'information',
        `Font substituted: ${after(note, 'font substituted: ')}.`
      );
    else if (note.startsWith('unknown preset geometry: '))
      diag.add(
        'geometry.unknown-preset',
        'approximation',
        `Unknown preset geometry drawn as a rectangle: ${after(note, 'unknown preset geometry: ')}.`
      );
    else if (note.startsWith('effect not rendered: '))
      diag.add(
        'effect.not-rendered',
        'unsupported',
        `Effect not rendered: ${after(note, 'effect not rendered: ')}.`
      );
  }
}

/** Paint one slide layout onto a new page. */
export async function paintSlide(
  layout: SlideLayout,
  res: PaintResources,
  diag: SlideDiagnostics
): Promise<void> {
  const { width, height } = layout.slide;
  const page = res.doc.addPage([width / EMU_PER_PT, height / EMU_PER_PT]);
  const p = new PageOps(page, res, diag);
  layoutDiagnostics(layout, diag);
  // EMU, y down → points, y up.
  p.push(
    pushGraphicsState(),
    concatTransformationMatrix(1 / EMU_PER_PT, 0, 0, -1 / EMU_PER_PT, 0, height / EMU_PER_PT)
  );
  const scope: FillScope = { sourcePart: layout.slide.backgroundPart };
  // A page is white, as the SVG paint's background without a fill.
  await paintFill(
    layout.slide.background,
    width,
    height,
    () => traceRect(0, 0, width, height, p),
    p,
    scope,
    'the background'
  );
  for (const item of layout.items) await drawItem(item, p, scope, 1);
  p.push(popGraphicsState());
  p.flush();
}
