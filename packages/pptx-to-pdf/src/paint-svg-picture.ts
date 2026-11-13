/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// SVG pictures as PDF vector paint. The editor draws an SVG picture as vector art when it is a
// plain icon; `svgPictureDrawing` reads the same icons into paths and solid paints, and this
// module writes them as PDF paths. A picture the reader refuses draws its PNG fallback.

import {
  drawObject,
  popGraphicsState,
  pushGraphicsState,
  setFillingRgbColor,
  setLineCap,
  setLineJoin,
  setLineWidth,
  setStrokingRgbColor,
  LineCapStyle,
  LineJoinStyle,
  PDFNumber,
  PDFOperator,
  PDFOperatorNames,
} from 'pdf-lib';
import {
  imageMime,
  svgPictureDrawing,
  type SvgPictureDrawing,
  type SvgPictureItem,
} from '@pptx-editor.dev/core/editor';
import { binaryPartOf } from '@pptx-editor.dev/core/store';
import {
  clipEvenOdd,
  clipNonZero,
  endPath,
  fillEvenOdd,
  tracePath,
  type PageOps,
} from './page-ops.ts';

const drawings = new WeakMap<PageOps['res'], Map<string, SvgPictureDrawing | null>>();

/** The drawing of the SVG picture part `target`, read once per document; null when not drawn. */
export function svgDrawingFor(target: string, p: PageOps): SvgPictureDrawing | null {
  let cache = drawings.get(p.res);
  if (!cache) drawings.set(p.res, (cache = new Map()));
  const known = cache.get(target);
  if (known !== undefined) return known;
  const bytes = imageMime(target) === 'image/svg+xml' ? binaryPartOf(p.res.pkg, target) : undefined;
  const drawing = (bytes && svgPictureDrawing(bytes)) || null;
  cache.set(target, drawing);
  return drawing;
}

const CAP = { butt: LineCapStyle.Butt, round: LineCapStyle.Round, square: LineCapStyle.Projecting };
const JOIN = { miter: LineJoinStyle.Miter, round: LineJoinStyle.Round, bevel: LineJoinStyle.Bevel };
const rgb = (c: { r: number; g: number; b: number }): [number, number, number] => [
  c.r / 255,
  c.g / 255,
  c.b / 255,
];

function drawItems(items: readonly SvgPictureItem[], bbox: number[], p: PageOps): void {
  for (const item of items) {
    if (item.kind === 'group') {
      p.push(pushGraphicsState());
      if (item.clip) {
        tracePath(item.clip.path, p);
        p.push(item.clip.evenOdd ? clipEvenOdd() : clipNonZero(), endPath());
      }
      if (item.opacity < 1) {
        // A layer: its items draw together, then the layer takes the opacity.
        const ops = p.captureNow(() => drawItems(item.items, bbox, p));
        const form = p.form(ops, bbox, { S: 'Transparency' });
        p.alpha(item.opacity, item.opacity);
        p.push(drawObject(p.name('XObject', form)));
      } else drawItems(item.items, bbox, p);
      p.push(popGraphicsState());
      continue;
    }
    const { fill, stroke } = item;
    p.push(pushGraphicsState());
    p.matrix(item.matrix);
    p.alpha(fill?.color.a ?? 1, stroke?.color.a ?? 1);
    if (fill) p.push(setFillingRgbColor(...rgb(fill.color)));
    if (stroke)
      p.push(
        setStrokingRgbColor(...rgb(stroke.color)),
        setLineWidth(stroke.width),
        setLineCap(CAP[stroke.cap]),
        setLineJoin(JOIN[stroke.join]),
        PDFOperator.of(PDFOperatorNames.SetLineMiterLimit, [PDFNumber.of(stroke.miterLimit)])
      );
    tracePath(item.path, p);
    const evenOdd = fill?.evenOdd ?? false;
    const op = fill
      ? stroke
        ? evenOdd
          ? PDFOperatorNames.FillEvenOddAndStroke
          : PDFOperatorNames.FillNonZeroAndStroke
        : evenOdd
          ? null
          : PDFOperatorNames.FillNonZero
      : PDFOperatorNames.StrokePath;
    p.push(op === null ? fillEvenOdd() : PDFOperator.of(op), popGraphicsState());
  }
}

/**
 * The largest magnitude of a number that the view box's matrix writes. The drawing's own
 * numbers are bounded by `svgPictureDrawing`; a tiny view box over a large frame could still
 * scale past what a PDF real holds as a plain decimal.
 */
const MAX_MATRIX_NUMBER = 1e12;

/**
 * Draw an SVG picture's drawing over the box `x`, `y`, `w` × `h` (slide EMU, y down): its view
 * box stretches over the box, as the editor's image draws with `preserveAspectRatio="none"`,
 * and the view box clips. `alpha` is the picture's transparency. Draws nothing and returns
 * false when the view box's matrix holds a number out of range.
 */
export function drawSvgPicture(
  drawing: SvgPictureDrawing,
  x: number,
  y: number,
  w: number,
  h: number,
  alpha: number,
  p: PageOps
): boolean {
  const [vx, vy, vw, vh] = drawing.viewBox;
  const sx = w / vw;
  const sy = h / vh;
  const raw = [sx, 0, 0, sy, x - vx * sx, y - vy * sy];
  if (!raw.every((v) => Number.isFinite(v) && Math.abs(v) <= MAX_MATRIX_NUMBER)) return false;
  // Rounding noise near zero is written as zero, so no decimal runs long.
  const [a, b, c, d, e, f] = raw.map((v) => (Math.abs(v) < 1e-9 ? 0 : v)) as [
    number,
    number,
    number,
    number,
    number,
    number,
  ];
  const matrix = [a, b, c, d, e, f] as const;
  const bbox = [vx, vy, vx + vw, vy + vh];
  p.push(pushGraphicsState());
  p.matrix(matrix);
  tracePath(
    [
      { op: 'M', x: vx, y: vy },
      { op: 'L', x: vx + vw, y: vy },
      { op: 'L', x: vx + vw, y: vy + vh },
      { op: 'L', x: vx, y: vy + vh },
      { op: 'Z' },
    ],
    p
  );
  p.push(clipNonZero(), endPath());
  if (alpha < 1) {
    const ops = p.captureNow(() => drawItems(drawing.items, bbox, p));
    const form = p.form(ops, bbox, { S: 'Transparency' });
    p.alpha(alpha, alpha);
    p.push(drawObject(p.name('XObject', form)));
  } else drawItems(drawing.items, bbox, p);
  p.push(popGraphicsState());
  return true;
}
