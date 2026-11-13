/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Bevels, contours, and extrusion as PDF paint: the polygons of the installed 3-D renderer
// (`@pptx-editor.dev/pro/drawing`), the same ones the SVG paint fills. The editor's renderer
// gives them, so the two outputs agree. Without the renderer, a 3-D shape draws flat.

import {
  fill,
  popGraphicsState,
  pushGraphicsState,
  setFillingRgbColor,
  setLineJoin,
  setLineWidth,
  setStrokingRgbColor,
  LineJoinStyle,
  PDFOperator,
  PDFOperatorNames,
} from 'pdf-lib';
import {
  drawingRenderers,
  type LayoutItem,
  type Matrix,
  type ThreeDFacet,
  type ThreeDPaint,
} from '@pptx-editor.dev/core/layout';
import { finiteFacet, threeDColor } from '@pptx-editor.dev/core/output';
import { rgb, tracePath, type PageOps } from './page-ops.ts';

type Drawable = Extract<LayoutItem, { kind: 'shape' } | { kind: 'picture' }>;

/** The 3-D paint of a shape or picture, or null when it draws flat. */
export function threeDOf(item: Drawable): ThreeDPaint | null {
  const effects = item.element.effects;
  const render = effects?.threeD ? drawingRenderers().effects?.threeD : undefined;
  if (!render || !effects) return null;
  return render({
    effects,
    paths: item.geometry.paths,
    matrix: item.matrix,
    width: item.element.xfrm.cx,
    height: item.element.xfrm.cy,
    color: threeDColor(item.element.fill),
  });
}

/** Whether the camera turns the face: its transform is not the identity. */
export function turnsFace(paint: ThreeDPaint | null): paint is ThreeDPaint {
  const id: Matrix = [1, 0, 0, 1, 0, 0];
  return paint !== null && paint.front.some((v, i) => Math.abs(v - id[i]!) > 1e-12);
}

/** The transform of the face and its effects: the camera's turn, else the identity. */
export function faceMatrix(paint: ThreeDPaint | null): Matrix {
  return paint?.front ?? [1, 0, 0, 1, 0, 0];
}

/** The width of the outline of an opaque facet, half a pixel in EMU, as in the SVG paint. */
const SEAM = 9_525 / 2;

/**
 * Fill facets in the current space, each in its color and alpha, as one nonzero path. An opaque
 * facet also has a thin outline of its own color, as in the SVG paint, so neighbors leave no
 * antialiased seam that shows the face below. A facet with a value that is not finite is
 * skipped, as the SVG paint skips it.
 */
export function fillFacets(facets: readonly ThreeDFacet[], p: PageOps): void {
  for (const f of facets) {
    const polygons = f.polygons.filter((poly) => poly.length >= 3);
    if (polygons.length === 0 || f.color.a <= 0 || !finiteFacet(f)) continue;
    const opaque = f.color.a >= 1;
    p.push(pushGraphicsState());
    p.alpha(f.color.a, 1);
    p.push(setFillingRgbColor(...rgb(f.color)));
    if (opaque)
      p.push(
        setStrokingRgbColor(...rgb(f.color)),
        setLineWidth(SEAM),
        setLineJoin(LineJoinStyle.Round)
      );
    tracePath(
      polygons.flatMap((poly) => [
        ...poly.map((q, i) => ({ op: i ? ('L' as const) : ('M' as const), x: q.x, y: q.y })),
        { op: 'Z' as const },
      ]),
      p
    );
    p.push(
      opaque ? PDFOperator.of(PDFOperatorNames.FillNonZeroAndStroke) : fill(),
      popGraphicsState()
    );
  }
}
