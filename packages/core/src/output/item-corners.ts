// The slide-space corners that an item can paint into: its box, geometry, outline, and text
// frame. Effect filter regions and 3-D faces grow to these corners.

import {
  apply as applyMatrix,
  hasCharacters,
  isOpenPath,
  lineEndReach,
  multiply as multiplyMatrix,
  type LayoutItem,
  type Matrix,
} from '../layout/index.ts';
import { boundsOf, cornersUnder, grow, PX } from './effect-transforms.ts';

type BoxItem = Extract<LayoutItem, { kind: 'shape' } | { kind: 'picture' }>;

/**
 * How far the outline can paint past the shape's geometry, in slide space: a picture's outline
 * lies outside it, a shape's is centered, a miter join reaches up to half the miter limit times
 * the width, and line ends reach further. Effect filter regions grow by this, so they never cut
 * the outline off.
 */
function outlineReach(item: BoxItem): number {
  const line = item.element.line;
  if (!line || line.fill.type === 'none') return 0;
  const width = Math.max(line.width, PX);
  const join = line.join === 'miter' ? Math.max(2, line.miterLimit / 2) : 2;
  const ends =
    item.kind === 'shape' && item.geometry.paths.some((p) => p.stroke && isOpenPath(p.commands));
  return Math.max(width * join, ends ? lineEndReach(line) : 0);
}

/**
 * The slide-space corners of a shape's box and geometry, with its outline. Geometry can reach
 * past the box (a callout tail, `a:custGeom`); a Bezier curve lies inside its control points.
 */
export function boxCorners(item: BoxItem): { x: number; y: number }[] {
  const { cx, cy } = item.element.xfrm;
  const m = item.matrix;
  const points = cornersUnder(m, cx, cy);
  for (const path of item.geometry.paths)
    for (const c of path.commands) {
      if (c.op === 'Z') continue;
      if (c.op === 'C') points.push(applyMatrix(m, c.x1, c.y1), applyMatrix(m, c.x2, c.y2));
      else if (c.op === 'Q') points.push(applyMatrix(m, c.x1, c.y1));
      points.push(applyMatrix(m, c.x, c.y));
    }
  const b = grow(boundsOf(points), outlineReach(item));
  return [
    { x: b.x, y: b.y },
    { x: b.r, y: b.y },
    { x: b.r, y: b.b },
    { x: b.x, y: b.b },
  ];
}

/** Slide-space corners of a shape's box, its outline, and its text frame (text can overflow the box). */
export function contentCorners(item: BoxItem): { x: number; y: number }[] {
  const out = boxCorners(item);
  if (item.kind === 'shape' && item.text && item.textMatrix && hasCharacters(item.text)) {
    const b = item.text.box;
    const h = Math.max(b.height, item.text.contentHeight);
    const frame: Matrix = [1, 0, 0, 1, b.x, b.y];
    // Rotated text can extend either way: take the box as a square of its larger side.
    const side = item.text.rotation % 180 !== 0 ? Math.max(b.width, h) : 0;
    out.push(...cornersUnder(multiplyMatrix(item.textMatrix, frame), side || b.width, side || h));
  }
  return out;
}
