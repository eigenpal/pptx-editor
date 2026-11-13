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

/**
 * How far an item's own copies (shadows, glow, reflection) can paint past its content, in slide
 * space, for content whose larger side is `side`. A scaled shadow or a reflection can move a
 * copy of the whole content, so they reach past it by its side.
 */
function effectReach(item: LayoutItem, side: number): number {
  const fx = item.element.kind === 'graphicFrame' ? undefined : item.element.effects;
  if (!fx) return 0;
  let reach = 0;
  for (const s of [fx.outerShadow, fx.secondShadow]) {
    if (!s) continue;
    const scale = Math.max(Math.abs(s.sx), Math.abs(s.sy), 1);
    const skew = Math.abs(s.kx) > 0 || Math.abs(s.ky) > 0 ? side : 0;
    reach = Math.max(reach, s.dist + s.blur + side * (scale - 1) + skew);
  }
  if (fx.glow) reach = Math.max(reach, fx.glow.radius);
  if (fx.reflection) reach = Math.max(reach, fx.reflection.dist + fx.reflection.blur + side);
  if (fx.blur?.grow) reach = Math.max(reach, fx.blur.radius);
  return reach;
}

/**
 * The slide-space corners of a group's box and of everything its children paint, their own
 * effects included. The group's own effects copy and filter this content.
 */
export function groupContentCorners(
  group: Extract<LayoutItem, { kind: 'group' }>
): { x: number; y: number }[] {
  const out = cornersUnder(group.matrix, group.element.xfrm.cx, group.element.xfrm.cy);
  for (const child of group.children) out.push(...itemCorners(child));
  return out;
}

/**
 * The slide-space corners of everything an item paints: a shape's or picture's content, a
 * frame's box, table, picture, diagram, or chart, and the children of a group. An item's own
 * shadows, glow, and reflection widen its corners, so the effects of a group that holds it copy
 * them whole.
 */
export function itemCorners(item: LayoutItem): { x: number; y: number }[] {
  let out: { x: number; y: number }[];
  switch (item.kind) {
    case 'shape':
    case 'picture':
      out = contentCorners(item);
      break;
    case 'group':
      out = groupContentCorners(item);
      break;
    case 'frame': {
      const { cx, cy } = item.element.xfrm;
      // A table's rows grow with their text, past the frame's box.
      const w = Math.max(cx, item.table?.width ?? 0);
      const h = Math.max(cy, item.table?.height ?? 0);
      out = cornersUnder(item.matrix, w, h);
      if (item.poster) out.push(...itemCorners(item.poster));
      for (const child of item.drawing ?? item.chart ?? []) out.push(...itemCorners(child));
      break;
    }
  }
  const b = boundsOf(out);
  const reach = effectReach(item, Math.max(b.r - b.x, b.b - b.y));
  if (!(reach > 0)) return out;
  const g = grow(b, reach);
  return [
    { x: g.x, y: g.y },
    { x: g.r, y: g.y },
    { x: g.r, y: g.b },
    { x: g.x, y: g.b },
  ];
}
