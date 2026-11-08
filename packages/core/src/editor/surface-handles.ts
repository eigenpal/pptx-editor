// The selection handles the surface paints in its overlay: the two end points of a single line,
// else the eight box handles and the rotation handle of one shape, or the box handles of the
// common bounds of several. A table frame has no rotation handle: tables do not rotate, as in
// slide editors. Its right and bottom edges are grid lines that resize the last column or row,
// so it has no east or south side handle there: one edge does one thing. Moved out of
// `surface.ts` (the 1000-line cap).

import { apply, type LayoutItem, type Matrix } from '../layout/index.ts';
import { SVG_NS } from '../output/index.ts';
import { rotationHandle, unionBounds } from '../automation/commands/placement.ts';
import { elbowCorners, segmentAxis, segmentMovable } from '../automation/commands/connectors.ts';
import { HANDLE_POS, HANDLES } from './surface-drag.ts';

/** A table frame: it has a grid, and no rotation handle. */
const isTable = (item: LayoutItem): boolean => item.kind === 'frame' && !!item.table;

/**
 * The handles of `selected`, in slide units (`px` is one CSS pixel): the end points of `line`
 * when one line is selected, with a handle on each segment of an elbow or curve connector
 * (titled `segmentLabel`), else the box handles.
 */
export function handleElements(
  doc: Document,
  selected: readonly LayoutItem[],
  line: { readonly start: { x: number; y: number }; readonly end: { x: number; y: number } } | null,
  px: number,
  segmentLabel = ''
): Element[] {
  const parts: Element[] = [];
  if (line) {
    if (selected.length === 1) parts.push(...segmentHandles(doc, selected[0], px, segmentLabel));
    for (const [handle, p] of [
      ['start', line.start],
      ['end', line.end],
    ] as const) {
      const node = doc.createElementNS(SVG_NS, 'circle');
      node.setAttribute('class', 'pptx-surface__handle pptx-surface__handle--point');
      node.setAttribute('data-handle', handle);
      node.setAttribute('cx', String(p.x));
      node.setAttribute('cy', String(p.y));
      node.setAttribute('r', String(px * 5));
      node.setAttribute('stroke-width', String(px));
      parts.push(node);
    }
    return parts;
  }
  if (selected.length === 0) return parts;
  const single = selected.length === 1 ? selected[0] : undefined;
  const u = unionBounds(selected);
  // One shape: handles on its own (rotated) box. Several: on their common bounds.
  const m: Matrix = single ? single.matrix : [1, 0, 0, 1, u.x, u.y];
  // A table's handles sit on its drawn grid, which rows that grow can make taller than xfrm.
  const grid = single?.kind === 'frame' ? single.table : undefined;
  const cx = grid ? grid.width : single ? single.element.xfrm.cx : u.r - u.x;
  const cy = grid ? grid.height : single ? single.element.xfrm.cy : u.b - u.y;
  for (const handle of HANDLES) {
    if (handle === 'rot' && (!single || isTable(single))) continue;
    if ((handle === 'e' || handle === 's') && single && isTable(single)) continue;
    const [hx, hy] = HANDLE_POS[handle];
    const p = apply(m, hx * cx, hy * cy);
    const hp = handle === 'rot' && single ? rotationHandle(single, 22 * px) : p;
    const node = doc.createElementNS(SVG_NS, handle === 'rot' ? 'circle' : 'rect');
    node.setAttribute('class', `pptx-surface__handle pptx-surface__handle--${handle}`);
    node.setAttribute('data-handle', handle);
    if (handle === 'rot') {
      node.setAttribute('cx', String(hp.x));
      node.setAttribute('cy', String(hp.y));
      node.setAttribute('r', String(px * 5));
    } else {
      node.setAttribute('x', String(p.x - px * 4));
      node.setAttribute('y', String(p.y - px * 4));
      node.setAttribute('width', String(px * 8));
      node.setAttribute('height', String(px * 8));
    }
    node.setAttribute('stroke-width', String(px));
    parts.push(node);
  }
  return parts;
}

/** Segments shorter than this, in CSS pixels, show no handle: it would cover the ends. */
const SHORTEST_SEGMENT = 24;

type Point = { readonly x: number; readonly y: number };

/** Points along an item's first drawn path on the slide, 32 for each curve. */
function drawnPoints(item: LayoutItem): Point[] {
  if (item.kind !== 'shape') return [];
  const out: Point[] = [];
  let at: Point = { x: 0, y: 0 };
  for (const c of item.geometry.paths[0]?.commands ?? []) {
    if (c.op === 'Z') continue;
    if (c.op === 'C' || c.op === 'Q') {
      // A quadratic as the cubic that draws it: each control two thirds of the way to its one.
      const q = (s: number, k: number): number => s + ((k - s) * 2) / 3;
      const a = c.op === 'C' ? { x: c.x1, y: c.y1 } : { x: q(at.x, c.x1), y: q(at.y, c.y1) };
      const b = c.op === 'C' ? { x: c.x2, y: c.y2 } : { x: q(c.x, c.x1), y: q(c.y, c.y1) };
      for (let k = 1; k <= 32; k += 1) {
        const t = k / 32;
        const u = 1 - t;
        const f = (s: number, p: number, r: number, e: number): number =>
          u * u * u * s + 3 * u * u * t * p + 3 * u * t * t * r + t * t * t * e;
        out.push(apply(item.matrix, f(at.x, a.x, b.x, c.x), f(at.y, a.y, b.y, c.y)));
      }
    } else out.push(apply(item.matrix, c.x, c.y));
    at = { x: c.x, y: c.y };
  }
  return out;
}

/** The point of `points` nearest `p`, or undefined when there is none. */
function nearest(points: readonly Point[], p: Point): Point | undefined {
  let best: Point | undefined;
  for (const q of points)
    if (!best || Math.hypot(q.x - p.x, q.y - p.y) < Math.hypot(best.x - p.x, best.y - p.y))
      best = q;
  return best;
}

/**
 * A diamond at the middle of each segment of an elbow or curve connector. A drag moves the
 * segment across (`moveElbowSegment`); on an end segment it adds a bend.
 */
function segmentHandles(doc: Document, item: LayoutItem, px: number, label: string): Element[] {
  const corners = elbowCorners(item);
  if (!corners) return [];
  // A curve bends between the corners: its handles sit on the drawn curve, nearest each middle.
  const curve = item.kind === 'shape' && item.geometry.paths[0]?.commands.some((c) => c.op === 'C');
  const drawn = curve ? drawnPoints(item) : [];
  const out: Element[] = [];
  for (let i = 0; i < corners.length - 1; i += 1) {
    const a = apply(item.parentMatrix, corners[i].x, corners[i].y);
    const b = apply(item.parentMatrix, corners[i + 1].x, corners[i + 1].y);
    if (Math.hypot(b.x - a.x, b.y - a.y) < SHORTEST_SEGMENT * px) continue;
    // A segment whose drag would need more segments than a preset draws shows no handle.
    if (!segmentMovable(item, corners, i)) continue;
    const mid = { x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 };
    const m = nearest(drawn, mid) ?? mid;
    const r = px * 5;
    const axis = segmentAxis(a, b);
    const titled = (node: Element): Element => {
      node.setAttribute('data-handle', `segment${i}`);
      const title = doc.createElementNS(SVG_NS, 'title');
      title.textContent = label;
      node.append(title);
      return node;
    };
    // A wider, unseen hit area first, so a finger or a quick pointer finds the small diamond.
    const hit = doc.createElementNS(SVG_NS, 'circle');
    hit.setAttribute('class', `pptx-surface__handle-hit pptx-surface__handle--segment-${axis}`);
    hit.setAttribute('cx', String(m.x));
    hit.setAttribute('cy', String(m.y));
    hit.setAttribute('r', String(px * 11));
    const node = doc.createElementNS(SVG_NS, 'path');
    node.setAttribute(
      'class',
      `pptx-surface__handle pptx-surface__handle--adjust pptx-surface__handle--segment-${axis}`
    );
    node.setAttribute(
      'd',
      `M${m.x} ${m.y - r}L${m.x + r} ${m.y}L${m.x} ${m.y + r}L${m.x - r} ${m.y}Z`
    );
    node.setAttribute('stroke-width', String(px));
    out.push(titled(hit), titled(node));
  }
  return out;
}

/**
 * The dashed common box of several selected shapes, in slide units (`px` is one CSS pixel). The
 * surface paints it under the member outlines (`memberOutline`), so each shape's own outline
 * stays visible where it meets the common box.
 */
export function commonOutline(
  doc: Document,
  selected: readonly LayoutItem[],
  px: number
): SVGRectElement {
  const u = unionBounds(selected);
  const box = doc.createElementNS(SVG_NS, 'rect') as SVGRectElement;
  box.setAttribute('class', 'pptx-surface__group-outline');
  box.setAttribute('x', String(u.x));
  box.setAttribute('y', String(u.y));
  box.setAttribute('width', String(u.r - u.x));
  box.setAttribute('height', String(u.b - u.y));
  box.setAttribute('stroke-width', String(px));
  box.setAttribute('stroke-dasharray', `${px * 2} ${px * 2}`);
  return box;
}

/** How far outside a shape's box the outline of one shape of several sits, in CSS pixels. */
export const MEMBER_OUTSET = 1;

/**
 * The outline of one shape of a multiple selection, in slide units (`px` is one CSS pixel): a
 * 1 px line `MEMBER_OUTSET` px outside the shape's box, so it never covers the shape's text or
 * border, at any zoom. The common box of the selection carries the handles.
 */
export function memberOutline(doc: Document, item: LayoutItem, px: number): SVGGElement {
  const { cx, cy } = item.element.xfrm;
  const [a, b, c, d] = item.matrix;
  // The item's own units: a shape in a scaled group has a scaled frame.
  const sx = Math.hypot(a, b) || 1;
  const sy = Math.hypot(c, d) || 1;
  const ox = (px * MEMBER_OUTSET) / sx;
  const oy = (px * MEMBER_OUTSET) / sy;
  const g = doc.createElementNS(SVG_NS, 'g') as SVGGElement;
  g.setAttribute('transform', `matrix(${item.matrix.join(' ')})`);
  const rect = doc.createElementNS(SVG_NS, 'rect');
  rect.setAttribute('class', 'pptx-surface__outline pptx-surface__outline--member');
  rect.setAttribute('x', String(-ox));
  rect.setAttribute('y', String(-oy));
  rect.setAttribute('width', String(cx + 2 * ox));
  rect.setAttribute('height', String(Math.max(cy, 1) + 2 * oy));
  rect.setAttribute('stroke-width', String(px / Math.sqrt(sx * sy)));
  g.append(rect);
  return g;
}

/** An item's box outline in slide space, `width` wide. */
export function boxOutline(doc: Document, item: LayoutItem, width: number): SVGGElement {
  const { cx, cy } = item.element.xfrm;
  const g = doc.createElementNS(SVG_NS, 'g') as SVGGElement;
  g.setAttribute('transform', `matrix(${item.matrix.join(' ')})`);
  const rect = doc.createElementNS(SVG_NS, 'rect');
  rect.setAttribute('class', 'pptx-surface__outline');
  rect.setAttribute('width', String(cx));
  rect.setAttribute('height', String(Math.max(cy, 1)));
  rect.setAttribute('stroke-width', String(width));
  g.append(rect);
  return g;
}
