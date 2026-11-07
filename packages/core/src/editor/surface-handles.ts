// The selection handles the surface paints in its overlay: the two end points of a single line,
// else the eight box handles and the rotation handle of one shape, or the box handles of the
// common bounds of several. A table frame has no rotation handle: tables do not rotate, as in
// slide editors. Its right and bottom edges are grid lines that resize the last column or row,
// so it has no east or south side handle there: one edge does one thing. Moved out of
// `surface.ts` (the 1000-line cap).

import { apply, type LayoutItem, type Matrix } from '../layout/index.ts';
import { SVG_NS } from '../output/index.ts';
import { rotationHandle, unionBounds } from '../automation/commands/placement.ts';
import { HANDLE_POS, HANDLES } from './surface-drag.ts';

/** A table frame: it has a grid, and no rotation handle. */
const isTable = (item: LayoutItem): boolean => item.kind === 'frame' && !!item.table;

/**
 * The handles of `selected`, in slide units (`px` is one CSS pixel): the end points of `line`
 * when one line is selected, else the box handles.
 */
export function handleElements(
  doc: Document,
  selected: readonly LayoutItem[],
  line: { readonly start: { x: number; y: number }; readonly end: { x: number; y: number } } | null,
  px: number
): Element[] {
  const parts: Element[] = [];
  if (line) {
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
  if (!single) {
    const box = doc.createElementNS(SVG_NS, 'rect');
    box.setAttribute('class', 'pptx-surface__group-outline');
    box.setAttribute('x', String(u.x));
    box.setAttribute('y', String(u.y));
    box.setAttribute('width', String(cx));
    box.setAttribute('height', String(cy));
    box.setAttribute('stroke-width', String(px));
    box.setAttribute('stroke-dasharray', `${px * 2} ${px * 2}`);
    parts.push(box);
  }
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
