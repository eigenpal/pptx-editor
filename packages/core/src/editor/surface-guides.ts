// The surface side of smart guides: what a drag snaps against (the other shapes of its scope,
// by their rotated bounding boxes), which lines a resize handle moves, and the guides as
// overlay elements. The geometry itself is in `snapping.ts`.

import { apply, type LayoutItem, type SlideLayout } from '../layout/index.ts';
import { SVG_NS } from '../output/index.ts';
import { parentGroupOf, scopeItems, unionBounds } from '../automation/commands/placement.ts';
import {
  ALL_LINES,
  moveBounds,
  resizeBounds,
  snapBox,
  type SnapBounds,
  type SnapGuide,
  type SnapLine,
  type SnapLines,
  type SnapResult,
} from './snapping.ts';

/** The snap distance, in screen pixels. */
export const SNAP_PX = 6;

/** What a drag snaps against: the other visible shapes of its scope, and the slide. */
export interface SnapContext {
  readonly targets: readonly SnapBounds[];
  readonly slide: { readonly width: number; readonly height: number };
}

/** The snap context of a drag of `ids`: their siblings (in the slide or the entered group). */
export function snapContext(layout: SlideLayout, ids: readonly string[]): SnapContext {
  const scope = ids[0] ? parentGroupOf(layout.items, ids[0]) : undefined;
  const targets = scopeItems(layout.items, scope)
    .filter((i) => i.element.layer === 'slide' && !ids.includes(i.element.nodeId))
    .map((i) => unionBounds([i]));
  return { targets, slide: { width: layout.slide.width, height: layout.slide.height } };
}

/** The bounds lines a handle at (`hx`, `hy`) of an axis-aligned box moves. */
export function handleLines(hx: number, hy: number): SnapLines {
  const one = (h: number): SnapLine[] => (h === 0 ? ['start'] : h === 1 ? ['end'] : []);
  return { x: one(hx), y: one(hy) };
}

/**
 * The slide-space bounds lines a resize handle of one shape moves, or null when the shape is
 * turned by an angle that is not a multiple of 90°: its edges are not the bounds' edges.
 */
export function resizeLines(item: LayoutItem, hx: number, hy: number): SnapLines | null {
  const [a, b, c, d] = item.matrix;
  const tiny = (v: number): boolean => Math.abs(v) < 1e-9;
  if (!((tiny(b) && tiny(c)) || (tiny(a) && tiny(d)))) return null;
  const { cx, cy } = item.element.xfrm;
  const p = apply(item.matrix, hx * cx, hy * cy);
  const u = unionBounds([item]);
  const one = (v: number, lo: number, hi: number): SnapLine[] => {
    const eps = 1 + (hi - lo) * 1e-6;
    if (hi - lo <= eps) return [];
    return Math.abs(v - lo) <= eps ? ['start'] : Math.abs(v - hi) <= eps ? ['end'] : [];
  };
  return { x: one(p.x, u.x, u.r), y: one(p.y, u.y, u.b) };
}

/**
 * A move of the drag bounds by (`dx`, `dy`), snapped when `tolerance` is a number (null: Alt
 * or View > Smart guides is off). Shift keeps the move on one axis, which does not snap.
 */
export function snapMove(
  ctx: SnapContext,
  bounds: SnapBounds,
  dx: number,
  dy: number,
  shift: boolean,
  tolerance: number | null
): SnapResult {
  const lock = shift ? (Math.abs(dx) >= Math.abs(dy) ? 'y' : 'x') : null;
  const mx = lock === 'x' ? 0 : dx;
  const my = lock === 'y' ? 0 : dy;
  if (tolerance === null) return { dx: mx, dy: my, guides: [] };
  const lines = lock ? { ...ALL_LINES, [lock]: [] } : ALL_LINES;
  const r = snapBox({ ...ctx, moving: moveBounds(bounds, mx, my), tolerance, lines });
  return { dx: mx + r.dx, dy: my + r.dy, guides: r.guides };
}

/**
 * A resize by a handle at (`hx`, `hy`) dragged by (`dx`, `dy`), with the dragged edges snapped.
 * It does not snap when Shift keeps the aspect ratio at a corner, or for a turned shape.
 */
export function snapResize(
  ctx: SnapContext,
  items: readonly LayoutItem[],
  bounds: SnapBounds,
  [hx, hy]: readonly [number, number],
  [dx, dy]: readonly [number, number],
  aspect: boolean,
  tolerance: number | null
): SnapResult {
  const corner = aspect && hx !== 0.5 && hy !== 0.5;
  // One shape resizes in its own frame; several scale their common bounds.
  const lines =
    tolerance === null || corner
      ? null
      : items.length === 1
        ? resizeLines(items[0], hx, hy)
        : handleLines(hx, hy);
  if (!lines || tolerance === null) return { dx, dy, guides: [] };
  const r = snapBox({ ...ctx, moving: resizeBounds(bounds, lines, dx, dy), tolerance, lines });
  return { dx: dx + r.dx, dy: dy + r.dy, guides: r.guides };
}

/** The guides as overlay elements, one device pixel wide (`px` is one CSS pixel in EMU). */
export function guideElements(doc: Document, guides: readonly SnapGuide[], px: number): Element[] {
  const width = String(px / (doc.defaultView?.devicePixelRatio || 1));
  const line = (cls: string, x1: number, y1: number, x2: number, y2: number): Element => {
    const node = doc.createElementNS(SVG_NS, 'line');
    node.setAttribute('class', cls);
    node.setAttribute('x1', String(x1));
    node.setAttribute('y1', String(y1));
    node.setAttribute('x2', String(x2));
    node.setAttribute('y2', String(y2));
    node.setAttribute('stroke-width', width);
    return node;
  };
  // Points on the slide from (along, across) on a guide's axis.
  const xy = (axis: 'x' | 'y', along: number, across: number): [number, number] =>
    axis === 'x' ? [along, across] : [across, along];
  const out: Element[] = [];
  for (const g of guides) {
    if (g.kind === 'align') {
      // Dashed; the equal-gap bars stay solid.
      const node = line(
        'pptx-surface__guide',
        ...xy(g.axis, g.at, g.from),
        ...xy(g.axis, g.at, g.to)
      );
      node.setAttribute('stroke-dasharray', `${px * 4} ${px * 3}`);
      out.push(node);
      continue;
    }
    // An equal gap: a bar between the two shapes with a tick at each end.
    const tick = px * 5;
    const cls = 'pptx-surface__guide pptx-surface__guide--gap';
    out.push(line(cls, ...xy(g.axis, g.from, g.at), ...xy(g.axis, g.to, g.at)));
    for (const end of [g.from, g.to])
      out.push(line(cls, ...xy(g.axis, end, g.at - tick), ...xy(g.axis, end, g.at + tick)));
  }
  return out;
}
