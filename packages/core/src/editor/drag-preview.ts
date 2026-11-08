// The preview of a shape drag (surface-drag.ts): each dragged shape laid out again at its new
// box, with nothing written, and painted in place of the shape on the slide. Text reflows in
// the new box, outlines keep their width, and a line's end points follow the pointer, as the
// release writes them. A resized shape takes the autofit the release writes (`relayoutItem`):
// an `a:spAutoFit` box sizes to its text, and `a:normAutofit` text scales to its box.
//
// A move only translates the shapes, so the preview places their layout as it is. A table
// resize lays the table out with the columns and rows the release writes (`resizeTable`), and
// the frame takes the laid-out table's size, as `refitTable` gives it. An object's picture
// (`poster`) takes the frame's new box. A frame the engine does not draw is stretched.
//
// A dragged group child is painted inside its real ancestor groups (with no other children),
// so it takes its group's fill (`a:grpFill`) and its group's scale for line widths, as on the
// slide. The preview holds no hit targets: the surface hits shapes through the layout.

import type { TransformItem } from '../contracts/index.ts';
import {
  boxMatrix,
  invert,
  multiply,
  type LayoutItem,
  type Matrix,
  type SlideLayout,
} from '../layout/index.ts';
import type { Xfrm } from '../cascade/index.ts';
import { paintSlide, type PaintOptions } from '../output/index.ts';
import { scaledRowHeights } from '../automation/commands/table-fit.ts';
import { findItem } from '../automation/commands/placement.ts';
import { followingConnectors } from '../automation/commands/connectors.ts';
import type { Drag } from './surface-drag.ts';

/** Lays out an item again at a new box in its parent space, or null when it cannot. */
export type Relayout = (item: LayoutItem, xfrm: Xfrm) => LayoutItem | null;

/** A file transform of `item` as a box in its parent space, flips included. */
export function parentBox(item: LayoutItem, t: TransformItem): Xfrm {
  const s = item.toSource;
  const el = item.element.xfrm;
  const box = s
    ? { x: (t.x - s.ox) * s.sx, y: (t.y - s.oy) * s.sy, cx: t.cx * s.sx, cy: t.cy * s.sy }
    : { x: t.x, y: t.y, cx: t.cx, cy: t.cy };
  return { ...box, rot: t.rot, flipH: t.flipH ?? el.flipH, flipV: t.flipV ?? el.flipV };
}

/** An item and its children placed by `m` (slide space → slide space). */
function moved(item: LayoutItem, m: Matrix): LayoutItem {
  const placed = { ...item, matrix: multiply(m, item.matrix) } as LayoutItem;
  if (placed.kind === 'shape' && placed.textMatrix)
    return { ...placed, textMatrix: multiply(m, placed.textMatrix) };
  if (placed.kind === 'group')
    return { ...placed, children: placed.children.map((c) => moved(c, m)) };
  // An object's picture is drawn in the frame's box: it moves with the frame.
  if (placed.kind === 'frame' && placed.poster)
    return { ...placed, poster: moved(placed.poster, m) as typeof placed.poster };
  return placed;
}

/**
 * A table frame laid out at a new box: the columns and rows scaled as the release scales them,
 * and the frame at the laid-out table's size.
 */
function resizedTable(
  item: Extract<LayoutItem, { kind: 'frame' }>,
  box: Xfrm,
  relayout: Relayout
): LayoutItem | null {
  const table = item.element.table;
  const laid = item.table;
  if (!table || !laid) return null;
  const old = item.element.xfrm;
  const fx = old.cx > 0 && Math.abs(box.cx - old.cx) > 1 ? box.cx / old.cx : 1;
  const columns = table.columns.map((w) => Math.max(1, Math.round(w * fx)));
  const mins = table.rows.map((r) => r.height);
  const drawn = table.rows.map((_, i) => (laid.rowY[i + 1] ?? 0) - (laid.rowY[i] ?? 0));
  const total = drawn.reduce((n, d) => n + d, 0);
  const heights =
    old.cy > 0 && Math.abs(box.cy - old.cy) > 1 && total > 0
      ? scaledRowHeights(drawn, mins, (total * box.cy) / old.cy)
      : mins;
  const rows = table.rows.map((r, i) => ({ ...r, height: Math.max(1, Math.round(heights[i]!)) }));
  const element = { ...item.element, table: { ...table, columns, rows } };
  const again = relayout({ ...item, element } as LayoutItem, box);
  if (again?.kind !== 'frame' || !again.table) return again;
  // The file box takes the laid-out table's size; a group draws it at its scale.
  const sx = item.toSource?.sx ?? 1;
  const sy = item.toSource?.sy ?? 1;
  const fitted = {
    ...box,
    cx: Math.round(again.table.width) * sx,
    cy: Math.round(again.table.height) * sy,
  };
  return {
    ...again,
    matrix: multiply(item.parentMatrix, boxMatrix(fitted)),
    element: { ...again.element, xfrm: fitted },
  };
}

/** A frame stretched from its old box onto `box`, with its element keeping the old size. */
function stretched(item: LayoutItem, box: Xfrm): LayoutItem {
  const old = item.element.xfrm;
  const sx = old.cx ? box.cx / old.cx : 1;
  const sy = old.cy ? box.cy / old.cy : 1;
  const to = multiply(item.parentMatrix, boxMatrix(box));
  return moved(item, multiply(multiply(to, [sx, 0, 0, sy, 0, 0]), invert(item.matrix)));
}

/** A connector with the preset and adjust values of its new route, if the transform has one. */
function routed(item: LayoutItem, t: TransformItem): LayoutItem {
  const route = t.connector;
  if (item.kind !== 'shape' || item.element.geometry.kind !== 'preset') return item;
  if (route?.preset === undefined && route?.adjust === undefined) return item;
  const geometry = {
    kind: 'preset' as const,
    preset: route.preset ?? item.element.geometry.preset,
    adjust: new Map(Object.entries(route.adjust ?? {})),
  };
  return { ...item, element: { ...item.element, geometry } };
}

/**
 * Preview the boxes `items`: keep them, and each shape laid out at its new box. The connectors
 * glued to the moved shapes follow them, as the release writes them.
 */
export function preview(d: Drag, items: readonly TransformItem[], relayout?: Relayout): void {
  const shown = new Map<string, LayoutItem>();
  for (const t of items) {
    const item = d.items.find((i) => i.item.element.nodeId === t.nodeId)?.item;
    if (!item) continue;
    const box = parentBox(item, t);
    shown.set(t.nodeId, laidOut(d, routed(item, t), box, relayout) ?? stretched(item, box));
  }
  if (relayout)
    for (const t of followingConnectors(d.slideItems, items)) {
      const item = findItem(d.slideItems, t.nodeId);
      const again = item ? relayout(routed(item, t), parentBox(item, t)) : null;
      if (again) shown.set(t.nodeId, again);
    }
  d.pending = items;
  d.shown = shown;
}

/** An item laid out again at `box`, or null when the preview stretches it instead. */
function laidOut(d: Drag, item: LayoutItem, box: Xfrm, relayout?: Relayout): LayoutItem | null {
  if (d.handle === 'move' || !relayout) return null;
  if (item.kind !== 'frame') return relayout(item, box);
  return item.table ? resizedTable(item, box, relayout) : relayout(item, box);
}

/**
 * Whether an item is a slide shape. Node ids are unique in one part only, so a layout or
 * master shape can have the node id of a dragged slide shape.
 */
const slideShape = (item: LayoutItem): boolean => item.element.layer === 'slide';

/** The layout the overlay draws while a drag previews: each dragged shape at its new box. */
export function previewLayout(layout: SlideLayout, d: Drag): SlideLayout {
  if (d.shown.size === 0) return layout;
  const map = (items: readonly LayoutItem[]): LayoutItem[] =>
    items.map((item) => {
      const shown = slideShape(item) ? d.shown.get(item.element.nodeId) : undefined;
      if (shown) return shown;
      return item.kind === 'group'
        ? ({ ...item, children: map(item.children) } as LayoutItem)
        : item;
    });
  return { ...layout, items: map(layout.items) };
}

/** `items` with only the shapes `shown` (at their preview) and the groups that hold them. */
function pruned(
  items: readonly LayoutItem[],
  shown: ReadonlyMap<string, LayoutItem>
): LayoutItem[] {
  const out: LayoutItem[] = [];
  for (const item of items) {
    const at = slideShape(item) ? shown.get(item.element.nodeId) : undefined;
    if (at) out.push(at);
    else if (item.kind === 'group') {
      const children = pruned(item.children, shown);
      if (children.length) out.push({ ...item, children } as LayoutItem);
    }
  }
  return out;
}

/** The hit target attributes the painter writes, which a preview must not repeat. */
const HIT_ATTRIBUTES = ['data-node-id', 'data-part', 'data-layer'] as const;

/**
 * The preview on the painted slide: each dragged shape hidden, with its preview painted after
 * it, so it keeps its place in the stacking order. `clear` puts the slide back.
 */
export class PaintedPreview {
  readonly #hidden = new Map<Element, string | null>();
  readonly #added: Element[] = [];
  #count = 0;

  /**
   * Paint the shapes `d.shown` over `slide`, the painted slide of `layout`. `paintLayout` adds
   * what the editor paints over a layout (placeholder prompts).
   */
  show(
    slide: Element | null,
    layout: SlideLayout,
    d: Drag,
    opts: PaintOptions,
    paintLayout?: (l: SlideLayout) => SlideLayout
  ): void {
    this.clear();
    if (!slide || d.shown.size === 0) return;
    const doc = slide.ownerDocument;
    this.#count += 1;
    const mini: SlideLayout = {
      ...layout,
      slide: { ...layout.slide, background: { type: 'none' } },
      items: pruned(layout.items, d.shown),
    };
    const painted = paintSlide(doc, paintLayout?.(mini) ?? mini, {
      ...opts,
      idPrefix: `${opts.idPrefix}d${this.#count}`,
    });
    const shapes = byNodeId(painted, d.shown);
    const defs = painted.querySelector('defs');
    let first = true;
    for (const [id, el] of byNodeId(slide, d.shown)) {
      this.#hidden.set(el, el.getAttribute('display'));
      el.setAttribute('display', 'none');
      const shape = shapes.get(id);
      if (!shape) continue;
      for (const node of [shape, ...shape.querySelectorAll('[data-node-id]')])
        for (const name of HIT_ATTRIBUTES) node.removeAttribute(name);
      const g = doc.createElementNS(slide.namespaceURI, 'g');
      g.setAttribute('data-kind', 'drag-preview');
      g.setAttribute('data-preview-of', id);
      // The preview's gradients, markers, and clip paths go with the first shape.
      if (first && defs) g.append(defs);
      first = false;
      g.append(shape);
      el.after(g);
      this.#added.push(g);
    }
  }
  /** Put every shape back where the slide painted it. */
  clear(): void {
    for (const [el, display] of this.#hidden) {
      if (display === null) el.removeAttribute('display');
      else el.setAttribute('display', display);
    }
    this.#hidden.clear();
    for (const g of this.#added.splice(0)) g.remove();
  }
}

/** The painted elements of the shapes `ids` in `root`, in one pass. */
function byNodeId(root: Element, ids: ReadonlyMap<string, unknown>): Map<string, Element> {
  const out = new Map<string, Element>();
  for (const el of root.querySelectorAll('[data-node-id]')) {
    const id = el.getAttribute('data-node-id') ?? '';
    // A layout or master shape can have the node id of a slide shape (`slideShape`).
    if (el.getAttribute('data-layer') !== 'slide') continue;
    if (ids.has(id) && !out.has(id)) out.set(id, el);
  }
  return out;
}
