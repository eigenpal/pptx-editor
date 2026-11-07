// Coordinate conversions for editing: slide space (where the pointer is) → an item's parent
// space (where its laid-out box lives) → the file's coordinates (where `a:xfrm` lives). Top-level
// items have the identity for both steps; group children go through their group's transform
// and its `chOff`/`chExt` scale (design D2 of `group-and-multi-edit`).

import { apply, invert, type LayoutItem, type Matrix } from '../../layout/index.ts';
import type { EditorSelection, TransformItem } from '../../contracts/neutral.ts';

/** A box in an item's parent space, EMU and degrees. */
export interface Box {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
  readonly rot: number;
}

/** A slide point in the item's parent space. */
export function toParent(item: LayoutItem, x: number, y: number): { x: number; y: number } {
  return apply(invert(item.parentMatrix), x, y);
}

/** A slide-space displacement in the item's parent space (linear part only). */
export function deltaToParent(item: LayoutItem, dx: number, dy: number): { x: number; y: number } {
  const m: Matrix = invert(item.parentMatrix);
  return { x: m[0] * dx + m[2] * dy, y: m[1] * dx + m[3] * dy };
}

/** A parent-space box as the file's transform for this item. */
export function toFile(item: LayoutItem, box: Box): TransformItem {
  const s = item.toSource;
  if (!s) return { nodeId: item.element.nodeId, ...box };
  return {
    nodeId: item.element.nodeId,
    x: box.x / s.sx + s.ox,
    y: box.y / s.sy + s.oy,
    cx: box.cx / s.sx,
    cy: box.cy / s.sy,
    rot: box.rot,
  };
}

/** The item's current box in parent space. */
export function boxOf(item: LayoutItem): Box {
  const x = item.element.xfrm;
  return { x: x.x, y: x.y, cx: x.cx, cy: x.cy, rot: x.rot };
}

/** The item's four corners in slide space. */
export function cornersOf(item: LayoutItem): { x: number; y: number }[] {
  const { cx, cy } = item.element.xfrm;
  return [
    apply(item.matrix, 0, 0),
    apply(item.matrix, cx, 0),
    apply(item.matrix, cx, cy),
    apply(item.matrix, 0, cy),
  ];
}

/** The axis-aligned slide-space bounds of several items. */
export function unionBounds(items: readonly LayoutItem[]): {
  x: number;
  y: number;
  r: number;
  b: number;
} {
  let x = Infinity;
  let y = Infinity;
  let r = -Infinity;
  let b = -Infinity;
  for (const item of items)
    for (const p of cornersOf(item)) {
      x = Math.min(x, p.x);
      y = Math.min(y, p.y);
      r = Math.max(r, p.x);
      b = Math.max(b, p.y);
    }
  return { x, y, r, b };
}

/** Find an item anywhere in the tree by node id. */
export function findItem(items: readonly LayoutItem[], id: string): LayoutItem | undefined {
  let byId = itemsById.get(items);
  if (!byId) {
    // One walk per item list: a command on thousands of selected shapes looks each one up.
    // The first item with an id wins, as an in-order search finds it.
    byId = new Map();
    const stack: LayoutItem[] = [...items].reverse();
    while (stack.length > 0) {
      const item = stack.pop()!;
      if (!byId.has(item.element.nodeId)) byId.set(item.element.nodeId, item);
      if (item.kind === 'group')
        for (let i = item.children.length - 1; i >= 0; i -= 1) stack.push(item.children[i]!);
    }
    itemsById.set(items, byId);
  }
  return byId.get(id);
}

/** The items of each laid-out item list by node id (`findItem`). Lists never change. */
const itemsById = new WeakMap<readonly LayoutItem[], Map<string, LayoutItem>>();

/** The items of a selection scope: the slide's top level, or a group's children. */
export function scopeItems(
  items: readonly LayoutItem[],
  group: string | undefined
): readonly LayoutItem[] {
  if (!group) return items;
  const g = findItem(items, group);
  return g && g.kind === 'group' ? g.children : [];
}

/**
 * Move a box so its center moves to `center` (slide space), keeping size and rotation; the
 * result is in the item's parent space.
 */
export function boxAtSlideCenter(
  item: LayoutItem,
  center: { x: number; y: number },
  cx: number,
  cy: number
): Box {
  const c = toParent(item, center.x, center.y);
  return { x: c.x - cx / 2, y: c.y - cy / 2, cx, cy, rot: item.element.xfrm.rot };
}

/** The item's center in slide space. */
export function slideCenter(item: LayoutItem): { x: number; y: number } {
  const { cx, cy } = item.element.xfrm;
  return apply(item.matrix, cx / 2, cy / 2);
}

/**
 * The rotation handle in slide space: `offset` above the middle of the box's top edge in the
 * shape's own frame, flips included. A `flipV` shape (or one turned
 * 180°) shows it under the box on the slide, never inside it.
 */
export function rotationHandle(item: LayoutItem, offset: number): { x: number; y: number } {
  return apply(item.matrix, item.element.xfrm.cx / 2, -offset);
}

/**
 * The rotation after a drag of the rotation handle from `from` to `to` (slide space): the
 * start rotation plus the change of the pointer angle around the box center, in the item's
 * parent space. The grab point does not matter, so a flipped shape never jumps on the first
 * move. `snap` (Shift) rounds to 15° steps. Degrees in [0, 360).
 */
export function rotationAfterDrag(
  item: LayoutItem,
  box: Box,
  from: { x: number; y: number },
  to: { x: number; y: number },
  snap: boolean
): number {
  const cx = box.x + box.cx / 2;
  const cy = box.y + box.cy / 2;
  const a = toParent(item, from.x, from.y);
  const b = toParent(item, to.x, to.y);
  const delta = (Math.atan2(b.y - cy, b.x - cx) - Math.atan2(a.y - cy, a.x - cx)) * (180 / Math.PI);
  let rot = box.rot + delta;
  if (snap) rot = Math.round(rot / 15) * 15;
  return ((rot % 360) + 360) % 360;
}

/** The id of the group that directly contains an item, if any. */
export function parentGroupOf(items: readonly LayoutItem[], id: string): string | undefined {
  for (const item of items) {
    if (item.kind !== 'group') continue;
    if (item.children.some((c) => c.element.nodeId === id)) return item.element.nodeId;
    const hit = parentGroupOf(item.children, id);
    if (hit) return hit;
  }
  return undefined;
}

/** The group scope a selection works in. */
export function selectionScope(
  items: readonly LayoutItem[],
  selection: EditorSelection
): string | undefined {
  if (selection.kind === 'shapes') return selection.group;
  if (selection.kind === 'text' || selection.kind === 'cells')
    return parentGroupOf(items, selection.nodeId);
  return undefined;
}

// Presets whose path runs from the box's top-left to its bottom-right (flips set direction).
const LINE_PRESET = /^(line|straightConnector1|bentConnector[2-5]|curvedConnector[2-5])$/;

/** Whether an item is a line or connector, edited by its two end points. */
export function isLineItem(item: LayoutItem): boolean {
  if (item.kind !== 'shape') return false;
  const el = item.element;
  return (
    el.kind === 'connector' ||
    (el.geometry.kind === 'preset' && LINE_PRESET.test(el.geometry.preset))
  );
}

/** A line's start and end points in slide space (the first and last points of its path). */
export function lineEnds(
  item: LayoutItem
): { start: { x: number; y: number }; end: { x: number; y: number } } | null {
  if (item.kind !== 'shape') return null;
  const cmds = item.geometry.paths[0]?.commands ?? [];
  const first = cmds[0];
  const points = cmds.filter((c) => c.op !== 'Z');
  const last = points[points.length - 1];
  if (!first || first.op !== 'M' || !last) return null;
  return {
    start: apply(item.matrix, first.x, first.y),
    end: apply(item.matrix, last.x, last.y),
  };
}

/**
 * The transform that puts a line's start and end on two parent-space points. The line keeps
 * its rotation frame; its box spans the points and its flips give the direction.
 */
export function lineTransform(
  item: LayoutItem,
  start: { x: number; y: number },
  end: { x: number; y: number }
): TransformItem {
  const rot = item.element.xfrm.rot;
  const r = (rot * Math.PI) / 180;
  const c = Math.cos(r);
  const s = Math.sin(r);
  // Into the line's unrotated frame, where the box is axis-aligned.
  const un = (p: { x: number; y: number }) => ({ x: p.x * c + p.y * s, y: -p.x * s + p.y * c });
  const a = un(start);
  const b = un(end);
  const cx = Math.abs(b.x - a.x);
  const cy = Math.abs(b.y - a.y);
  const mid = { x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 };
  // Back to parent space: the box rotates about its center.
  const center = { x: mid.x * c - mid.y * s, y: mid.x * s + mid.y * c };
  const box: Box = { x: center.x - cx / 2, y: center.y - cy / 2, cx, cy, rot };
  return { ...toFile(item, box), flipH: b.x < a.x, flipV: b.y < a.y };
}
