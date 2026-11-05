// Group editing: transforms of several shapes in one transaction, group bounds that follow
// their children, and group / ungroup (design D3–D5 of `group-and-multi-edit`). Values here are
// the file's: a group child's box is in its group's child space (`chOff`/`chExt`).

import { attr, child, elements, NS, type Element } from '../../store/index.ts';
import { apply, boxMatrix, type LayoutItem } from '../../layout/index.ts';
import type { TransformItem } from '../../contracts/neutral.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import {
  coordinate,
  nextShapeId,
  parentOf,
  positiveCoordinate,
  spTreeOf,
} from './package-edits.ts';
import { findItem, toFile, boxOf } from './placement.ts';
import { setTransform, specFromXml, transformHost } from './shape-commands.ts';
import { removeShapes } from './shape-removal.ts';

/** A transform as stored in the file, EMU and degrees. */
export interface FileXfrm {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
  readonly rot: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
}

const SHAPE_KINDS = new Set(['sp', 'pic', 'grpSp', 'cxnSp', 'graphicFrame', 'contentPart']);

const num = (el: Element | undefined, name: string): number | undefined => {
  const v = el ? attr(el, name) : undefined;
  const n = v === undefined ? NaN : Number(v);
  return Number.isFinite(n) ? n : undefined;
};
const flag = (el: Element, name: string): boolean => {
  const v = attr(el, name);
  return v === '1' || v === 'true';
};

/** The `a:xfrm` (or a frame's `p:xfrm`) of a shape, if the file has one. */
function xfrmElement(el: Element): Element | undefined {
  if (el.kind === 'graphicFrame') return child(el, 'p:xfrm');
  const host = transformHost(el);
  return host ? child(host, 'a:xfrm') : undefined;
}

/** The shape's transform in the file; inherited (placeholder) transforms come from layout. */
export function fileXfrm(
  batch: OpBatch,
  part: string,
  id: string,
  items: readonly LayoutItem[]
): FileXfrm {
  const el = batch.element(part, id);
  const xfrm = xfrmElement(el);
  const off = xfrm ? child(xfrm, 'a:off') : undefined;
  const ext = xfrm ? child(xfrm, 'a:ext') : undefined;
  if (xfrm && off && ext)
    return {
      x: num(off, 'x') ?? 0,
      y: num(off, 'y') ?? 0,
      cx: num(ext, 'cx') ?? 0,
      cy: num(ext, 'cy') ?? 0,
      rot: (num(xfrm, 'rot') ?? 0) / 60_000,
      flipH: flag(xfrm, 'flipH'),
      flipV: flag(xfrm, 'flipV'),
    };
  const item = findItem(items, id);
  if (!item) throw new CommandRefusal('disabled.unsupportedElement');
  const box = toFile(item, boxOf(item));
  return { ...box, flipH: item.element.xfrm.flipH, flipV: item.element.xfrm.flipV };
}

/** A group's child space: `chOff`/`chExt`, defaulting to its own box. */
function childSpace(
  group: Element,
  box: FileXfrm
): { x: number; y: number; cx: number; cy: number } {
  const xfrm = xfrmElement(group);
  const chOff = xfrm ? child(xfrm, 'a:chOff') : undefined;
  const chExt = xfrm ? child(xfrm, 'a:chExt') : undefined;
  return {
    x: num(chOff, 'x') ?? box.x,
    y: num(chOff, 'y') ?? box.y,
    cx: num(chExt, 'cx') ?? box.cx,
    cy: num(chExt, 'cy') ?? box.cy,
  };
}

/** Shape children of a shape tree or group, in z-order. */
export function shapesOf(parent: Element): Element[] {
  return parent.children.filter(
    (c): c is Element => c.type === 'element' && c.ns === NS.p && SHAPE_KINDS.has(c.local)
  );
}

/** Whether a shape is a placeholder (`p:nv*Pr/p:nvPr/p:ph`). */
export function isPlaceholder(el: Element): boolean {
  return !!child(child(elements(el)[0], 'p:nvPr'), 'p:ph');
}

/** Axis-aligned bounds of a box after its rotation, in the same space. */
function rotatedBounds(b: FileXfrm): { x: number; y: number; r: number; b: number } {
  const m = boxMatrix({ ...b, flipH: false, flipV: false });
  const pts = [apply(m, 0, 0), apply(m, b.cx, 0), apply(m, b.cx, b.cy), apply(m, 0, b.cy)];
  return {
    x: Math.min(pts[0].x, pts[1].x, pts[2].x, pts[3].x),
    y: Math.min(pts[0].y, pts[1].y, pts[2].y, pts[3].y),
    r: Math.max(pts[0].x, pts[1].x, pts[2].x, pts[3].x),
    b: Math.max(pts[0].y, pts[1].y, pts[2].y, pts[3].y),
  };
}

function union(boxes: readonly FileXfrm[]): { x: number; y: number; cx: number; cy: number } {
  let x = Infinity;
  let y = Infinity;
  let r = -Infinity;
  let b = -Infinity;
  for (const box of boxes) {
    const u = rotatedBounds(box);
    x = Math.min(x, u.x);
    y = Math.min(y, u.y);
    r = Math.max(r, u.r);
    b = Math.max(b, u.b);
  }
  return { x, y, cx: r - x, cy: b - y };
}

function setChildSpace(
  batch: OpBatch,
  part: string,
  group: string,
  space: { x: number; y: number; cx: number; cy: number }
): void {
  const xfrm = xfrmElement(batch.element(part, group));
  if (!xfrm) return;
  const write = (local: string, a: string, av: string, b: string, bv: string): void => {
    const el = child(xfrm, `a:${local}`);
    if (!el) return;
    for (const [name, value] of [
      [a, av],
      [b, bv],
    ] as const)
      batch.push({ op: 'setAttribute', part, node: el.id, name: { ns: '', local: name }, value });
  };
  write('chOff', 'x', coordinate(space.x, 'x'), 'y', coordinate(space.y, 'y'));
  write(
    'chExt',
    'cx',
    positiveCoordinate(Math.max(0, space.cx), 'cx'),
    'cy',
    positiveCoordinate(Math.max(0, space.cy), 'cy')
  );
}

/**
 * Make a group's child space the union of its children, moving and resizing the group's own
 * box so nothing moves on the slide (design D3). Refits enclosing groups outward.
 */
export function refitGroup(
  batch: OpBatch,
  part: string,
  groupId: string,
  items: readonly LayoutItem[]
): void {
  for (let id: string | null = groupId, depth = 0; id && depth < 64; depth += 1) {
    const group = batch.element(part, id);
    if (group.kind !== 'grpSp' || !xfrmElement(group)) return;
    const kids = shapesOf(group);
    if (kids.length === 0) return;
    const box = fileXfrm(batch, part, id, items);
    const space = childSpace(group, box);
    const u = union(kids.map((k) => fileXfrm(batch, part, k.id, items)));
    if (
      Math.abs(u.x - space.x) < 1 &&
      Math.abs(u.y - space.y) < 1 &&
      Math.abs(u.cx - space.cx) < 1 &&
      Math.abs(u.cy - space.cy) < 1
    )
      return;
    const sx = space.cx ? box.cx / space.cx : 1;
    const sy = space.cy ? box.cy / space.cy : 1;
    // The new box's center, through the group's current transform; rotating around the new
    // center then maps every child exactly where it was.
    const local = { x: (u.x - space.x) * sx, y: (u.y - space.y) * sy };
    const size = { cx: u.cx * sx, cy: u.cy * sy };
    const center = apply(boxMatrix(box), local.x + size.cx / 2, local.y + size.cy / 2);
    setTransform(
      batch,
      part,
      id,
      center.x - size.cx / 2,
      center.y - size.cy / 2,
      size.cx,
      size.cy,
      box.rot
    );
    setChildSpace(batch, part, id, u);
    const parent = parentOf(batch, part, id);
    id = parent.kind === 'grpSp' ? parent.id : null;
  }
}

/** Write several transforms, then refit every group whose children moved. */
export function setTransforms(
  batch: OpBatch,
  part: string,
  transforms: readonly TransformItem[],
  items: readonly LayoutItem[]
): void {
  const groups = new Set<string>();
  for (const t of transforms) {
    if (![t.x, t.y, t.cx, t.cy, t.rot].every(Number.isFinite))
      throw new CommandRefusal('errors.unsupported', 'transform');
    const flips =
      t.flipH !== undefined || t.flipV !== undefined
        ? { h: t.flipH ?? false, v: t.flipV ?? false }
        : undefined;
    setTransform(batch, part, t.nodeId, t.x, t.y, t.cx, t.cy, t.rot, flips);
    const parent = parentOf(batch, part, t.nodeId);
    if (parent.kind === 'grpSp') groups.add(parent.id);
  }
  for (const g of groups) refitGroup(batch, part, g, items);
}

/** Group shapes that share a parent; returns the new group's id (design D5). */
export function groupShapes(
  batch: OpBatch,
  part: string,
  ids: readonly string[],
  items: readonly LayoutItem[]
): string {
  if (ids.length < 2) throw new CommandRefusal('disabled.groupNeedsTwo');
  const parent = parentOf(batch, part, ids[0]);
  if (ids.some((id) => parentOf(batch, part, id).id !== parent.id))
    throw new CommandRefusal('disabled.groupScope');
  for (const id of ids)
    if (isPlaceholder(batch.element(part, id)))
      throw new CommandRefusal('disabled.groupPlaceholder');
  const ordered = shapesOf(parent).filter((s) => ids.includes(s.id));
  if (ordered.length !== new Set(ids).size) throw new CommandRefusal('disabled.unsupportedElement');
  // Materialize inherited transforms so the shapes keep their place inside the group.
  const boxes = ordered.map((s) => {
    const box = fileXfrm(batch, part, s.id, items);
    if (!xfrmElement(s)) setTransform(batch, part, s.id, box.x, box.y, box.cx, box.cy, box.rot);
    return box;
  });
  const u = union(boxes);
  const id = nextShapeId(batch, spTreeOf(batch, part));
  const i = (v: number): string => coordinate(v, 'group');
  const size = (v: number): string => positiveCoordinate(v, 'group');
  const at = parent.children.findIndex((c) => c.id === ordered[ordered.length - 1].id) + 1;
  const group = batch.insert({
    op: 'insertNode',
    part,
    parent: parent.id,
    index: at,
    spec: specFromXml(
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="Group ${id - 1}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
        `<p:grpSpPr><a:xfrm><a:off x="${i(u.x)}" y="${i(u.y)}"/><a:ext cx="${size(u.cx)}" cy="${size(u.cy)}"/>` +
        `<a:chOff x="${i(u.x)}" y="${i(u.y)}"/><a:chExt cx="${size(u.cx)}" cy="${size(u.cy)}"/></a:xfrm></p:grpSpPr></p:grpSp>`
    ),
  });
  for (const s of ordered) {
    const g = batch.element(part, group);
    batch.push({ op: 'moveNode', part, node: s.id, parent: group, index: g.children.length });
  }
  if (parent.kind === 'grpSp') refitGroup(batch, part, parent.id, items);
  return group;
}

/** Ungroup groups into their parent, keeping every child's place; returns the children. */
export function ungroupShapes(
  batch: OpBatch,
  part: string,
  ids: readonly string[],
  items: readonly LayoutItem[]
): string[] {
  const out: string[] = [];
  const groups = ids.filter((id) => batch.element(part, id).kind === 'grpSp');
  if (groups.length === 0) throw new CommandRefusal('disabled.notAGroup');
  for (const groupId of groups) {
    const group = batch.element(part, groupId);
    const box = fileXfrm(batch, part, groupId, items);
    const space = childSpace(group, box);
    const sx = space.cx ? box.cx / space.cx : 1;
    const sy = space.cy ? box.cy / space.cy : 1;
    const m = boxMatrix(box);
    const mirrored = box.flipH !== box.flipV;
    const kids = shapesOf(group);
    // Content this editor does not understand (for example `mc:AlternateContent`) would be lost.
    if (
      elements(group).some(
        (e) => !kids.includes(e) && !['nvGrpSpPr', 'grpSpPr', 'extLst'].includes(e.local)
      )
    )
      throw new CommandRefusal('disabled.unsupportedElement');
    const parent = parentOf(batch, part, groupId);
    let at = parent.children.findIndex((c) => c.id === groupId);
    for (const kid of kids) {
      const k = fileXfrm(batch, part, kid.id, items);
      const cx = k.cx * sx;
      const cy = k.cy * sy;
      const center = apply(m, (k.x + k.cx / 2 - space.x) * sx, (k.y + k.cy / 2 - space.y) * sy);
      const rot = mirrored ? box.rot - k.rot : box.rot + k.rot;
      batch.push({ op: 'moveNode', part, node: kid.id, parent: parent.id, index: at });
      at += 1;
      setTransform(batch, part, kid.id, center.x - cx / 2, center.y - cy / 2, cx, cy, rot, {
        h: k.flipH !== box.flipH,
        v: k.flipV !== box.flipV,
      });
      out.push(kid.id);
    }
    // The children keep their ids; only the group's own id leaves the slide, and with it any
    // animation of the group.
    removeShapes(batch, part, [groupId]);
    if (parent.kind === 'grpSp') refitGroup(batch, part, parent.id, items);
  }
  return out;
}
