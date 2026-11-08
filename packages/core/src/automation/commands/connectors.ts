// Connectors glued to shapes (ECMA-376 Part 1, 19.3.1.13 `p:cxnSp`, 20.1.2.2.4
// `CT_Connection`, and 20.1.9.10 `a:cxnLst`).
//
// A connector end glues to a connection site of a shape: `a:stCxn` or `a:endCxn` names the
// shape's drawing id and the index of the site in its geometry's `a:cxnLst`. A geometry with
// no `a:cxnLst` has the four sites of `rect`: top, left, bottom, right. Each site has an angle,
// the direction a connector leaves the shape in.
//
// When a glued shape moves or resizes, its connectors follow it: a straight connector puts its
// end on the site, and an elbow (`bentConnector2` to `5`) or a curve (`curvedConnector2` to
// `5`) routes again. The route is a path of horizontal and vertical segments that leaves the
// start site in its direction and enters the end site against its direction; its segment
// count picks the preset, and its corners give the adjust values. The editor's drag preview
// and the command that writes the drag share these functions, so the preview is the result.
//
// DOM-free: the editor's surface and the commands both call it.

import {
  apply,
  boxMatrix,
  connectionSites,
  evaluateGeometry,
  invert,
  multiply,
  type LayoutItem,
  type Matrix,
} from '../../layout/index.ts';
import type { GeometrySpec, Xfrm } from '../../cascade/index.ts';
import type { ConnectionTarget, ConnectorRoute, TransformItem } from '../../contracts/neutral.ts';
import {
  deltaToParent,
  lineEnds,
  lineTransform,
  toFile,
  toParent,
  unionBounds,
  type Box,
} from './placement.ts';

type Point = { readonly x: number; readonly y: number };

/** A connection site on the slide: where it is, and the unit direction a connector leaves in. */
export interface PlacedSite extends ConnectionTarget {
  readonly x: number;
  readonly y: number;
  readonly dx: number;
  readonly dy: number;
}

/** One end of a route: a slide point, with the direction it leaves in when it is glued. */
export interface RouteEnd {
  readonly p: Point;
  readonly dir?: Point;
}

const RECT: GeometrySpec = { kind: 'preset', preset: 'rect', adjust: new Map() };
/** How far an elbow runs straight out of a site before it turns: 1/4 in. */
const STUB = 228_600;
/** Guide values are 32-bit integers (`ST_GeomGuideFormula` literals). */
const GUIDE_LIMIT = 2_147_483_647;

const STRAIGHT = /^(line|straightConnector1)$/;
const ROUTED = /^(bentConnector|curvedConnector)[2-5]$/;
/** The presets a connector route writes. */
export const CONNECTOR_PRESET = /^(line|straightConnector1|(bentConnector|curvedConnector)[2-5])$/;

/** Whether an item is a connector (`p:cxnSp`): only a connector's ends glue. */
export function isConnector(item: LayoutItem): boolean {
  return item.kind === 'shape' && item.element.kind === 'connector';
}

/** Whether a connector ends glue to `item`: a slide shape, picture, or frame, never a group. */
function takesGlue(item: LayoutItem): boolean {
  if (item.kind === 'group' || isConnector(item)) return false;
  return item.element.layer === 'slide' && item.element.drawingId !== undefined;
}

/** Every item of the tree, groups and their children included. */
export function flatItems(items: readonly LayoutItem[]): LayoutItem[] {
  const out: LayoutItem[] = [];
  const walk = (list: readonly LayoutItem[]): void => {
    for (const item of list) {
      out.push(item);
      if (item.kind === 'group') walk(item.children);
    }
  };
  walk(items);
  return out;
}

/** A file transform of `item` as a box in its parent space, flips included. */
export function parentXfrm(item: LayoutItem, t: TransformItem): Xfrm {
  const s = item.toSource;
  const el = item.element.xfrm;
  const box = s
    ? { x: (t.x - s.ox) * s.sx, y: (t.y - s.oy) * s.sy, cx: t.cx * s.sx, cy: t.cy * s.sy }
    : { x: t.x, y: t.y, cx: t.cx, cy: t.cy };
  return { ...box, rot: t.rot, flipH: t.flipH ?? el.flipH, flipV: t.flipV ?? el.flipV };
}

/**
 * The connection sites of `item` on the slide, in `a:cxnLst` order, at its laid-out box or at
 * `box` (a box in its parent space). Items that take no glue have none.
 */
export function placedSites(item: LayoutItem, box?: Xfrm): PlacedSite[] {
  if (!takesGlue(item)) return [];
  const xf = box ?? item.element.xfrm;
  const m: Matrix = box ? multiply(item.parentMatrix, boxMatrix(box)) : item.matrix;
  const spec = item.kind === 'shape' || item.kind === 'picture' ? item.element.geometry : RECT;
  let sites = connectionSites(spec, xf.cx, xf.cy);
  if (sites.length === 0) sites = connectionSites(RECT, xf.cx, xf.cy);
  return sites.map((s, index) => {
    const at = apply(m, s.x, s.y);
    const r = (s.angle * Math.PI) / 180;
    const ux = Math.cos(r);
    const uy = Math.sin(r);
    const dx = m[0] * ux + m[2] * uy;
    const dy = m[1] * ux + m[3] * uy;
    const len = Math.hypot(dx, dy) || 1;
    return { nodeId: item.element.nodeId, index, x: at.x, y: at.y, dx: dx / len, dy: dy / len };
  });
}

/**
 * The connection sites near a slide point: the sites of each shape that `p` is over or within
 * `reach` of, and the nearest of them within `reach` (the site an end glues to), or null.
 * `exclude` names items that are not targets (the connector itself).
 */
export function sitesNear(
  items: readonly LayoutItem[],
  p: Point,
  reach: number,
  exclude: ReadonlySet<string>
): { readonly shown: PlacedSite[]; readonly snap: PlacedSite | null } {
  const shown: PlacedSite[] = [];
  let snap: PlacedSite | null = null;
  let best = reach;
  for (const item of flatItems(items)) {
    if (exclude.has(item.element.nodeId) || item.element.hidden || !takesGlue(item)) continue;
    const b = unionBounds([item]);
    if (p.x < b.x - reach || p.x > b.r + reach || p.y < b.y - reach || p.y > b.b + reach) continue;
    for (const s of placedSites(item)) {
      shown.push(s);
      const d = Math.hypot(s.x - p.x, s.y - p.y);
      if (d <= best) {
        best = d;
        snap = s;
      }
    }
  }
  return { shown, snap };
}

/** The items by drawing id (`p:cNvPr/@id`), for the slide's shapes that take glue. */
export function byDrawingId(items: readonly LayoutItem[]): Map<string, LayoutItem> {
  const out = new Map<string, LayoutItem>();
  for (const item of flatItems(items))
    if (takesGlue(item) && !out.has(item.element.drawingId!))
      out.set(item.element.drawingId!, item);
  return out;
}

/** The site a connector end glues to now, or undefined when the end is free. */
export function gluedSite(
  connector: LayoutItem,
  which: 'start' | 'end',
  targets: ReadonlyMap<string, LayoutItem>,
  box?: (item: LayoutItem) => Xfrm | undefined
): PlacedSite | undefined {
  if (connector.kind !== 'shape') return undefined;
  const ref = connector.element.connections?.[which];
  const target = ref ? targets.get(ref.id) : undefined;
  if (!ref || !target) return undefined;
  return placedSites(target, box?.(target))[ref.idx];
}

/** The axis direction nearest to a direction. */
function axisOf(d: Point): Point {
  return Math.abs(d.x) >= Math.abs(d.y)
    ? { x: d.x < 0 ? -1 : 1, y: 0 }
    : { x: 0, y: d.y < 0 ? -1 : 1 };
}

const AXES: readonly Point[] = [
  { x: 1, y: 0 },
  { x: -1, y: 0 },
  { x: 0, y: 1 },
  { x: 0, y: -1 },
];

/**
 * The corners of an elbow route from `s`, leaving in direction `ds`, to `e`, entered against
 * direction `de` (both axis directions, in one space). `m` is the length a route runs out of a
 * site before it turns back. Two to five segments.
 */
export function elbowPoints(s: Point, ds: Point, e: Point, de: Point, m = STUB): Point[] {
  const ahead = (from: Point, dir: Point, to: Point): number =>
    (to.x - from.x) * dir.x + (to.y - from.y) * dir.y;
  const coord = (dir: Point, p: Point): number => (dir.x !== 0 ? p.x : p.y);
  // `q` moved to the coordinate `c` on the axis of `dir`.
  const at = (dir: Point, c: number, q: Point): Point =>
    dir.x !== 0 ? { x: c, y: q.y } : { x: q.x, y: c };
  const step = (p: Point, dir: Point, k: number): Point => ({
    x: p.x + dir.x * k,
    y: p.y + dir.y * k,
  });
  const sign = ds.x + ds.y;
  if ((ds.x !== 0) !== (de.x !== 0)) {
    // Perpendicular: one corner, when it is ahead of both ends.
    const corner = at(ds, coord(ds, e), s);
    if (ahead(s, ds, corner) > 0 && ahead(e, de, corner) > 0) return [s, corner, e];
    // Else out of the start, across beyond the end, and into the end.
    const gap = ahead(s, ds, e);
    const k = gap > 2 * m ? coord(ds, s) + (sign * gap) / 2 : coord(ds, step(s, ds, m));
    const mid = (coord(de, s) + coord(de, e)) / 2;
    const c = ahead(e, de, at(de, mid, e)) >= m ? mid : coord(de, step(e, de, m));
    const a = at(ds, k, s);
    return [s, a, at(de, c, a), at(de, c, e), e];
  }
  if (ds.x === -de.x && ds.y === -de.y) {
    const gap = ahead(s, ds, e);
    // Facing each other: turn half way.
    if (gap > 0) {
      const k = coord(ds, s) + (sign * gap) / 2;
      return [s, at(ds, k, s), at(ds, k, e), e];
    }
    // Back to back: out of each end, and across between them.
    const across: Point = { x: ds.y !== 0 ? 1 : 0, y: ds.x !== 0 ? 1 : 0 };
    const a = step(s, ds, m);
    const d = step(e, de, m);
    let mid = (coord(across, s) + coord(across, e)) / 2;
    if (Math.abs(coord(across, s) - coord(across, e)) < 1) mid += m;
    return [s, a, at(across, mid, a), at(across, mid, d), d, e];
  }
  // The same direction: around both ends.
  const k =
    sign > 0 ? Math.max(coord(ds, s), coord(ds, e)) + m : Math.min(coord(ds, s), coord(ds, e)) - m;
  return [s, at(ds, k, s), at(ds, k, e), e];
}

/** The total length of a path. */
function lengthOf(pts: readonly Point[]): number {
  let n = 0;
  for (let i = 1; i < pts.length; i += 1)
    n += Math.abs(pts[i].x - pts[i - 1].x) + Math.abs(pts[i].y - pts[i - 1].y);
  return n;
}

/** The elbow route with the fewest segments for the directions each free end could take. */
function bestRoute(s: Point, ds: Point | undefined, e: Point, de: Point | undefined): Point[] {
  let best: Point[] | null = null;
  for (const a of ds ? [ds] : AXES)
    for (const b of de ? [de] : AXES) {
      const pts = elbowPoints(s, a, e, b);
      if (
        !best ||
        pts.length < best.length ||
        (pts.length === best.length && lengthOf(pts) < lengthOf(best) - 1)
      )
        best = pts;
    }
  return best!;
}

/**
 * The box, flips, and adjust values that draw the elbow path `pts` (parent space) with the
 * preset of `pts.length - 1` segments. The preset's path leaves its box's top-left corner along
 * the box's x axis: the box turns 90° when the first segment is vertical, and its flips point
 * it at the end.
 */
function fitElbow(
  pts: readonly Point[],
  vertical: boolean
): { box: Box; flipH: boolean; flipV: boolean; adjust: Record<string, number> } {
  const s = pts[0];
  // Into the preset's frame, before flips: a turn of -90° when the first segment is vertical.
  const local = (p: Point): Point => {
    const v = { x: p.x - s.x, y: p.y - s.y };
    return vertical ? { x: v.y, y: -v.x } : v;
  };
  const span = local(pts[pts.length - 1]);
  const flipH = span.x < 0;
  const flipV = span.y < 0;
  const q = pts.map((p) => {
    const l = local(p);
    return { u: flipH ? -l.x : l.x, v: flipV ? -l.y : l.y };
  });
  let w = Math.abs(span.x);
  let h = Math.abs(span.y);
  const n = pts.length - 1;
  // The corners each adjust value places: x of the first corner, y of the second, x of the third.
  const us = n >= 3 ? [q[1].u, ...(n === 5 ? [q[3].u] : [])] : [];
  const vs = n >= 4 ? [q[2].v] : [];
  // A route that turns away from a zero-width box needs some width for its adjust value.
  const need = (vals: readonly number[]): number => {
    let most = 0;
    for (const v of vals) most = Math.max(most, Math.ceil((Math.abs(v) * 100_000) / GUIDE_LIMIT));
    return most;
  };
  w = Math.max(w, need(us));
  h = Math.max(h, need(vs));
  const value = (c: number, size: number): number =>
    size > 0
      ? Math.max(-GUIDE_LIMIT, Math.min(GUIDE_LIMIT, Math.round((c / size) * 100_000)))
      : 50_000;
  const adjust: Record<string, number> = {};
  if (n === 3) adjust.adj1 = value(q[1].u, w);
  if (n >= 4) {
    adjust.adj1 = value(q[1].u, w);
    adjust.adj2 = value(q[2].v, h);
  }
  if (n === 5) adjust.adj3 = value(q[3].u, w);
  // Back to the parent space: the end at (w, h) of the flipped, turned frame.
  const lu = flipH ? -w : w;
  const lv = flipV ? -h : h;
  const end = vertical ? { x: s.x - lv, y: s.y + lu } : { x: s.x + lu, y: s.y + lv };
  const cx = (s.x + end.x) / 2;
  const cy = (s.y + end.y) / 2;
  return {
    box: { x: cx - w / 2, y: cy - h / 2, cx: w, cy: h, rot: vertical ? 90 : 0 },
    flipH,
    flipV,
    adjust,
  };
}

/** The connector family of a preset: straight, elbow, curve, or none (it is not a connector). */
function familyOf(item: LayoutItem): 'straight' | 'bentConnector' | 'curvedConnector' | null {
  if (item.kind !== 'shape' || item.element.geometry.kind !== 'preset') return null;
  const preset = item.element.geometry.preset;
  if (STRAIGHT.test(preset)) return 'straight';
  if (ROUTED.test(preset)) return preset.startsWith('bent') ? 'bentConnector' : 'curvedConnector';
  return null;
}

/**
 * The transform that puts a connector's ends on `start` and `end` (slide points). A straight
 * connector, or one whose ends are both free, stretches to them. An elbow or a curve routes
 * from the start's direction to the end's, and the route sets its preset and adjust values.
 */
export function connectorRoute(item: LayoutItem, start: RouteEnd, end: RouteEnd): TransformItem {
  const s = toParent(item, start.p.x, start.p.y);
  const e = toParent(item, end.p.x, end.p.y);
  const family = familyOf(item);
  if (family === null || family === 'straight' || (!start.dir && !end.dir))
    return lineTransform(item, s, e);
  // Directions into the parent space, as axis directions.
  const dirOf = (from: Point, d: Point | undefined): Point | undefined => {
    if (!d) return undefined;
    const a = toParent(item, from.x, from.y);
    const b = toParent(item, from.x + d.x * 1000, from.y + d.y * 1000);
    return axisOf({ x: b.x - a.x, y: b.y - a.y });
  };
  const pts = bestRoute(s, dirOf(start.p, start.dir), e, dirOf(end.p, end.dir));
  const vertical = pts[1].x === pts[0].x && pts[1].y !== pts[0].y;
  const fit = fitElbow(pts, vertical);
  const route: ConnectorRoute = { preset: `${family}${pts.length - 1}`, adjust: fit.adjust };
  return { ...toFile(item, fit.box), flipH: fit.flipH, flipV: fit.flipV, connector: route };
}

/**
 * The connectors glued to the shapes that `moved` transforms, routed to their sites at the new
 * boxes. A connector in `moved` keeps its own transform.
 */
export function followingConnectors(
  items: readonly LayoutItem[],
  moved: readonly TransformItem[]
): TransformItem[] {
  const byNode = new Map(moved.map((t) => [t.nodeId, t]));
  const flat = flatItems(items);
  const targets = byDrawingId(items);
  const boxOf = (target: LayoutItem): Xfrm | undefined => {
    const t = byNode.get(target.element.nodeId);
    return t ? parentXfrm(target, t) : undefined;
  };
  const out: TransformItem[] = [];
  for (const item of flat) {
    // A layout or master connector never follows: its node id can name another slide shape.
    if (item.element.layer !== 'slide') continue;
    if (!isConnector(item) || item.kind !== 'shape' || byNode.has(item.element.nodeId)) continue;
    const c = item.element.connections;
    if (!c) continue;
    const follows = [c.start, c.end].some((ref) => {
      const target = ref ? targets.get(ref.id) : undefined;
      return !!target && byNode.has(target.element.nodeId);
    });
    if (!follows) continue;
    const ends = lineEnds(item);
    if (!ends) continue;
    const endOf = (which: 'start' | 'end'): RouteEnd => {
      const site = gluedSite(item, which, targets, boxOf);
      return site ? { p: site, dir: { x: site.dx, y: site.dy } } : { p: ends[which] };
    };
    out.push(connectorRoute(item, endOf('start'), endOf('end')));
  }
  return out;
}

/** The most segments a connector preset draws (`bentConnector5`). */
const MOST_SEGMENTS = 5;

/**
 * The corners of an elbow or curve connector in its parent space, from its start to its end,
 * with no segment of zero length, or null for any other item. A curve has the corners of the elbow with the same values: its
 * segments bend between them.
 */
export function elbowCorners(item: LayoutItem): Point[] | null {
  const family = familyOf(item);
  if (item.kind !== 'shape' || item.element.geometry.kind !== 'preset') return null;
  if (family !== 'bentConnector' && family !== 'curvedConnector') return null;
  const { preset, adjust } = item.element.geometry;
  const { cx, cy } = item.element.xfrm;
  const spec: GeometrySpec = { kind: 'preset', preset: `bentConnector${preset.slice(-1)}`, adjust };
  const cmds = evaluateGeometry(spec, cx, cy).paths[0]?.commands ?? [];
  // The item's own box to its parent's space: the item matrix, then out of the parent.
  const own = multiply(invert(item.parentMatrix), item.matrix);
  const pts: Point[] = [];
  for (const c of cmds) if (c.op === 'M' || c.op === 'L') pts.push(apply(own, c.x, c.y));
  // Adjust values can put a corner on an end or on another corner: those segments have no
  // length, so the route is the corners without them. A straight route keeps one segment.
  const route = simplified(pts);
  return route.length >= 2 ? route : null;
}

/** The corners without repeated points and without corners on a straight run. */
function simplified(pts: readonly Point[]): Point[] {
  const near = (a: number, b: number): boolean => Math.abs(a - b) < 1;
  const out: Point[] = [];
  for (const p of pts) {
    const q = out[out.length - 1];
    if (!q || !near(q.x, p.x) || !near(q.y, p.y)) out.push(p);
  }
  for (let i = 1; i < out.length - 1;) {
    const [a, b, c] = [out[i - 1], out[i], out[i + 1]];
    const straight = (near(a.x, b.x) && near(b.x, c.x)) || (near(a.y, b.y) && near(b.y, c.y));
    if (straight) out.splice(i, 1);
    else i += 1;
  }
  return out;
}

/**
 * Whether a drag of segment `index` of `corners` can draw a route a preset has: an end
 * segment adds one bend, or two at a glued end, and no preset draws more than five segments.
 */
export function segmentMovable(
  item: LayoutItem,
  corners: readonly Point[],
  index: number
): boolean {
  const glue = item.kind === 'shape' ? item.element.connections : undefined;
  const last = corners.length - 2;
  const added =
    (index === 0 ? (glue?.start ? 2 : 1) : 0) + (index === last ? (glue?.end ? 2 : 1) : 0);
  return corners.length - 1 + added <= MOST_SEGMENTS;
}

/** The axis a segment of a route moves along: `y` for a horizontal segment, `x` for a vertical. */
export function segmentAxis(a: Point, b: Point): 'x' | 'y' {
  return Math.abs(a.y - b.y) < Math.abs(a.x - b.x) ? 'y' : 'x';
}

/**
 * The transform of an elbow or curve connector after segment `index` of its route moves across
 * as far as the pointer moved from slide point `from` to `to`. A middle segment moves with its two corners. An end segment keeps
 * its end and gains a bend there. A glued end keeps a short run in its site's direction
 * before the bend, so the connector still meets the shape square. Segments that come in line
 * merge, and a route back where it started is `same`. Within `snap` (slide units) of another corner's line or of its own first line, the
 * segment snaps to it, so a bend is easy to take out and a slip is easy to undo. The ends and their glue stay. Null when the route would need more
 * segments than a preset draws, or when the item is no elbow or curve.
 */
export function moveElbowSegment(
  item: LayoutItem,
  index: number,
  from: Point,
  to: Point,
  snap: number | null = null
): TransformItem | 'same' | null {
  const corners = elbowCorners(item);
  const family = familyOf(item);
  if (!corners || index < 0 || index >= corners.length - 1) return null;
  const a = corners[index];
  const b = corners[index + 1];
  const across = segmentAxis(a, b);
  // The segment moves by the pointer's move, so a press beside the handle does not jump it.
  const p0 = toParent(item, from.x, from.y);
  const p1 = toParent(item, to.x, to.y);
  const at = { ...a, [across]: a[across] + p1[across] - p0[across] };
  if (snap !== null) {
    const d = deltaToParent(item, snap, 0);
    const reach = Math.hypot(d.x, d.y);
    const lines = [a, ...corners.filter((_, i) => i !== index && i !== index + 1)];
    const off = (q: Point): number => Math.abs(q[across] - at[across]);
    const nearest = lines.sort((u, v) => off(u) - off(v))[0];
    if (off(nearest) <= reach) at[across] = nearest[across];
  }
  const moved = (q: Point): Point => ({ ...q, [across]: at[across] });
  const last = corners.length - 2;
  const glue = item.kind === 'shape' ? item.element.connections : undefined;
  // The point on the segment a short run from end `q`, toward `to`.
  const stub = (q: Point, to: Point): Point => {
    const len = Math.hypot(to.x - q.x, to.y - q.y);
    const k = len > 0 ? Math.min(STUB, len / 2) / len : 0;
    return { x: q.x + (to.x - q.x) * k, y: q.y + (to.y - q.y) * k };
  };
  // An end segment: the end, its run out of a glued site, then the moved segment.
  const head = (end: Point, other: Point, glued: boolean): Point[] =>
    glued ? [end, stub(end, other), moved(stub(end, other))] : [end, moved(end)];
  const route = simplified([
    ...corners.slice(0, index),
    ...(index === 0 ? head(a, b, !!glue?.start) : [moved(a)]),
    ...(index === last ? head(b, a, !!glue?.end).reverse() : [moved(b)]),
    ...corners.slice(index + 2),
  ]);
  // Back on its first route (a straight elbow back in line too): nothing to write.
  const same =
    route.length === corners.length &&
    route.every((q, i) => Math.abs(q.x - corners[i].x) < 1 && Math.abs(q.y - corners[i].y) < 1);
  if (same || route.length < 3) return 'same';
  const n = route.length - 1;
  if (n > MOST_SEGMENTS) return null;
  const vertical = Math.abs(route[1].x - route[0].x) < 1;
  const fit = fitElbow(route, vertical);
  const connector: ConnectorRoute = { preset: `${family}${n}`, adjust: fit.adjust };
  return { ...toFile(item, fit.box), flipH: fit.flipH, flipV: fit.flipV, connector };
}
