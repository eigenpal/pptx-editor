// Line dashes and line ends: the geometry that both the SVG paint (`output/paint.ts`) and the PDF
// export (`@pptx-editor.dev/pptx-to-pdf`) draw, so the editor and the PDF agree. DOM-free.
//
// ECMA-376 Part 1 names the dashes (20.1.10.48 ST_PresetLineDashVal, 20.1.8.21 custDash) and the
// line ends (20.1.10.33 ST_LineEndType, 20.1.10.34 ST_LineEndWidth, 20.1.10.35
// ST_LineEndLength), but not their sizes. The sizes below are measured on the reference PDFs of
// the `line-ends-dashes` and `lines-arrows` fixtures:
//
// - A line end scales with the line width, but never with less than 2 pt: a thinner line gets
//   the heads of a 2 pt line. Its length and width are 2, 3, or 5 such units (sm, med, lg).
// - A triangle's tip and a stealth arrow's tip lie on the end point. The line stops half a unit
//   inside a triangle's base, and at a stealth arrow's notch (1, 2, or 3 units from the tip).
// - An open arrow is a stroke as wide as the line, with round caps and a mitered tip on the end
//   point. Its arms reach 2, 3, or 4.5 units back and 1.25, 1.75, or 2.5 units to each side.
//   The line stops at its vertex.
// - An oval and a diamond are centered on the end point, and the line runs to the end point.
// - On a curve, a head points along the chord from the point where the line stops.
// - A dash pattern is in line widths. With round caps, each dash is one width shorter and each
//   gap one width longer, so the caps keep the rhythm. With square caps, a preset dash keeps its
//   flat length, and only the ends of the line get the half-width square cap: the start, and the
//   end when it falls in a dash. A custom dash with square caps gets a cap on every dash.

import type { PathCommand } from '@eigenpal/ooxml-core';
import type { LineEnd, ResolvedLine } from '../cascade/index.ts';

/** Preset dash patterns, in multiples of the line width (ECMA-376 ST_PresetLineDashVal). */
const PRESET_DASHES: Readonly<Record<string, readonly number[]>> = Object.freeze({
  dot: [1, 3],
  dash: [4, 3],
  lgDash: [8, 3],
  dashDot: [4, 3, 1, 3],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDash: [3, 1],
  sysDot: [1, 1],
  sysDashDot: [3, 1, 1, 1],
  sysDashDotDot: [3, 1, 1, 1, 1, 1],
});

/** A zero-width outline draws as a hairline of one pixel at 96 dpi, in EMU. */
const HAIRLINE_WIDTH = 9525;

/** The smallest unit a line end scales with: 2 pt, in EMU. */
const LINE_END_MIN_UNIT = 25_400;

/**
 * The drawn width of an outline in the local units of a shape whose group scales by `scale`.
 * @internal
 */
export function strokeWidth(line: ResolvedLine, scale = 1): number {
  return (line.width > 0 ? line.width : HAIRLINE_WIDTH) / (scale || 1);
}

/**
 * The dash pattern of an outline in multiples of its width, or undefined for a solid line. A
 * `prstDash/@val` comes from the file: an own-key check keeps prototype names out.
 */
function dashMultiples(line: ResolvedLine): readonly number[] | undefined {
  const pattern =
    line.dash === 'custom'
      ? line.dashes
      : Object.hasOwn(PRESET_DASHES, line.dash)
        ? PRESET_DASHES[line.dash]
        : undefined;
  // A pattern with no length would draw nothing, or never end.
  if (!pattern?.length || !pattern.every((v) => Number.isFinite(v) && v >= 0)) return undefined;
  return pattern.some((v) => v > 0) ? pattern : undefined;
}

/**
 * The dash array of an outline of `width`, in the same units, or undefined for a solid line.
 * With round caps each dash is one width shorter and each gap one width longer.
 * @internal
 */
export function dashArray(line: ResolvedLine, width: number): number[] | undefined {
  const pattern = dashMultiples(line);
  if (!pattern) return undefined;
  const round = line.cap === 'rnd';
  return pattern.map((v, i) => {
    const len = v * width + (round ? (i % 2 === 0 ? -width : width) : 0);
    // A dash of zero length still draws its round cap: a dot.
    return Math.max(i % 2 === 0 ? 0.001 : 0, len);
  });
}

/** Whether the square caps of a dashed outline show only at its ends (`planStroke` adds them). */
function capsAtEndsOnly(line: ResolvedLine): boolean {
  return line.cap === 'sq' && line.dash !== 'custom' && dashMultiples(line) !== undefined;
}

/**
 * The cap that each dash and each end of an outline gets. A preset dash with square caps draws
 * butt dashes, and `planStroke` adds the square caps at the line's ends.
 * @internal
 */
export function strokeCap(line: ResolvedLine): 'butt' | 'round' | 'square' {
  if (line.cap === 'rnd') return 'round';
  if (line.cap === 'sq') return capsAtEndsOnly(line) ? 'butt' : 'square';
  return 'butt';
}

/**
 * The dash array of a picture's border, which lies outside the picture: a stroke twice as wide,
 * of which only the outer half shows. A round or square cap of the wider stroke would reach
 * twice as far along the path, so its dashes get butt caps, and each dash grows by the caps of
 * a centered line. Undefined for a solid border.
 * @internal
 */
export function outsideDashArray(line: ResolvedLine, width: number): number[] | undefined {
  const dashes = dashArray(line, width);
  if (!dashes || strokeCap(line) === 'butt') return dashes;
  return dashes.map((v, i) => Math.max(0, v + (i % 2 === 0 ? width : -width)));
}

/** A line end's length and width in units (`sm`, `med`, `lg`). */
const END_SIZE = { sm: 2, med: 3, lg: 5 } as const;
/** How far an open arrow's arms reach back along the line, in units (`len`). */
const ARROW_LENGTH = { sm: 2, med: 3, lg: 4.5 } as const;
/** How far an open arrow's arms reach to each side of the line, in units (`w`). */
const ARROW_REACH = { sm: 1.25, med: 1.75, lg: 2.5 } as const;
/** How far a stealth arrow's notch lies from its tip, in units. */
const STEALTH_NOTCH = { sm: 1, med: 2, lg: 3 } as const;

/** The line end types that draw. */
const END_TYPES = new Set(['triangle', 'stealth', 'diamond', 'oval', 'arrow']);

/** The unit that the line ends of an outline scale with, in the shape's local units. */
function lineEndUnit(line: ResolvedLine, scale = 1): number {
  return Math.max(line.width, LINE_END_MIN_UNIT) / (scale || 1);
}

/**
 * How far from its end point, in any direction, a line end of this outline can paint, in the
 * shape's local units: the farthest point of the head or the tail, with an arrow's stroke. A
 * head on a line shorter than itself reaches past the other end of the line, so this bounds the
 * ends of every line, short or long.
 * @internal
 */
export function lineEndReach(line: ResolvedLine, scale = 1): number {
  const unit = lineEndUnit(line, scale);
  const width = strokeWidth(line, scale);
  let reach = 0;
  for (const end of [line.head, line.tail]) {
    if (!drawsEnd(end)) continue;
    const shape = endShape(end, { x: 0, y: 0 }, { x: 1, y: 0 }, unit, width);
    const pad = shape.kind === 'stroke' ? shape.width / 2 : 0;
    // A curve lies inside its control points, so they bound it too.
    for (const c of shape.commands) {
      if (c.op === 'Z') continue;
      const points = c.op === 'C' ? [c.x1, c.y1, c.x2, c.y2, c.x, c.y] : [c.x, c.y];
      for (let i = 0; i < points.length; i += 2)
        reach = Math.max(reach, Math.hypot(points[i], points[i + 1]) + pad);
    }
  }
  return reach;
}

function drawsEnd(end: LineEnd | undefined): end is LineEnd {
  return end !== undefined && END_TYPES.has(end.type);
}

/**
 * The miter limit of an open arrow's tip: its sharpest tip (sm width, lg length) needs 3.74.
 * @internal
 */
export const ARROW_MITER_LIMIT = 10;

/**
 * A drawn line end: an outline to fill, or a path to stroke with round caps and a mitered tip.
 * @internal
 */
export type LineEndShape =
  | { readonly kind: 'fill'; readonly commands: readonly PathCommand[] }
  | { readonly kind: 'stroke'; readonly commands: readonly PathCommand[]; readonly width: number };

/**
 * The parts of a stroked path: the line itself, its line ends, and the square caps it needs.
 * @internal
 */
export interface StrokePlan {
  /** The path to stroke, shortened where a head covers its end. */
  readonly commands: readonly PathCommand[];
  /** Line ends to draw over the line, in the line's paint. */
  readonly ends: readonly LineEndShape[];
  /** Short solid paths to stroke with butt caps: the square caps of a dashed line. */
  readonly caps: readonly (readonly PathCommand[])[];
}

/**
 * Whether a path is open: it has no `close`. Line ends (`a:headEnd`, `a:tailEnd`) show only on
 * open paths, so a rectangle or a callout box never gets an arrowhead.
 * @internal
 */
export function isOpenPath(commands: readonly PathCommand[]): boolean {
  return commands.length >= 2 && !commands.some((c) => c.op === 'Z');
}

interface Pt {
  readonly x: number;
  readonly y: number;
}

/** One drawing segment with its start point. */
type Seg =
  | { readonly op: 'L'; readonly from: Pt; readonly to: Pt }
  | { readonly op: 'Q'; readonly from: Pt; readonly c1: Pt; readonly to: Pt }
  | { readonly op: 'C'; readonly from: Pt; readonly c1: Pt; readonly c2: Pt; readonly to: Pt };

interface Subpath {
  readonly start: Pt;
  readonly segs: Seg[];
}

const sub = (a: Pt, b: Pt): Pt => ({ x: a.x - b.x, y: a.y - b.y });
const len = (v: Pt): number => Math.hypot(v.x, v.y);
const lerp = (a: Pt, b: Pt, t: number): Pt => ({
  x: a.x + (b.x - a.x) * t,
  y: a.y + (b.y - a.y) * t,
});

/** Split an open path into subpaths. A path that does not start with a move starts at 0, 0. */
function subpathsOf(commands: readonly PathCommand[]): Subpath[] {
  const out: Subpath[] = [];
  let current: Subpath | undefined;
  let at: Pt = { x: 0, y: 0 };
  for (const c of commands) {
    if (c.op === 'Z') continue;
    if (c.op === 'M') {
      current = { start: { x: c.x, y: c.y }, segs: [] };
      out.push(current);
      at = current.start;
      continue;
    }
    if (!current) {
      current = { start: at, segs: [] };
      out.push(current);
    }
    const to = { x: c.x, y: c.y };
    if (c.op === 'L') current.segs.push({ op: 'L', from: at, to });
    else if (c.op === 'Q') current.segs.push({ op: 'Q', from: at, c1: { x: c.x1, y: c.y1 }, to });
    else if (c.op === 'C')
      current.segs.push({
        op: 'C',
        from: at,
        c1: { x: c.x1, y: c.y1 },
        c2: { x: c.x2, y: c.y2 },
        to,
      });
    at = to;
  }
  return out;
}

function commandsOf(subpaths: readonly Subpath[]): PathCommand[] {
  const out: PathCommand[] = [];
  for (const s of subpaths) {
    out.push({ op: 'M', x: s.start.x, y: s.start.y });
    for (const g of s.segs) {
      if (g.op === 'L') out.push({ op: 'L', x: g.to.x, y: g.to.y });
      else if (g.op === 'Q') out.push({ op: 'Q', x1: g.c1.x, y1: g.c1.y, x: g.to.x, y: g.to.y });
      else
        out.push({
          op: 'C',
          x1: g.c1.x,
          y1: g.c1.y,
          x2: g.c2.x,
          y2: g.c2.y,
          x: g.to.x,
          y: g.to.y,
        });
    }
  }
  return out;
}

function reverseSeg(g: Seg): Seg {
  if (g.op === 'L') return { op: 'L', from: g.to, to: g.from };
  if (g.op === 'Q') return { op: 'Q', from: g.to, c1: g.c1, to: g.from };
  return { op: 'C', from: g.to, c1: g.c2, c2: g.c1, to: g.from };
}

function pointAt(g: Seg, t: number): Pt {
  if (g.op === 'L') return lerp(g.from, g.to, t);
  if (g.op === 'Q') return lerp(lerp(g.from, g.c1, t), lerp(g.c1, g.to, t), t);
  const a = lerp(g.from, g.c1, t);
  const b = lerp(g.c1, g.c2, t);
  const c = lerp(g.c2, g.to, t);
  return lerp(lerp(a, b, t), lerp(b, c, t), t);
}

/** The part of a segment from its start to `t` (de Casteljau). */
function headOf(g: Seg, t: number): Seg {
  if (g.op === 'L') return { op: 'L', from: g.from, to: lerp(g.from, g.to, t) };
  if (g.op === 'Q') {
    const c1 = lerp(g.from, g.c1, t);
    return { op: 'Q', from: g.from, c1, to: pointAt(g, t) };
  }
  const a = lerp(g.from, g.c1, t);
  const b = lerp(g.c1, g.c2, t);
  const ab = lerp(a, b, t);
  return { op: 'C', from: g.from, c1: a, c2: ab, to: pointAt(g, t) };
}

/** The direction in which a segment arrives at its end point, or undefined when it has none. */
function arrival(g: Seg): Pt | undefined {
  const candidates = g.op === 'L' ? [g.from] : g.op === 'Q' ? [g.c1, g.from] : [g.c2, g.c1, g.from];
  for (const c of candidates) {
    const d = sub(g.to, c);
    const l = len(d);
    if (l > 1e-9) return { x: d.x / l, y: d.y / l };
  }
  return undefined;
}

/**
 * Cut the segments back from their end point `end` to the first point that lies `by` from it.
 * Returns the kept segments and that point, or undefined when the whole path lies within `by`.
 */
function cutBack(segs: readonly Seg[], by: number): { segs: Seg[]; at: Pt } | undefined {
  const end = segs[segs.length - 1].to;
  for (let i = segs.length - 1; i >= 0; i -= 1) {
    const g = segs[i];
    if (len(sub(g.from, end)) < by) continue;
    // The segment leaves the circle of radius `by` around the end: find where, by bisection.
    let lo = 0;
    let hi = 1;
    for (let k = 0; k < 40; k += 1) {
      const mid = (lo + hi) / 2;
      if (len(sub(pointAt(g, mid), end)) >= by) lo = mid;
      else hi = mid;
    }
    const kept = headOf(g, lo);
    return { segs: [...segs.slice(0, i), kept], at: kept.to };
  }
  return undefined;
}

/** A line end at `tip`, pointing along the unit vector `dir`, for a line of `width`. */
function endShape(end: LineEnd, tip: Pt, dir: Pt, unit: number, width: number): LineEndShape {
  const length = END_SIZE[end.len] * unit;
  const half = (END_SIZE[end.w] * unit) / 2;
  // `back` along the line from the tip, `side` across it.
  const at = (back: number, side: number): Pt => ({
    x: tip.x - dir.x * back - dir.y * side,
    y: tip.y - dir.y * back + dir.x * side,
  });
  const poly = (...pts: Pt[]): PathCommand[] => [
    ...pts.map((p, i): PathCommand => ({ op: i === 0 ? 'M' : 'L', x: p.x, y: p.y })),
    { op: 'Z' },
  ];
  switch (end.type) {
    case 'triangle':
      return { kind: 'fill', commands: poly(at(0, 0), at(length, half), at(length, -half)) };
    case 'stealth':
      return {
        kind: 'fill',
        commands: poly(
          at(0, 0),
          at(length, half),
          at(STEALTH_NOTCH[end.len] * unit, 0),
          at(length, -half)
        ),
      };
    case 'diamond':
      return {
        kind: 'fill',
        commands: poly(at(-length / 2, 0), at(0, half), at(length / 2, 0), at(0, -half)),
      };
    case 'oval': {
      // Four quarter arcs; 0.5523 places the controls of a cubic quarter circle.
      const a = length / 2;
      const b = half;
      const k = 0.5522847498;
      const p = (u: number, v: number): Pt => at(-u, v);
      const c = (c1: Pt, c2: Pt, to: Pt): PathCommand => ({
        op: 'C',
        x1: c1.x,
        y1: c1.y,
        x2: c2.x,
        y2: c2.y,
        x: to.x,
        y: to.y,
      });
      const start = p(a, 0);
      return {
        kind: 'fill',
        commands: [
          { op: 'M', x: start.x, y: start.y },
          c(p(a, k * b), p(k * a, b), p(0, b)),
          c(p(-k * a, b), p(-a, k * b), p(-a, 0)),
          c(p(-a, -k * b), p(-k * a, -b), p(0, -b)),
          c(p(k * a, -b), p(a, -k * b), start),
          { op: 'Z' },
        ],
      };
    }
    default: {
      // `arrow`: a stroke from one arm's end over the vertex to the other's.
      const reach = ARROW_REACH[end.w] * unit;
      const arm = ARROW_LENGTH[end.len] * unit;
      const vertex = arrowVertex(end, unit, width);
      const pts = [at(vertex + arm, reach), at(vertex, 0), at(vertex + arm, -reach)];
      return {
        kind: 'stroke',
        width,
        commands: pts.map((q, i): PathCommand => ({ op: i === 0 ? 'M' : 'L', x: q.x, y: q.y })),
      };
    }
  }
}

/** How far an open arrow's vertex lies behind its mitered tip, for a line of `width`. */
function arrowVertex(end: LineEnd, unit: number, width: number): number {
  const angle = Math.atan2(ARROW_REACH[end.w] * unit, ARROW_LENGTH[end.len] * unit);
  return width / 2 / Math.sin(angle);
}

/** How far the line stops short of its end point under a line end. */
function inset(end: LineEnd, unit: number, width: number): number {
  if (end.type === 'triangle') return END_SIZE[end.len] * unit - unit / 2;
  if (end.type === 'stealth') return STEALTH_NOTCH[end.len] * unit;
  if (end.type === 'arrow') return arrowVertex(end, unit, width);
  return 0;
}

/**
 * Cut the last subpath back for a line end, and return the end's shape. A head points along
 * the chord from the point where the line stops, or along the path's tangent with no inset.
 */
function placeEnd(
  path: Subpath,
  end: LineEnd,
  unit: number,
  width: number
): { path: Subpath; shape?: LineEndShape } {
  if (!path.segs.length) return { path };
  const tip = path.segs[path.segs.length - 1].to;
  const by = inset(end, unit, width);
  const cut = by > 0 ? cutBack(path.segs, by) : undefined;
  let dir: Pt | undefined;
  if (cut) {
    const d = sub(tip, cut.at);
    dir = { x: d.x / len(d), y: d.y / len(d) };
  } else {
    for (let i = path.segs.length - 1; i >= 0 && !dir; i -= 1) dir = arrival(path.segs[i]);
  }
  if (!dir) return { path };
  // A line shorter than its inset lies wholly under the end, so none of it draws.
  const kept = cut
    ? { start: path.start, segs: cut.segs }
    : by > 0
      ? { start: tip, segs: [] }
      : path;
  return { path: kept, shape: endShape(end, tip, dir, unit, width) };
}

function reversed(path: Subpath): Subpath {
  const segs = path.segs.map(reverseSeg).reverse();
  return { start: segs.length ? segs[0].from : path.start, segs };
}

/** The length of a subpath, with curves measured over 32 chords each. */
function lengthOf(path: Subpath): number {
  let total = 0;
  for (const g of path.segs) {
    if (g.op === 'L') {
      total += len(sub(g.to, g.from));
      continue;
    }
    let prev = g.from;
    for (let i = 1; i <= 32; i += 1) {
      const p = pointAt(g, i / 32);
      total += len(sub(p, prev));
      prev = p;
    }
  }
  return total;
}

/** Whether a dash pattern of `pattern` draws at `distance` along the line. */
function inDash(pattern: readonly number[], distance: number): boolean {
  const period = pattern.reduce((s, v) => s + v, 0);
  if (!(period > 0)) return true;
  let r = distance % period;
  for (let i = 0; i < pattern.length; i += 1) {
    if (r < pattern[i]) return i % 2 === 0;
    r -= pattern[i];
  }
  return false;
}

/** A square cap: half a width past `point` in the direction `dir`. */
function capPath(point: Pt, dir: Pt, width: number): PathCommand[] {
  return [
    { op: 'M', x: point.x, y: point.y },
    { op: 'L', x: point.x + (dir.x * width) / 2, y: point.y + (dir.y * width) / 2 },
  ];
}

/**
 * Plan the stroke of one path of a shape whose group scales by `scale`: shorten an open path
 * under its line ends, shape the ends, and add the square caps of a dashed line. A closed path
 * keeps its commands and gets neither.
 * @internal
 */
export function planStroke(
  commands: readonly PathCommand[],
  line: ResolvedLine,
  scale = 1
): StrokePlan {
  const head = drawsEnd(line.head) ? line.head : undefined;
  const tail = drawsEnd(line.tail) ? line.tail : undefined;
  const plain = { commands, ends: [], caps: [] };
  if (!isOpenPath(commands) || (!head && !tail && !capsAtEndsOnly(line))) return plain;
  const width = strokeWidth(line, scale);
  const unit = lineEndUnit(line, scale);
  // A lone move draws nothing, so the ends go on the first and last subpaths that draw.
  const subpaths = subpathsOf(commands).filter((s) => s.segs.length);
  if (!subpaths.length) return plain;
  const ends: LineEndShape[] = [];
  let first = subpaths[0];
  const lastIndex = subpaths.length - 1;
  let changed = false;
  if (head) {
    const back = reversed(first);
    const placed = placeEnd(back, head, unit, width);
    if (placed.path !== back) {
      first = reversed(placed.path);
      changed = true;
    }
    if (placed.shape) ends.push(placed.shape);
  }
  const parts = [first, ...subpaths.slice(1)];
  if (tail) {
    // A head can take all of a short line; the tail then still points along the line itself.
    const emptied = !parts[lastIndex].segs.length;
    const placed = placeEnd(emptied ? subpaths[lastIndex] : parts[lastIndex], tail, unit, width);
    if (!emptied) {
      changed ||= placed.path !== parts[lastIndex];
      parts[lastIndex] = placed.path;
    }
    if (placed.shape) ends.push(placed.shape);
  }
  const caps: PathCommand[][] = [];
  const pattern = dashArray(line, width);
  if (pattern && capsAtEndsOnly(line)) {
    // The dash pattern starts again with a dash on each subpath, so each start gets its cap.
    for (const part of parts) {
      if (!part.segs.length) continue;
      const out = arrival(reverseSeg(part.segs[0]));
      if (out) caps.push(capPath(part.start, out, width));
      const lastSeg = part.segs[part.segs.length - 1];
      const dir = arrival(lastSeg);
      if (dir && inDash(pattern, lengthOf(part))) caps.push(capPath(lastSeg.to, dir, width));
    }
  }
  return { commands: changed ? commandsOf(parts) : commands, ends, caps };
}
