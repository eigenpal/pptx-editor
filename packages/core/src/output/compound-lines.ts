// Compound lines (design D5 of `fill-and-line-fidelity`). SVG strokes are centered, so a band at
// distances `a`…`b` from the center line is a stroke of width `2b` minus a stroke of width `2a`,
// built in a mask. On a closed path the outer bands are masked to the outside of the path and
// the inner bands are clipped to its inside. A solid line instead draws each band as a stroke of
// the path offset to the band's center, so every band edge is vector geometry, as in the
// reference PDF. On an open path the outer side is the left of the path's direction (measured on
// a `thickThin` line), as on a clockwise closed path. The specification does not give the band
// fractions. They are measured on `fixtures/fills-lines.pdf`, slide 2, which draws every band as
// a filled outline.

import type { PathCommand } from '../layout/index.ts';
import { svg, num, type Defs } from './svg.ts';

/**
 * Line and gap widths, outer side first, as fractions of the line width. Measured on
 * `fixtures/fills-lines.pdf`, slide 2: 6 pt `dbl` lines are 2 + 2 + 2 pt, `thickThin` and
 * `thinThick` 3.6 + 1.2 + 1.2 pt, and `tri` 1 + 1 + 2 + 1 + 1 pt, on closed and open paths alike.
 */
export const COMPOUND_BANDS: Readonly<Record<string, readonly number[]>> = {
  dbl: [1 / 3, 1 / 3, 1 / 3],
  thickThin: [0.6, 0.2, 0.2],
  thinThick: [0.2, 0.2, 0.6],
  tri: [1 / 6, 1 / 6, 1 / 3, 1 / 6, 1 / 6],
};

/**
 * A drawn band as distances from the center line, on one side (`a` < `b`).
 * @internal
 */
export interface Band {
  readonly a: number;
  readonly b: number;
}

/** Whether a compound type draws different bands on its two sides. */
export function isAsymmetric(cmpd: string): boolean {
  return cmpd === 'thickThin' || cmpd === 'thinThick';
}

/**
 * The drawn bands of a compound line of `width`, split by side of the center line: `outer` is
 * the outside of a closed path, or the left of an open path's direction.
 */
export function compoundBands(
  cmpd: string,
  width: number
): { outer: Band[]; inner: Band[] } | null {
  const bands = Object.hasOwn(COMPOUND_BANDS, cmpd) ? COMPOUND_BANDS[cmpd] : undefined;
  if (!bands) return null;
  const outer: Band[] = [];
  const inner: Band[] = [];
  let p = 0;
  bands.forEach((f, i) => {
    const p0 = p;
    p += f;
    if (i % 2 === 1) return; // a gap
    // Offset from the center line: positive on the outer side.
    const o0 = (0.5 - p0) * width;
    const o1 = (0.5 - p) * width;
    if (o0 > 0) outer.push({ a: Math.max(0, o1), b: o0 });
    if (o1 < 0) inner.push({ a: Math.max(0, -o0), b: -o1 });
  });
  return { outer, inner };
}

/** Mask content for bands: widest first, so each black stroke cuts only its own band. */
function bandStrokes(
  doc: Document,
  d: string,
  bands: readonly Band[],
  stroke: Record<string, string | undefined>
): SVGElement[] {
  const out: SVGElement[] = [];
  const common = {
    d,
    fill: 'none',
    'stroke-linejoin': stroke['stroke-linejoin'],
    'stroke-linecap': stroke['stroke-linecap'],
    'stroke-miterlimit': stroke['stroke-miterlimit'],
    'stroke-dasharray': stroke['stroke-dasharray'],
  };
  for (const band of [...bands].sort((x, y) => y.b - x.b)) {
    out.push(svg(doc, 'path', { ...common, stroke: '#ffffff', 'stroke-width': num(2 * band.b) }));
    if (band.a > 0)
      out.push(svg(doc, 'path', { ...common, stroke: '#000000', 'stroke-width': num(2 * band.a) }));
  }
  return out;
}

/** The most points of one path that offset bands flatten; a longer path draws masked bands. */
export const MAX_OFFSET_POINTS = 4096;

/**
 * The most points that all offset bands of one path write, round corners included; past it the
 * path draws masked bands. A file cannot make one outline write megabytes of path data.
 */
export const MAX_OFFSET_BAND_POINTS = 16_384;

/** The segments that flatten one curve, for offset bands. */
const CURVE_STEPS = 16;

/** The largest angle (radians) of one arc segment of a round offset corner. */
const ARC_STEP = Math.PI / 18;

/**
 * A point of a polyline.
 * @internal
 */
export type Point = { readonly x: number; readonly y: number };

/**
 * One subpath as a polyline (curves flattened).
 * @internal
 */
export interface Polyline {
  readonly points: readonly Point[];
  readonly closed: boolean;
}

/** The subpaths of `commands` as polylines, or null past `MAX_OFFSET_POINTS`. */
function polylines(commands: readonly PathCommand[]): Polyline[] | null {
  type Open = { points: Point[]; closed: boolean };
  const out: Open[] = [];
  let line: Open | undefined;
  let count = 0;
  const add = (p: Point): void => {
    if (!line) out.push((line = { points: [], closed: false }));
    const last = line.points[line.points.length - 1];
    if (last && Math.abs(last.x - p.x) < 1e-6 && Math.abs(last.y - p.y) < 1e-6) return;
    line.points.push(p);
    count += 1;
  };
  let start: Point = { x: 0, y: 0 };
  for (const c of commands) {
    if (count > MAX_OFFSET_POINTS) return null;
    const from = line?.points[line.points.length - 1] ?? start;
    switch (c.op) {
      case 'M':
        line = undefined;
        start = { x: c.x, y: c.y };
        add(start);
        break;
      case 'L':
        add({ x: c.x, y: c.y });
        break;
      case 'Q':
      case 'C':
        for (let i = 1; i <= CURVE_STEPS; i += 1) {
          const t = i / CURVE_STEPS;
          const u = 1 - t;
          if (c.op === 'Q')
            add({
              x: u * u * from.x + 2 * u * t * c.x1 + t * t * c.x,
              y: u * u * from.y + 2 * u * t * c.y1 + t * t * c.y,
            });
          else
            add({
              x: u ** 3 * from.x + 3 * u * u * t * c.x1 + 3 * u * t * t * c.x2 + t ** 3 * c.x,
              y: u ** 3 * from.y + 3 * u * u * t * c.y1 + 3 * u * t * t * c.y2 + t ** 3 * c.y,
            });
        }
        break;
      case 'Z':
        if (line) {
          line.closed = true;
          const pts = line.points;
          const [first, last] = [pts[0], pts[pts.length - 1]];
          if (
            pts.length > 1 &&
            Math.abs(first.x - last.x) < 1e-6 &&
            Math.abs(first.y - last.y) < 1e-6
          )
            pts.pop();
        }
        line = undefined;
        break;
    }
  }
  if (count > MAX_OFFSET_POINTS) return null;
  return out.filter((l) => l.points.length > 1);
}

/** The left normal of a segment, in y-down coordinates. */
function leftNormal(a: Point, b: Point): Point {
  const len = Math.hypot(b.x - a.x, b.y - a.y);
  return { x: (b.y - a.y) / len, y: -(b.x - a.x) / len };
}

/** Twice the signed area of the closed polylines: positive when they run clockwise on screen. */
function windingArea(lines: readonly Polyline[]): number {
  let sum = 0;
  for (const { points, closed } of lines) {
    if (!closed) continue;
    for (let i = 0; i < points.length; i += 1) {
      const p = points[i];
      const q = points[(i + 1) % points.length];
      sum += p.x * q.y - q.x * p.y;
    }
  }
  return sum;
}

/**
 * `line` offset by `k` to the left of its direction (to the right for a negative `k`). A corner
 * takes the line's join on its outer side, as the stroke's own edge does, and the meeting point of
 * the two offset segments on its inner side. Null when it takes more than `budget.left` points;
 * otherwise its points are taken from the budget.
 */
function offsetLine(
  line: Polyline,
  k: number,
  join: string,
  miterLimit: number,
  budget: { left: number }
): Polyline | null {
  const { points, closed } = line;
  const n = points.length;
  const out: Point[] = [];
  const at = (p: Point, nv: Point): Point => ({ x: p.x + nv.x * k, y: p.y + nv.y * k });
  for (let i = 0; i < n; i += 1) {
    if (out.length > budget.left) return null;
    const p = points[i];
    const prev = i > 0 ? points[i - 1] : closed ? points[n - 1] : undefined;
    const next = i < n - 1 ? points[i + 1] : closed ? points[0] : undefined;
    const n0 = prev ? leftNormal(prev, p) : undefined;
    const n1 = next ? leftNormal(p, next) : undefined;
    if (!n0 || !n1) {
      out.push(at(p, (n0 ?? n1)!));
      continue;
    }
    const cos = Math.max(-1, Math.min(1, n0.x * n1.x + n0.y * n1.y));
    const angle = Math.acos(cos);
    // A right turn on screen (positive cross product) bulges to the left.
    const cross = n0.x * n1.y - n0.y * n1.x;
    const outer = cross * k > 0;
    const miter = 1 / Math.sqrt((1 + cos) / 2); // the miter length over the offset
    const meet = (): Point => {
      const s = k / (1 + cos);
      return { x: p.x + (n0.x + n1.x) * s, y: p.y + (n0.y + n1.y) * s };
    };
    if (angle < ARC_STEP / 2 || (!outer && 1 + cos > 1e-6 && miter <= 64)) {
      out.push(meet());
      continue;
    }
    if (!outer || join === 'bevel' || (join !== 'round' && miter > miterLimit)) {
      out.push(at(p, n0), at(p, n1));
      continue;
    }
    if (join !== 'round') {
      out.push(meet());
      continue;
    }
    // A round corner: an arc about the corner from one offset normal to the other.
    const a0 = Math.atan2(n0.y, n0.x);
    const steps = Math.ceil(angle / ARC_STEP);
    const turn = Math.sign(cross) * (angle / steps);
    for (let s = 0; s <= steps; s += 1) {
      const a = a0 + turn * s;
      out.push(at(p, { x: Math.cos(a), y: Math.sin(a) }));
    }
  }
  if (out.length > budget.left) return null;
  budget.left -= out.length;
  return { points: out, closed };
}

/** Path data for polylines. */
function polylineData(lines: readonly Polyline[]): string {
  return lines
    .map(
      ({ points, closed }) =>
        points.map((p, i) => `${i ? 'L' : 'M'}${num(p.x)} ${num(p.y)}`).join('') +
        (closed ? 'Z' : '')
    )
    .join('');
}

/**
 * One band of a compound line: polylines to stroke at `width`, with miter joins.
 * @internal
 */
export interface CompoundBandLine {
  readonly width: number;
  readonly lines: readonly Polyline[];
}

/**
 * The bands of a compound line as polylines of the path offset to each band's center: stroked
 * at the band's width with miter joins, every band edge is vector geometry. The outer side is the
 * outside of a closed path (by its winding) and the left of an open path's direction. `join` is
 * the line's join (`round`, `bevel`, or `miter`). Null for a single line or an unknown type, or
 * past `MAX_OFFSET_POINTS` or `MAX_OFFSET_BAND_POINTS`. The SVG paint and `pptx-to-pdf` share it.
 * @internal
 */
export function compoundBandLines(
  commands: readonly PathCommand[],
  cmpd: string,
  width: number,
  join: string,
  miterLimit: number
): CompoundBandLine[] | null {
  if (!Object.hasOwn(COMPOUND_BANDS, cmpd) || !(width > 0)) return null;
  const lines = polylines(commands);
  if (!lines || lines.length === 0) return null;
  const side = windingArea(lines) < 0 ? -1 : 1;
  const out: CompoundBandLine[] = [];
  const budget = { left: MAX_OFFSET_BAND_POINTS };
  const fractions = COMPOUND_BANDS[cmpd];
  let p = 0;
  for (let i = 0; i < fractions.length; i += 1) {
    const p0 = p;
    p += fractions[i];
    if (i % 2 === 1) continue; // a gap
    const o0 = (0.5 - p0) * width;
    const o1 = (0.5 - p) * width;
    const k = (side * (o0 + o1)) / 2;
    const offset: Polyline[] = [];
    for (const line of lines) {
      const moved = offsetLine(line, k, join, miterLimit, budget);
      if (!moved) return null;
      offset.push(moved);
    }
    out.push({ width: o0 - o1, lines: offset });
  }
  return out;
}

/** The bands of a compound line as offset strokes (`compoundBandLines`), as SVG paths. */
function offsetBands(
  doc: Document,
  bands: readonly CompoundBandLine[],
  stroke: Record<string, string | undefined>,
  limit: number
): SVGElement[] {
  // The corners are in the offset geometry: a miter join keeps them as built.
  return bands.map((band) =>
    svg(doc, 'path', {
      d: polylineData(band.lines),
      fill: 'none',
      ...stroke,
      'stroke-width': num(band.width),
      'stroke-linejoin': 'miter',
      'stroke-miterlimit': num(Math.max(limit, 4)),
    })
  );
}

/**
 * One closed subpath that starts with its closing segment, or undefined for any other path. The
 * dashes of a compound line on a closed path start where the closing segment starts
 * (`fixtures/fills-lines.pdf`, slide 2: a dashed `tri` `roundRect`), and SVG and PDF dashes start
 * at the first point.
 */
function closingFirstCommands(commands: readonly PathCommand[]): PathCommand[] | undefined {
  const n = commands.length;
  const first = commands[0];
  const last = commands[n - 2];
  if (n < 3 || first.op !== 'M' || commands[n - 1].op !== 'Z' || last.op === 'Z') return undefined;
  if (commands.slice(1, -1).some((c) => c.op === 'M' || c.op === 'Z')) return undefined;
  if (Math.abs(last.x - first.x) < 1e-6 && Math.abs(last.y - first.y) < 1e-6) return undefined;
  const out: PathCommand[] = [
    { op: 'M', x: last.x, y: last.y },
    { op: 'L', x: first.x, y: first.y },
  ];
  for (const c of commands.slice(1, -1)) out.push(c.op === 'M' ? { op: 'L', x: c.x, y: c.y } : c);
  out.push({ op: 'Z' });
  return out;
}

/** Path data for commands from `closingFirstCommands`. */
function closingFirstData(commands: readonly PathCommand[]): string {
  const pt = (x: number, y: number): string => `${num(x)} ${num(y)}`;
  let d = '';
  for (const c of commands) {
    if (c.op === 'M') d += `M${pt(c.x, c.y)}`;
    else if (c.op === 'L') d += `L${pt(c.x, c.y)}`;
    else if (c.op === 'C') d += `C${pt(c.x1, c.y1)} ${pt(c.x2, c.y2)} ${pt(c.x, c.y)}`;
    else if (c.op === 'Q') d += `Q${pt(c.x1, c.y1)} ${pt(c.x, c.y)}`;
    else d += 'Z';
  }
  return d;
}

/**
 * How a compound line draws (design D5 of `fill-and-line-fidelity`). The SVG paint and
 * `pptx-to-pdf` share it.
 *
 * - `offset`: each band is a stroke of the path offset to the band's center (`compoundBandLines`).
 * - `masked`: each band is the line's own stroke, cut to the band by a stroke of width `2b` minus
 *   a stroke of width `2a`. On a closed path, `outer` bands lie outside the path and `inner` bands
 *   inside it; on an open path, every band is in `outer`, on both sides. `commands`, when set,
 *   replace the path, so the dashes start where the reference PDF starts them.
 * @internal
 */
export type CompoundPlan =
  | { readonly kind: 'offset'; readonly bands: readonly CompoundBandLine[] }
  | {
      readonly kind: 'masked';
      readonly closed: boolean;
      readonly commands?: readonly PathCommand[];
      readonly outer: readonly Band[];
      readonly inner: readonly Band[];
    };

/**
 * The plan of a compound line of `width` (`CompoundPlan`), or null for a single line, an unknown
 * type, or no width. A dashed line keeps the masked bands, which cut every band at the dashes of
 * the center line, except where only offset bands can draw it: an asymmetric type on an open
 * path. Without `commands`, or past `MAX_OFFSET_POINTS`, the bands are masked, and an open path
 * draws asymmetric types with `dbl` bands.
 * @internal
 */
export function compoundPlan(
  cmpd: string,
  width: number,
  closed: boolean,
  dashed: boolean,
  join: string,
  miterLimit: number,
  commands?: readonly PathCommand[]
): CompoundPlan | null {
  if (!(width > 0) || !Object.hasOwn(COMPOUND_BANDS, cmpd)) return null;
  const offset =
    commands && (!dashed || (!closed && isAsymmetric(cmpd)))
      ? compoundBandLines(commands, cmpd, width, join, miterLimit)
      : null;
  if (offset) return { kind: 'offset', bands: offset };
  const moved = dashed && closed && commands ? closingFirstCommands(commands) : undefined;
  const bands = compoundBands(!closed && isAsymmetric(cmpd) ? 'dbl' : cmpd, width);
  if (!bands) return null;
  if (!closed) {
    // Symmetric bands: both sides of the center line at once.
    const all = [...bands.outer, ...bands.inner];
    const merged = all.filter((b, i) => all.findIndex((c) => c.a === b.a && c.b === b.b) === i);
    return { kind: 'masked', closed, commands: moved, outer: merged, inner: [] };
  }
  return { kind: 'masked', closed, commands: moved, outer: bands.outer, inner: bands.inner };
}

/**
 * The SVG for a compound outline of path `given` in a `w`×`h` local box, or null when `cmpd` is a
 * single line. `stroke` is the single-line stroke (its paint, width, joins, caps, dashes).
 * `commands` are the path's commands, for the offset bands. The bands follow `compoundPlan`.
 */
export function compoundStroke(
  defs: Defs,
  given: string,
  cmpd: string,
  stroke: Record<string, string | undefined>,
  w: number,
  h: number,
  commands?: readonly PathCommand[]
): SVGElement | null {
  const width = Number(stroke['stroke-width']);
  const closed = /z\s*$/i.test(given);
  if (!Number.isFinite(width) || width <= 0 || stroke.stroke === 'none') return null;
  if (!Object.hasOwn(COMPOUND_BANDS, cmpd)) return null;
  const doc = defs.doc;
  const dashed = !!stroke['stroke-dasharray'] && stroke['stroke-dasharray'] !== 'none';
  const join = stroke['stroke-linejoin'] ?? 'miter';
  const limit = Number(stroke['stroke-miterlimit']) || 4;
  const plan = compoundPlan(cmpd, width, closed, dashed, join, limit, commands);
  if (!plan) return null;
  if (plan.kind === 'offset')
    return svg(doc, 'g', { 'data-compound': cmpd }, offsetBands(doc, plan.bands, stroke, limit));
  const d = plan.commands ? closingFirstData(plan.commands) : given;
  const grow = width + Math.max(w, h);
  const area = {
    x: num(-grow),
    y: num(-grow),
    width: num(w + 2 * grow),
    height: num(h + 2 * grow),
  };
  const painted = (maskUrl: string): SVGElement =>
    svg(doc, 'path', {
      d,
      fill: 'none',
      ...stroke,
      'stroke-width': num(width),
      mask: maskUrl,
    });
  const out = svg(doc, 'g', { 'data-compound': cmpd });
  const mask = (content: SVGElement[]): string =>
    defs.add(
      svg(doc, 'mask', { id: defs.id('cmpd'), maskUnits: 'userSpaceOnUse', ...area }, content)
    );
  if (!closed) {
    out.appendChild(painted(mask(bandStrokes(doc, d, plan.outer, stroke))));
    return out;
  }
  if (plan.outer.length)
    out.appendChild(
      painted(
        mask([
          ...bandStrokes(doc, d, plan.outer, stroke),
          // Remove the inside of the path: these bands lie outside it.
          svg(doc, 'path', { d, fill: '#000000', stroke: 'none' }),
        ])
      )
    );
  if (plan.inner.length) {
    // Keep the inside of the path with a mask, not a clipPath: Chromium can misplace a clipPath
    // edge in EMU coordinates under a translation (design D7 of `fill-fidelity`).
    const inside = mask([svg(doc, 'path', { d, fill: '#ffffff', stroke: 'none' })]);
    out.appendChild(
      svg(doc, 'g', { mask: inside }, [painted(mask(bandStrokes(doc, d, plan.inner, stroke)))])
    );
  }
  return out;
}
