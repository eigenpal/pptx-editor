/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The outer edge of the union of a shape's filled paths, for the edge fades of soft edges and
// blurs. Each filled path fills by the even-odd rule, as the body paints it. A path whose bounds
// meet no other filled path keeps its own commands. A path that meets another is flattened, and
// only the pieces of its edge with the union on one side and not on the other stay: an edge
// inside another path, or shared by two paths on opposite sides, is no outer edge. The stroked
// pieces have round caps, so a stroke of them covers the points near the outer edge.

import type { PathCommand } from '@pptx-editor.dev/core/layout';

/** The largest deviation of a flattened curve from the curve: 0.05 pt in EMU. */
const FLATTEN_TOLERANCE = 635;
/** The most segments a curve flattens into. */
const MAX_CURVE_SEGMENTS = 64;
/** The distance from an edge piece at which its two sides are tested: 0.2 pt in EMU. */
const SIDE_OFFSET = 4 * FLATTEN_TOLERANCE;
/** The longest edge piece tested on its own: 1 pt in EMU, more for a very long edge. */
const PIECE_LENGTH = 12_700;
/** The most edge pieces tested for one shape. */
const MAX_PIECES = 20_000;
/** The most flattened segments for one shape. Past it, every path keeps its own commands. */
const MAX_SEGMENTS = 8_000;

type Pt = { x: number; y: number };
type Segment = { a: Pt; b: Pt };
type Box = { x0: number; y0: number; x1: number; y1: number };

/** The closed polygons of a path, its curves flattened. */
function flatten(commands: readonly PathCommand[]): Pt[][] {
  const out: Pt[][] = [];
  let ring: Pt[] = [];
  let at: Pt = { x: 0, y: 0 };
  let start: Pt = at;
  const close = (): void => {
    if (ring.length > 1) out.push(ring);
    ring = [];
  };
  const curve = (pts: Pt[]): void => {
    // Second differences bound the deviation of a chord from the curve.
    let dd = 0;
    for (let i = 0; i + 2 < pts.length; i += 1)
      dd = Math.max(
        dd,
        Math.hypot(
          pts[i].x - 2 * pts[i + 1].x + pts[i + 2].x,
          pts[i].y - 2 * pts[i + 1].y + pts[i + 2].y
        )
      );
    const n = Math.max(
      1,
      Math.min(MAX_CURVE_SEGMENTS, Math.ceil(Math.sqrt((0.75 * dd) / FLATTEN_TOLERANCE)))
    );
    for (let s = 1; s <= n; s += 1) {
      const t = s / n;
      const u = 1 - t;
      const w =
        pts.length === 3
          ? [u * u, 2 * u * t, t * t]
          : [u * u * u, 3 * u * u * t, 3 * u * t * t, t * t * t];
      ring.push({
        x: pts.reduce((sum, q, i) => sum + w[i] * q.x, 0),
        y: pts.reduce((sum, q, i) => sum + w[i] * q.y, 0),
      });
    }
  };
  for (const c of commands) {
    if (c.op === 'M') {
      close();
      start = { x: c.x, y: c.y };
      ring.push(start);
    } else if (c.op === 'L') ring.push({ x: c.x, y: c.y });
    else if (c.op === 'C')
      curve([at, { x: c.x1, y: c.y1 }, { x: c.x2, y: c.y2 }, { x: c.x, y: c.y }]);
    else if (c.op === 'Q') curve([at, { x: c.x1, y: c.y1 }, { x: c.x, y: c.y }]);
    else {
      close();
      ring = [start];
      at = start;
      continue;
    }
    at = { x: c.x, y: c.y };
  }
  close();
  return out;
}

function boundsOf(rings: readonly Pt[][]): Box {
  const box = { x0: Infinity, y0: Infinity, x1: -Infinity, y1: -Infinity };
  for (const ring of rings)
    for (const q of ring) {
      box.x0 = Math.min(box.x0, q.x);
      box.y0 = Math.min(box.y0, q.y);
      box.x1 = Math.max(box.x1, q.x);
      box.y1 = Math.max(box.y1, q.y);
    }
  return box;
}

const meets = (a: Box, b: Box, pad: number): boolean =>
  a.x0 <= b.x1 + pad && b.x0 <= a.x1 + pad && a.y0 <= b.y1 + pad && b.y0 <= a.y1 + pad;

/** Whether `q` is inside the polygons `rings` by the even-odd rule. */
function inside(q: Pt, rings: readonly Pt[][], box: Box): boolean {
  if (q.x < box.x0 || q.x > box.x1 || q.y < box.y0 || q.y > box.y1) return false;
  let odd = false;
  for (const ring of rings)
    for (let i = 0, j = ring.length - 1; i < ring.length; j = i, i += 1) {
      const a = ring[i];
      const b = ring[j];
      if (a.y > q.y !== b.y > q.y && q.x < ((b.x - a.x) * (q.y - a.y)) / (b.y - a.y) + a.x)
        odd = !odd;
    }
  return odd;
}

/** The segments of the closed polygons `rings`, each closing edge included. */
function segmentsOf(rings: readonly Pt[][]): Segment[][] {
  return rings.map((ring) =>
    ring
      .map((a, i) => ({ a, b: ring[(i + 1) % ring.length] }))
      .filter((s) => s.a.x !== s.b.x || s.a.y !== s.b.y)
  );
}

/**
 * The commands to stroke for the outer edge of the union of `filled`, one entry per filled path:
 * its own commands when its bounds meet no other path, else open polylines of the pieces of its
 * edge that lie on the outer edge (empty when none does).
 */
export function outerEdges(filled: readonly PathCommand[][]): PathCommand[][] {
  if (filled.length < 2) return filled.map((path) => [...path]);
  const rings = filled.map(flatten);
  const boxes = rings.map(boundsOf);
  const pad = 2 * SIDE_OFFSET;
  const overlapping = boxes.map((box, i) =>
    boxes.some((other, j) => j !== i && meets(box, other, pad))
  );
  if (!overlapping.some(Boolean)) return filled.map((path) => [...path]);
  const segments = rings.map(segmentsOf);
  const count = segments.reduce((n, path) => n + path.reduce((m, ring) => m + ring.length, 0), 0);
  if (count > MAX_SEGMENTS) return filled.map((path) => [...path]);
  let perimeter = 0;
  segments.forEach((path, i) => {
    if (!overlapping[i]) return;
    for (const ring of path)
      for (const s of ring) perimeter += Math.hypot(s.b.x - s.a.x, s.b.y - s.a.y);
  });
  const step = Math.max(PIECE_LENGTH, perimeter / MAX_PIECES);
  const covered = (q: Pt): boolean => rings.some((r, j) => inside(q, r, boxes[j]));
  return filled.map((path, i) => {
    if (!overlapping[i]) return [...path];
    const out: PathCommand[] = [];
    for (const ring of segments[i]) {
      let open = false;
      for (const { a, b } of ring) {
        const len = Math.hypot(b.x - a.x, b.y - a.y);
        const n = Math.max(1, Math.ceil(len / step));
        const nx = (-(b.y - a.y) / len) * SIDE_OFFSET;
        const ny = ((b.x - a.x) / len) * SIDE_OFFSET;
        for (let k = 0; k < n; k += 1) {
          const from = { x: a.x + ((b.x - a.x) * k) / n, y: a.y + ((b.y - a.y) * k) / n };
          const to = { x: a.x + ((b.x - a.x) * (k + 1)) / n, y: a.y + ((b.y - a.y) * (k + 1)) / n };
          const mx = (from.x + to.x) / 2;
          const my = (from.y + to.y) / 2;
          const outer = covered({ x: mx + nx, y: my + ny }) !== covered({ x: mx - nx, y: my - ny });
          if (!outer) open = false;
          else {
            if (!open) out.push({ op: 'M', x: from.x, y: from.y });
            out.push({ op: 'L', x: to.x, y: to.y });
            open = true;
          }
        }
      }
    }
    return out;
  });
}
