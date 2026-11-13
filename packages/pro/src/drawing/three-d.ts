/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Bevels, contours, and extrusion as flat polygons (ECMA-376 Part 1, 20.1.5.1 bevelT,
// 20.1.5.12 sp3d). Pure math: no DOM. The SVG paint and the PDF writer fill the same polygons.
//
// A bevel is a band along the outline, cut into strips. Each strip is tilted by the slope of
// the preset's profile, and the light of the scene shades it (`three-d-scene.ts`). The profiles
// of `ST_BevelPresetType` (20.1.10.9) are measured on the reference PDF of the `three-d-effects`
// fixture: `circle` and `angle` follow their names, and the others are slope tables. The light
// also shades the face itself (`faceColor`). An extrusion shows only when the camera turns: its
// sides are the outline swept along the turned depth. The face of a turned shape is the flat
// shape under the camera's linear transform, about its center.

import type { ResolvedThreeD, Rgba, BevelPresetType } from '@pptx-editor.dev/core/cascade';
import {
  apply,
  type GeometryPath,
  type Matrix,
  type ThreeDFacet,
  type ThreeDInput,
  type ThreeDPaint,
} from '@pptx-editor.dev/core/layout';
import {
  cameraView,
  lightingOf,
  normalize,
  turn,
  viewDirection,
  type Lighting,
} from './three-d-scene.ts';

type Pt = { readonly x: number; readonly y: number };

/** Strips per bevel: one for each measured slope. */
const STRIPS = 12;
/** Segments per curve when the outline is flattened. */
const CURVE_STEPS = 12;
/** The most outline points a 3-D paint takes: a resource limit. More draws the shape flat. */
const MAX_POINTS = 4096;
/**
 * The most outline points that bevel strips, contours, and extrusion sides follow: a longer
 * outline keeps every k-th point, so one shape paints at most about 15,000 polygons.
 */
const MAX_THREE_D_POINTS = 1024;

/**
 * The measured slopes of the presets that are not `circle` or `angle`: the rise of the profile
 * over each twelfth of the band, from the outer edge in, for a bevel as high as it is wide.
 * Each slope is the one that gives, under the light of the scene, the colors of the reference
 * PDF of the `three-d-effects` fixture (slide 1) on the four sides of the band. A negative slope
 * falls toward the face: an inset.
 */
const MEASURED_SLOPES: Readonly<Record<Exclude<BevelPresetType, 'circle' | 'angle'>, number[]>> = {
  relaxedInset: [5, 2.42, 1.5, 0.88, 0.82, -1.22, -1.32, -1.06, -0.86, -0.24, -0.24, -0.12],
  cross: [0.88, 0.02, 0, 0, 0, 1.08, 0.88, 0.04, 0, 0, 0, 0],
  coolSlant: [5, 4.04, 2.96, 1.96, 0.78, 0.18, 0.06, 0, 0, 0, 0, 0],
  softRound: [5, 2.74, -1.76, -1.84, -1.38, -1.02, -0.84, -0.68, -0.3, -0.28, -0.2, -0.08],
  convex: [0.88, 0.08, 0.24, 0.36, 0.68, 0.82, 1, 1.34, 1.64, 2.16, 1.9, 0.1],
  slope: [5, 5, 1.58, 0.16, 0.04, -0.02, -0.12, -0.78, -2.22, -5, -2.42, -0.14],
  divot: [5, 3.12, 1.7, 0.96, 0.74, 0.28, 0.18, 0, 0, 0.02, 0.1, 0.1],
  riblet: [5, 2.02, 1.42, 1.22, 1, 0.82, 0.28, 0.2, 0.06, -0.02, 0.1, 0.12],
  hardEdge: [5, 4.84, 0.16, 0.02, 0, 0, -0.08, -0.78, -0.66, -0.26, -0.12, -0.02],
  artDeco: [2.04, 0.8, 0.3, 0.18, 2.2, 1.68, 1, 0.74, 0.32, 0.3, 0.2, 0.08],
};

/** A profile that rises by `slopes[k] / n` over each of its `n` equal parts. */
function fromSlopes(slopes: readonly number[]): (t: number) => number {
  const n = slopes.length;
  const heights = [0];
  for (const s of slopes) heights.push(heights[heights.length - 1]! + s / n);
  return (t) => {
    const u = Math.max(0, Math.min(1, t)) * n;
    const k = Math.min(n - 1, Math.floor(u));
    return heights[k]! + (u - k) * (heights[k + 1]! - heights[k]!);
  };
}

/** The height of a bevel profile at `t` (0 at the outer edge, 1 at the inner edge). */
const PROFILES: Readonly<Record<BevelPresetType, (t: number) => number>> = {
  circle: (t) => Math.sqrt(Math.max(0, 1 - (1 - t) ** 2)),
  angle: (t) => t,
  ...(Object.fromEntries(
    Object.entries(MEASURED_SLOPES).map(([name, slopes]) => [name, fromSlopes(slopes)])
  ) as Record<Exclude<BevelPresetType, 'circle' | 'angle'>, (t: number) => number>),
};

/** The closed outlines of the filled paths, flattened to polygons in the local box. */
export function outlines(paths: readonly GeometryPath[]): Pt[][] | null {
  const out: Pt[][] = [];
  let total = 0;
  for (const path of paths) {
    if (path.fill === 'none') continue;
    let cur: Pt[] = [];
    let at: Pt = { x: 0, y: 0 };
    const close = (): void => {
      // Drop a closing point that repeats the first.
      const last = cur[cur.length - 1];
      if (last && cur.length > 1 && Math.hypot(last.x - cur[0]!.x, last.y - cur[0]!.y) < 1)
        cur.pop();
      if (cur.length >= 3) out.push(cur);
      total += cur.length;
      cur = [];
    };
    const add = (p: Pt): void => {
      const last = cur[cur.length - 1];
      if (!last || Math.hypot(last.x - p.x, last.y - p.y) >= 1) cur.push(p);
      at = p;
    };
    for (const c of path.commands) {
      if (c.op === 'M') {
        close();
        add({ x: c.x, y: c.y });
      } else if (c.op === 'L') add({ x: c.x, y: c.y });
      else if (c.op === 'C') {
        const p0 = at;
        for (let i = 1; i <= CURVE_STEPS; i += 1) {
          const t = i / CURVE_STEPS;
          const u = 1 - t;
          add({
            x: u * u * u * p0.x + 3 * u * u * t * c.x1 + 3 * u * t * t * c.x2 + t * t * t * c.x,
            y: u * u * u * p0.y + 3 * u * u * t * c.y1 + 3 * u * t * t * c.y2 + t * t * t * c.y,
          });
        }
      } else if (c.op === 'Q') {
        const p0 = at;
        for (let i = 1; i <= CURVE_STEPS; i += 1) {
          const t = i / CURVE_STEPS;
          const u = 1 - t;
          add({
            x: u * u * p0.x + 2 * u * t * c.x1 + t * t * c.x,
            y: u * u * p0.y + 2 * u * t * c.y1 + t * t * c.y,
          });
        }
      } else if (c.op === 'Z') {
        const first = cur[0];
        close();
        if (first) at = first;
      }
      if (total + cur.length > MAX_POINTS) return null;
    }
    close();
  }
  return out;
}

/** The outlines with at most `MAX_THREE_D_POINTS` points in all: every k-th point of each. */
function thinned(polys: readonly Pt[][]): Pt[][] {
  const total = polys.reduce((n, p) => n + p.length, 0);
  const k = Math.ceil(total / MAX_THREE_D_POINTS);
  if (k <= 1) return [...polys];
  return polys.map((p) => p.filter((_, i) => i % k === 0)).filter((p) => p.length >= 3);
}

function signedArea(poly: readonly Pt[]): number {
  let a = 0;
  for (let i = 0; i < poly.length; i += 1) {
    const p = poly[i]!;
    const q = poly[(i + 1) % poly.length]!;
    a += p.x * q.y - q.x * p.y;
  }
  return a / 2;
}

/** The outward unit normal of each edge `i` (from point `i` to `i + 1`). */
function edgeNormals(poly: readonly Pt[]): Pt[] {
  const sign = signedArea(poly) >= 0 ? 1 : -1;
  return poly.map((p, i) => {
    const q = poly[(i + 1) % poly.length]!;
    const ex = q.x - p.x;
    const ey = q.y - p.y;
    const l = Math.hypot(ex, ey) || 1;
    return { x: (sign * ey) / l, y: (sign * -ex) / l };
  });
}

/** The outline moved inward by `d` (outward when negative), with mitered corners. */
function offset(poly: readonly Pt[], normals: readonly Pt[], d: number): Pt[] {
  const n = poly.length;
  return poly.map((p, i) => {
    const a = normals[(i - 1 + n) % n]!;
    const b = normals[i]!;
    // The miter (a + b) / (1 + a·b), at most four times the distance on a sharp corner.
    const den = Math.max(0.125, 1 + a.x * b.x + a.y * b.y);
    let mx = (a.x + b.x) / den;
    let my = (a.y + b.y) / den;
    const l = Math.hypot(mx, my);
    if (l > 4) {
      mx *= 4 / l;
      my *= 4 / l;
    }
    return { x: p.x - mx * d, y: p.y - my * d };
  });
}

/** A color under a light level and a white highlight, as an opaque fill. */
function lit(c: Rgba, level: number, white: number): Rgba {
  // A degenerate normal gives no light value: keep the color unshaded.
  const k = Number.isFinite(level) ? level : 1;
  const wh = Number.isFinite(white) ? white : 0;
  const ch = (v: number): number => Math.max(0, Math.min(255, v * k + 255 * wh));
  return { r: ch(c.r), g: ch(c.g), b: ch(c.b), a: c.a };
}

/** The smallest outline area that a bevel shades, in EMU²: less is a line, with no inside. */
const MIN_AREA = 1;

type Quad = { readonly points: readonly Pt[]; readonly color: Rgba };

const finite = (q: Quad): boolean =>
  q.points.every((p) => Number.isFinite(p.x) && Number.isFinite(p.y)) &&
  [q.color.r, q.color.g, q.color.b, q.color.a].every(Number.isFinite);

const colorKey = (c: Rgba): string =>
  `${Math.round(c.r)},${Math.round(c.g)},${Math.round(c.b)},${Math.round(c.a * 1000)}`;

/**
 * Neighboring polygons of one color as one facet, so a long outline paints a path for each run
 * of a color, not for each polygon. Only neighbors merge: the paint order stays (far sides
 * first, and inner bevel strips over outer ones). Polygons with values that are not finite are
 * dropped.
 */
function merge(quads: readonly Quad[]): ThreeDFacet[] {
  const out: { polygons: (readonly Pt[])[]; color: Rgba; key: string }[] = [];
  for (const q of quads) {
    if (!finite(q)) continue;
    const key = colorKey(q.color);
    const last = out[out.length - 1];
    if (last?.key === key) last.polygons.push(q.points);
    else out.push({ polygons: [q.points], color: q.color, key });
  }
  return out.map(({ polygons, color }) => ({ polygons, color }));
}

/** The strips of one bevel on one outline, in the local box. */
function bevelStrips(
  poly: readonly Pt[],
  slideNormals: readonly Pt[],
  bevel: { readonly w: number; readonly h: number; readonly preset: BevelPresetType },
  width: number,
  color: Rgba,
  light: Lighting
): Quad[] {
  const w = Math.min(bevel.w, width);
  // A zero-width bevel or an outline with no inside (collinear points) has no strips.
  if (!(w > 0) || !(Math.abs(signedArea(poly)) >= MIN_AREA)) return [];
  const normals = edgeNormals(poly);
  const profile = PROFILES[bevel.preset];
  const rings = Array.from({ length: STRIPS + 1 }, (_, k) =>
    offset(poly, normals, (w * k) / STRIPS)
  );
  const out: Quad[] = [];
  for (let k = 0; k < STRIPS; k += 1) {
    const t0 = k / STRIPS;
    const t1 = (k + 1) / STRIPS;
    const slope = ((profile(t1) - profile(t0)) / (t1 - t0)) * (bevel.h / w);
    const outer = rings[k]!;
    const inner = rings[k + 1]!;
    for (let i = 0; i < poly.length; i += 1) {
      const j = (i + 1) % poly.length;
      const n = slideNormals[i]!;
      const { level, white } = light.shade(normalize([n.x * slope, n.y * slope, 1]));
      out.push({
        points: [outer[i]!, outer[j]!, inner[j]!, inner[i]!],
        color: lit(color, level, white),
      });
    }
  }
  return out;
}

/** The outward normals of an outline in slide space: the local normals under the matrix. */
function slideNormalsOf(poly: readonly Pt[], m: Matrix): Pt[] {
  return edgeNormals(poly.map((p) => apply(m, p.x, p.y)));
}

const IDENTITY_MATRIX: Matrix = [1, 0, 0, 1, 0, 0];

/** The smallest area scale of a turned face that still paints. */
const MIN_DETERMINANT = 1e-6;

/** Bevels, contour, and extrusion of a shape, or null when it draws flat. */
export function threeDPaint(input: ThreeDInput): ThreeDPaint | null {
  const t: ResolvedThreeD | undefined = input.effects.threeD;
  if (!t) return null;
  const full = outlines(input.paths);
  if (!full || full.length === 0) return null;
  const polys = thinned(full);
  const view = cameraView(t);
  const eye = viewDirection(view);
  const light = lightingOf(t, eye);
  const m = input.matrix;
  const c = apply(m, input.width / 2, input.height / 2);
  const front: Matrix =
    view.kind === 'turn'
      ? (() => {
          const R = view.rotation;
          const a = R[0][0];
          const b = R[1][0];
          const cc = R[0][1];
          const d = R[1][1];
          return [a, b, cc, d, c.x - a * c.x - cc * c.y, c.y - b * c.x - d * c.y];
        })()
      : IDENTITY_MATRIX;
  const color = { ...input.color, a: 1 };
  const over: Quad[] = [];
  const under: Quad[] = [];
  const behind: Quad[] = [];
  const minSide = Math.max(0, Math.min(input.width, input.height) / 2);
  if (t.bevelTop)
    for (const p of polys)
      over.push(...bevelStrips(p, slideNormalsOf(p, m), t.bevelTop, minSide, color, light));
  if (t.contourW > 0) {
    const cColor = t.contourColor ? { ...t.contourColor } : color;
    for (const p of polys) {
      const normals = edgeNormals(p);
      // The contour lies across the edge: half of it shows outside the outline, and the face
      // covers the rest (measured on the fixture, slide 2).
      const ring = offset(p, normals, -t.contourW / 2);
      for (let i = 0; i < p.length; i += 1) {
        const j = (i + 1) % p.length;
        under.push({ points: [ring[i]!, ring[j]!, p[j]!, p[i]!], color: cColor });
      }
    }
  }
  const depth = t.extrusionH + (t.bevelBottom?.h ?? 0) + (t.bevelTop?.h ?? 0);
  // The back of a turned extrusion, from the face: the turned depth axis.
  const back =
    depth > 0 && view.kind === 'turn'
      ? { x: view.rotation[0][2] * depth, y: view.rotation[1][2] * depth }
      : undefined;
  if (depth > 0 && view.kind !== 'front') {
    const side = t.extrusionColor ? { ...t.extrusionColor, a: 1 } : color;
    const faces: { facet: Quad; z: number }[] = [];
    for (const p of polys) {
      const s = p.map((q) => apply(m, q.x, q.y));
      const normals = edgeNormals(s);
      for (let i = 0; i < s.length; i += 1) {
        const a = s[i]!;
        const b = s[(i + 1) % s.length]!;
        const n = normals[i]!;
        if (view.kind === 'oblique') {
          if (n.x * view.dx + n.y * view.dy <= 0) continue;
          const e = { x: view.dx * depth, y: view.dy * depth };
          const { level, white } = light.shade(normalize([n.x, n.y, 0.35]));
          faces.push({
            facet: {
              points: [a, b, { x: b.x + e.x, y: b.y + e.y }, { x: a.x + e.x, y: a.y + e.y }],
              color: lit(side, level, white),
            },
            z: 0,
          });
          continue;
        }
        const R = view.rotation;
        const proj = (q: Pt, z: number): { x: number; y: number; z: number } => {
          const r = turn(R, [q.x - c.x, q.y - c.y, z]);
          return { x: c.x + r[0], y: c.y + r[1], z: r[2] };
        };
        const corners = [proj(a, 0), proj(b, 0), proj(b, depth), proj(a, depth)];
        // The light lies in the shape's axes: a side takes the light of its own normal.
        const { level, white } = light.shade([n.x, n.y, 0]);
        faces.push({
          facet: { points: corners.map(({ x, y }) => ({ x, y })), color: lit(side, level, white) },
          z: corners.reduce((sum, q) => sum + q.z, 0) / 4,
        });
      }
    }
    // Far sides first (depth grows away from the viewer), so near sides cover them.
    faces.sort((p, q) => q.z - p.z);
    behind.push(...faces.map((f) => f.facet));
    // The front cap in the side color, under the face: the face's antialiased edge then blends
    // with the extrusion, not with the slide.
    for (const p of polys) {
      const cap = p.map((q) => {
        const s = apply(m, q.x, q.y);
        return apply(front, s.x, s.y);
      });
      behind.push({ points: cap, color: side });
    }
  }
  if (over.length === 0 && under.length === 0 && behind.length === 0 && view.kind !== 'turn')
    return null;
  // A face turned edge-on has no area: only the extrusion shows (a turn of 90 degrees).
  const showsFace = Math.abs(front[0] * front[3] - front[1] * front[2]) >= MIN_DETERMINANT;
  // The light of the face itself, for a solid fill: the face is a surface of the scene too.
  // A camera behind the shape sees the back of the face.
  const faceShade = light.shade([0, 0, eye[2] < 0 ? -1 : 1]);
  return {
    front,
    showsFace,
    faceColor: lit(input.color, faceShade.level, faceShade.white),
    ...(back && Number.isFinite(back.x) && Number.isFinite(back.y) ? { backOffset: back } : {}),
    behind: merge(behind),
    under: merge(under),
    over: merge(over),
  };
}
