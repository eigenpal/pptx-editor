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
// are drawn from the preset names of `ST_BevelPresetType` (20.1.10.9) and the reference PDF of
// the `three-d-effects` fixture; only `circle` and `angle` are fitted. An extrusion shows only
// when the camera turns: its sides are the outline swept along the turned depth. The face of a
// turned shape is the flat shape under the camera's linear transform, about its center.

import type { ResolvedThreeD, Rgba, BevelPresetType } from '@pptx-editor.dev/core/cascade';
import {
  apply,
  type GeometryPath,
  type Matrix,
  type ThreeDFacet,
  type ThreeDInput,
  type ThreeDPaint,
} from '@pptx-editor.dev/core/layout';
import { cameraView, lightingOf, normalize, turn, type Lighting } from './three-d-scene.ts';

type Pt = { readonly x: number; readonly y: number };

/** Strips per bevel. */
const STRIPS = 8;
/** Segments per curve when the outline is flattened. */
const CURVE_STEPS = 12;
/** The most outline points a 3-D paint takes: a resource limit. More draws the shape flat. */
const MAX_POINTS = 4096;

/** The height of a bevel profile at `t` (0 at the outer edge, 1 at the inner edge), 0 to 1. */
const PROFILES: Readonly<Record<BevelPresetType, (t: number) => number>> = {
  circle: (t) => Math.sqrt(Math.max(0, 1 - (1 - t) ** 2)),
  angle: (t) => t,
  relaxedInset: (t) => {
    if (t < 0.4) return Math.sqrt(Math.max(0, 1 - (1 - t / 0.4) ** 2));
    const u = (t - 0.4) / 0.6;
    return 1 - 0.5 * u * (2 - u);
  },
  slope: (t) => Math.min(1, t / 0.6),
  cross: (t) => 1 - 0.5 * Math.max(0, 1 - Math.abs(t - 0.45) / 0.1),
  softRound: (t) => 1 - (1 - t) ** 3,
  convex: (t) => Math.sin((Math.PI * t) / 2),
  coolSlant: (t) => (t < 0.2 ? (0.6 * t) / 0.2 : 0.6 + (0.4 * (t - 0.2)) / 0.8),
  divot: (t) =>
    t < 0.5
      ? Math.sqrt(Math.max(0, 1 - (1 - t / 0.5) ** 2))
      : 1 - 0.3 * Math.sin((Math.PI * (t - 0.5)) / 0.5),
  riblet: (t) => 0.5 * t + 0.5 * Math.abs(Math.sin(2 * Math.PI * t)),
  hardEdge: (t) => Math.min(1, t / 0.15),
  artDeco: (t) => (t < 0.15 ? t / 0.3 : t < 0.35 ? 0.5 : t < 0.5 ? 0.5 + (t - 0.35) / 0.3 : 1),
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
  const ch = (v: number): number => Math.max(0, Math.min(255, v * level + 255 * white));
  return { r: ch(c.r), g: ch(c.g), b: ch(c.b), a: c.a };
}

/** The strips of one bevel on one outline, in the local box. */
function bevelStrips(
  poly: readonly Pt[],
  slideNormals: readonly Pt[],
  bevel: { readonly w: number; readonly h: number; readonly preset: BevelPresetType },
  width: number,
  color: Rgba,
  light: Lighting
): ThreeDFacet[] {
  const normals = edgeNormals(poly);
  const w = Math.min(bevel.w, width);
  const profile = PROFILES[bevel.preset];
  const rings = Array.from({ length: STRIPS + 1 }, (_, k) =>
    offset(poly, normals, (w * k) / STRIPS)
  );
  const out: ThreeDFacet[] = [];
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

/** Bevels, contour, and extrusion of a shape, or null when it draws flat. */
export function threeDPaint(input: ThreeDInput): ThreeDPaint | null {
  const t: ResolvedThreeD | undefined = input.effects.threeD;
  if (!t) return null;
  const polys = outlines(input.paths);
  if (!polys || polys.length === 0) return null;
  const view = cameraView(t);
  const light = lightingOf(t);
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
  const over: ThreeDFacet[] = [];
  const under: ThreeDFacet[] = [];
  const behind: ThreeDFacet[] = [];
  const minSide = Math.max(0, Math.min(input.width, input.height) / 2);
  if (t.bevelTop)
    for (const p of polys)
      over.push(...bevelStrips(p, slideNormalsOf(p, m), t.bevelTop, minSide, color, light));
  if (t.contourW > 0) {
    const cColor = t.contourColor ? { ...t.contourColor } : color;
    for (const p of polys) {
      const normals = edgeNormals(p);
      const ring = offset(p, normals, -t.contourW);
      for (let i = 0; i < p.length; i += 1) {
        const j = (i + 1) % p.length;
        under.push({ points: [ring[i]!, ring[j]!, p[j]!, p[i]!], color: cColor });
      }
    }
  }
  const depth = t.extrusionH + (t.bevelBottom?.h ?? 0) + (t.bevelTop?.h ?? 0);
  if (depth > 0 && view.kind !== 'front') {
    const side = t.extrusionColor ? { ...t.extrusionColor, a: 1 } : color;
    const faces: { facet: ThreeDFacet; z: number }[] = [];
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
        // The light turns with the camera: a side keeps the light of its own normal.
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
  return { front, behind, under, over };
}
