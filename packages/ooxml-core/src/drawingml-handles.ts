// Adjust handles of preset shapes (ECMA-376 Part 1, 20.1.9.1 `a:ahLst`): where each handle
// sits at a shape's size and adjust values, and the values that put a handle at a point.
//
// A handle's place (`a:pos`) is a pair of guide formulas over the adjust values, so a drag
// cannot read the values back from the point in closed form for every preset. `adjustForHandle`
// searches instead: for each guide the handle moves, in turn, it finds the value in the
// handle's range whose place is nearest to the point, and repeats while that brings the handle
// nearer. The preset's own formulas give every place, so the result is the one the shape
// draws.

import { presetAdjustments, presetGuides } from './drawingml-geometry.ts';

/** One adjust handle of a shape, in shape-local EMU, with the guides it moves. */
export interface PresetHandle {
  readonly kind: 'xy' | 'polar';
  readonly x: number;
  readonly y: number;
  /** The adjust guides this handle moves (`a:avLst` names): x and y, or radius and angle. */
  readonly guides: readonly string[];
}

/** A full circle in guide units (1/60 000 of a degree). */
const CIRCLE = 21_600_000;
/** Samples of one guide's range before the search narrows to the best one. */
const SAMPLES = 48;
/** Rounds of the coordinate search over the guides of one handle. */
const ROUNDS = 3;
/** Steps on each side of the start grid of a handle that moves two guides. */
const GRID = 16;

/**
 * The adjust handles of a preset at the size `w` by `h` (EMU) with the shape's own adjust
 * values, in `a:ahLst` order. Empty for a preset with no handles or one the engine does not
 * know, and for a shape of no size.
 */
export function presetHandles(
  preset: string,
  w: number,
  h: number,
  adjust: ReadonlyMap<string, number>
): PresetHandle[] {
  if (!(w > 0 && h > 0)) return [];
  const g = presetGuides(preset, w, h, adjust);
  if (!g?.def.hp) return [];
  const value = (token: string): number => {
    const n = Number(token);
    return Number.isFinite(n) ? n : (g.vars.get(token) ?? 0);
  };
  const avNames = new Set(g.def.av.map(([name]) => name));
  return g.def.hp.map(([kind, px, py, a, b]) => {
    const y = value(py);
    return {
      kind,
      x: value(px),
      y: g.mirrored ? h - y : y,
      guides: [a, b].filter((n): n is string => n !== null && avNames.has(n)),
    };
  });
}

/**
 * The values of the guides that handle `index` moves, which put it nearest to `point`
 * (shape-local EMU), each within its handle range. Undefined when the shape has no such handle.
 * Values are whole guide units, as `a:gd` stores them.
 */
export function adjustForHandle(
  preset: string,
  w: number,
  h: number,
  adjust: ReadonlyMap<string, number>,
  index: number,
  point: { readonly x: number; readonly y: number }
): Map<string, number> | undefined {
  const handle = presetHandles(preset, w, h, adjust)[index];
  if (!handle || handle.guides.length === 0) return undefined;
  const values = new Map<string, number>();
  for (const a of presetAdjustments(preset, w, h, adjust) ?? []) values.set(a.name, a.value);
  for (const [name, v] of adjust) values.set(name, v);
  const distance = (trial: Map<string, number>): number => {
    const p = presetHandles(preset, w, h, trial)[index];
    return p ? (p.x - point.x) ** 2 + (p.y - point.y) ** 2 : Number.POSITIVE_INFINITY;
  };
  const rangeOf = (guide: string): readonly [number, number] | undefined => {
    const slot = presetAdjustments(preset, w, h, values)?.find((a) => a.name === guide);
    if (!slot) return undefined;
    const lo = slot.min ?? 0;
    const hi = slot.max ?? (slot.angle ? CIRCLE - 1 : 100_000);
    return hi >= lo ? [lo, hi] : undefined;
  };
  // A handle that moves two guides starts from the best point of a coarse grid: one guide
  // alone can lead to a place where the other has no effect (a radius of zero).
  const [ga, gb] = handle.guides;
  const ra = ga === undefined ? undefined : rangeOf(ga);
  const rb = gb === undefined ? undefined : rangeOf(gb);
  if (ga !== undefined && gb !== undefined && ra && rb) {
    let start: [number, number] | null = null;
    let near = distance(values);
    for (let i = 0; i <= GRID; i++)
      for (let j = 0; j <= GRID; j++) {
        const va = ra[0] + ((ra[1] - ra[0]) * i) / GRID;
        const vb = rb[0] + ((rb[1] - rb[0]) * j) / GRID;
        const d = distance(new Map(values).set(ga, va).set(gb, vb));
        if (d < near) [near, start] = [d, [va, vb]];
      }
    if (start) values.set(ga, start[0]).set(gb, start[1]);
  }
  let best = distance(values);
  for (let round = 0; round < ROUNDS; round++) {
    const before = best;
    for (const guide of handle.guides) {
      // The range of this guide's handle, with the other values as they are now.
      const range = rangeOf(guide);
      if (!range) continue;
      const [lo, hi] = range;
      const at = (v: number): number => distance(new Map(values).set(guide, v));
      let x = values.get(guide) ?? lo;
      let fx = at(x);
      const step = (hi - lo) / SAMPLES;
      for (let i = 0; i <= SAMPLES; i++) {
        const v = lo + step * i;
        const fv = at(v);
        if (fv < fx) [x, fx] = [v, fv];
      }
      // Narrow the bracket around the best sample (golden section).
      let a = Math.max(lo, x - step);
      let b = Math.min(hi, x + step);
      const r = (Math.sqrt(5) - 1) / 2;
      while (b - a > 1) {
        const c = b - r * (b - a);
        const d = a + r * (b - a);
        if (at(c) < at(d)) b = d;
        else a = c;
      }
      for (const v of [Math.floor(a), Math.ceil(b), Math.round(x)]) {
        const clamped = Math.min(Math.floor(hi), Math.max(Math.ceil(lo), v));
        const fv = at(clamped);
        if (fv <= fx) [x, fx] = [clamped, fv];
      }
      values.set(guide, Math.round(x));
      best = fx;
    }
    if (best >= before) break;
  }
  return new Map(handle.guides.map((g) => [g, values.get(g) ?? 0]));
}
