// Affine matrices for layout: compose, apply, invert, and an element box's matrix.

import type { Xfrm } from '../cascade/index.ts';

/** An affine matrix `[a, b, c, d, e, f]` (SVG convention). EMU units. */
export type Matrix = readonly [number, number, number, number, number, number];

export const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];

/** `m × n`: apply `n` first, then `m`. */
export function multiply(m: Matrix, n: Matrix): Matrix {
  return [
    m[0] * n[0] + m[2] * n[1],
    m[1] * n[0] + m[3] * n[1],
    m[0] * n[2] + m[2] * n[3],
    m[1] * n[2] + m[3] * n[3],
    m[0] * n[4] + m[2] * n[5] + m[4],
    m[1] * n[4] + m[3] * n[5] + m[5],
  ];
}

/** Apply a matrix to a point. */
export function apply(m: Matrix, x: number, y: number): { x: number; y: number } {
  return { x: m[0] * x + m[2] * y + m[4], y: m[1] * x + m[3] * y + m[5] };
}

/** Invert an affine matrix. */
export function invert(m: Matrix): Matrix {
  const det = m[0] * m[3] - m[1] * m[2] || 1e-12;
  return [
    m[3] / det,
    -m[1] / det,
    -m[2] / det,
    m[0] / det,
    (m[2] * m[5] - m[3] * m[4]) / det,
    (m[1] * m[4] - m[0] * m[5]) / det,
  ];
}

const translate = (x: number, y: number): Matrix => [1, 0, 0, 1, x, y];
const scale = (sx: number, sy: number): Matrix => [sx, 0, 0, sy, 0, 0];
function rotate(deg: number): Matrix {
  const r = (deg * Math.PI) / 180;
  const c = Math.cos(r);
  const s = Math.sin(r);
  return [c, s, -s, c, 0, 0];
}

/** Box-local → parent: rotate and flip around the box center, then place at the offset. */
export function boxMatrix(x: Xfrm, withFlip = true, extraRotation = 0): Matrix {
  const cx = x.cx / 2;
  const cy = x.cy / 2;
  let m = translate(x.x + cx, x.y + cy);
  m = multiply(m, rotate(x.rot + extraRotation));
  if (withFlip) m = multiply(m, scale(x.flipH ? -1 : 1, x.flipV ? -1 : 1));
  return multiply(m, translate(-cx, -cy));
}
