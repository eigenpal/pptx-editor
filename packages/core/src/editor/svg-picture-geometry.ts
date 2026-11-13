// Geometry of the SVG icon subset, for `svg-picture-drawing.ts`: numbers, transform lists,
// view box fits, path data, and basic shapes, as absolute path commands with cubic curves.
// The input is untrusted: numbers are finite, and a path past `MAX_DRAWING_COMMANDS` throws
// `SvgUnsupported`.

import type { PathCommand } from '@eigenpal/ooxml-core';
import type { Matrix } from '../layout/index.ts';

/** The most path commands one drawing holds, with each `<use>` expanded and each arc as curves. */
export const MAX_DRAWING_COMMANDS = 200_000;

/** The most style rules (one per selector) a picture may hold; one with more draws its PNG. */
export const MAX_STYLE_RULES = 1_000;

/** Thrown for a feature the drawing reader does not draw; the reader then gives `undefined`. */
export class SvgUnsupported extends Error {}

/**
 * The largest magnitude of any number in a drawing: coordinates, matrix entries, and line
 * widths. Every number of a drawing reaches a PDF operator, which takes plain decimals.
 */
export const MAX_SVG_NUMBER = 1e9;
/** Numbers closer to zero than this are written as zero, so no decimal runs long. */
const NEAR_ZERO = 1e-9;

/** A number of a drawing, finite and within `MAX_SVG_NUMBER`; throws `SvgUnsupported` otherwise. */
export function drawingNumber(v: number): number {
  if (!Number.isFinite(v) || Math.abs(v) > MAX_SVG_NUMBER) throw new SvgUnsupported('number');
  return Math.abs(v) < NEAR_ZERO ? 0 : v;
}

/** A path with every number checked by `drawingNumber`. */
export const drawingPath = (path: readonly PathCommand[]): PathCommand[] =>
  path.map((c) => {
    if (c.op === 'Z') return c;
    const n = drawingNumber;
    if (c.op === 'C')
      return { op: 'C', x1: n(c.x1), y1: n(c.y1), x2: n(c.x2), y2: n(c.y2), x: n(c.x), y: n(c.y) };
    if (c.op === 'Q') return { op: 'Q', x1: n(c.x1), y1: n(c.y1), x: n(c.x), y: n(c.y) };
    return { op: c.op, x: n(c.x), y: n(c.y) };
  });

export const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];

export const multiply = (m: Matrix, n: Matrix): Matrix => [
  m[0] * n[0] + m[2] * n[1],
  m[1] * n[0] + m[3] * n[1],
  m[0] * n[2] + m[2] * n[3],
  m[1] * n[2] + m[3] * n[3],
  m[0] * n[4] + m[2] * n[5] + m[4],
  m[1] * n[4] + m[3] * n[5] + m[5],
];

const NUMBER = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i;

export function num(v: string | undefined): number | undefined {
  if (v === undefined) return undefined;
  const s = v.trim().replace(/px$/i, '');
  if (!NUMBER.test(s)) return undefined;
  const n = Number(s);
  return Number.isFinite(n) ? n : undefined;
}

export function numberList(v: string | undefined): number[] | undefined {
  if (v === undefined) return undefined;
  const words = v.trim().split(/[\s,]+/);
  if (words.length === 1 && words[0] === '') return [];
  const out: number[] = [];
  for (const w of words) {
    const n = num(w);
    if (n === undefined || /px$/i.test(w)) return undefined;
    out.push(n);
  }
  return out;
}

/** The matrix of a `transform` list, or undefined for one it cannot read. */
export function parseTransform(value: string): Matrix | undefined {
  let m = IDENTITY;
  const rest = value
    .trim()
    .replace(/(matrix|translate|scale|rotate|skewX|skewY)\s*\(([^()]*)\)[\s,]*/g, (_, f, a) => {
      const v = numberList(a);
      if (!v) return '!';
      const rad = (d: number) => (d * Math.PI) / 180;
      let n: Matrix | undefined;
      if (f === 'matrix' && v.length === 6) n = v as unknown as Matrix;
      else if (f === 'translate' && (v.length === 1 || v.length === 2))
        n = [1, 0, 0, 1, v[0]!, v[1] ?? 0];
      else if (f === 'scale' && (v.length === 1 || v.length === 2))
        n = [v[0]!, 0, 0, v[1] ?? v[0]!, 0, 0];
      else if (f === 'rotate' && (v.length === 1 || v.length === 3)) {
        const c = Math.cos(rad(v[0]!));
        const s = Math.sin(rad(v[0]!));
        const [cx, cy] = [v[1] ?? 0, v[2] ?? 0];
        n = [c, s, -s, c, cx - c * cx + s * cy, cy - s * cx - c * cy];
      } else if (f === 'skewX' && v.length === 1) n = [1, 0, Math.tan(rad(v[0]!)), 1, 0, 0];
      else if (f === 'skewY' && v.length === 1) n = [1, Math.tan(rad(v[0]!)), 0, 1, 0, 0];
      if (!n) return '!';
      m = multiply(m, n);
      return '';
    });
  return rest === '' && m.every(Number.isFinite) ? m : undefined;
}

/** The transform that fits a view box into a viewport, by `preserveAspectRatio` (no `slice`). */
export function viewBoxMatrix(
  vb: readonly number[],
  x: number,
  y: number,
  w: number,
  h: number,
  par: string | undefined
): Matrix {
  const sx = w / vb[2]!;
  const sy = h / vb[3]!;
  const align = (par ?? 'xMidYMid').trim().split(/\s+/)[0]!;
  if (align === 'none') return [sx, 0, 0, sy, x - vb[0]! * sx, y - vb[1]! * sy];
  const s = Math.min(sx, sy);
  const ax = /xMid/.test(align) ? 0.5 : /xMax/.test(align) ? 1 : 0;
  const ay = /YMid/.test(align) ? 0.5 : /YMax/.test(align) ? 1 : 0;
  const tx = x - vb[0]! * s + (w - vb[2]! * s) * ax;
  const ty = y - vb[1]! * s + (h - vb[3]! * s) * ay;
  return [s, 0, 0, s, tx, ty];
}

// --- Path data --------------------------------------------------------------------------

/** One SVG arc as cubic curves (SVG 1.1, F.6.5 and F.6.6). */
function arcCurves(
  x1: number,
  y1: number,
  rxIn: number,
  ryIn: number,
  angle: number,
  large: boolean,
  sweep: boolean,
  x2: number,
  y2: number,
  out: PathCommand[]
): void {
  let rx = Math.abs(rxIn);
  let ry = Math.abs(ryIn);
  if ((x1 === x2 && y1 === y2) || rx === 0 || ry === 0) {
    if (x1 !== x2 || y1 !== y2) out.push({ op: 'L', x: x2, y: y2 });
    return;
  }
  const phi = (angle * Math.PI) / 180;
  const cos = Math.cos(phi);
  const sin = Math.sin(phi);
  const dx = (x1 - x2) / 2;
  const dy = (y1 - y2) / 2;
  const xp = cos * dx + sin * dy;
  const yp = -sin * dx + cos * dy;
  // Shares of the radii, so that a huge radius cannot overflow a square (SVG 1.1, F.6.6).
  const lambda = (xp / rx) ** 2 + (yp / ry) ** 2;
  // Radii so large that the arc's bulge is far below a point draw as the straight chord, as the
  // arc of an infinite radius does.
  if (!(lambda > 1e-12)) {
    out.push({ op: 'L', x: x2, y: y2 });
    return;
  }
  if (lambda > 1) {
    // Out-of-range radii scale up until the ellipse just reaches both ends.
    rx *= Math.sqrt(lambda);
    ry *= Math.sqrt(lambda);
  }
  // (rx²ry² − rx²yp² − ry²xp²) / (rx²yp² + ry²xp²), divided through by rx²ry².
  let k = Math.sqrt(Math.max(0, (1 - Math.min(1, lambda)) / Math.min(1, lambda)));
  if (large === sweep) k = -k;
  const cxp = (k * rx * yp) / ry;
  const cyp = (-k * ry * xp) / rx;
  const cx = cos * cxp - sin * cyp + (x1 + x2) / 2;
  const cy = sin * cxp + cos * cyp + (y1 + y2) / 2;
  const ang = (ux: number, uy: number, vx: number, vy: number) => {
    const a = Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy);
    return a;
  };
  if (![cx, cy, rx, ry].every(Number.isFinite)) {
    out.push({ op: 'L', x: x2, y: y2 });
    return;
  }
  const t1 = ang(1, 0, (xp - cxp) / rx, (yp - cyp) / ry);
  let dt = ang((xp - cxp) / rx, (yp - cyp) / ry, (-xp - cxp) / rx, (-yp - cyp) / ry);
  if (!sweep && dt > 0) dt -= 2 * Math.PI;
  else if (sweep && dt < 0) dt += 2 * Math.PI;
  const segments = Math.max(1, Math.ceil(Math.abs(dt) / (Math.PI / 2) - 1e-9));
  const step = dt / segments;
  const alpha = (4 / 3) * Math.tan(step / 4);
  const point = (t: number): [number, number, number, number] => {
    const ct = Math.cos(t);
    const st = Math.sin(t);
    // The point and its derivative on the ellipse.
    return [
      cx + rx * ct * cos - ry * st * sin,
      cy + rx * ct * sin + ry * st * cos,
      -rx * st * cos - ry * ct * sin,
      -rx * st * sin + ry * ct * cos,
    ];
  };
  let t = t1;
  let [px, py, pdx, pdy] = point(t);
  for (let i = 0; i < segments; i += 1) {
    const t2 = t + step;
    const [qx, qy, qdx, qdy] = point(t2);
    out.push({
      op: 'C',
      x1: px + alpha * pdx,
      y1: py + alpha * pdy,
      x2: qx - alpha * qdx,
      y2: qy - alpha * qdy,
      x: i === segments - 1 ? x2 : qx,
      y: i === segments - 1 ? y2 : qy,
    });
    [px, py, pdx, pdy] = [qx, qy, qdx, qdy];
    t = t2;
  }
}

const COMMAND = /[MmLlHhVvCcSsQqTtAaZz]/;
const ARGS: Readonly<Record<string, number>> = {
  m: 2,
  l: 2,
  h: 1,
  v: 1,
  c: 6,
  s: 4,
  q: 4,
  t: 2,
  a: 7,
  z: 0,
};

/**
 * Path data as absolute commands, with quadratic curves and arcs as cubic curves. As SVG
 * renders it, data after an error is dropped and the path up to the error draws.
 */
export function parseSvgPath(d: string): PathCommand[] {
  const out: PathCommand[] = [];
  let i = 0;
  const n = d.length;
  const space = () => {
    while (i < n && /[\s,]/.test(d[i]!)) i += 1;
  };
  const readNumber = (): number | undefined => {
    space();
    const m = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/.exec(d.slice(i, i + 64));
    if (!m) return undefined;
    i += m[0].length;
    const v = Number(m[0]);
    return Number.isFinite(v) ? v : undefined;
  };
  const readFlag = (): boolean | undefined => {
    space();
    const c = d[i];
    if (c !== '0' && c !== '1') return undefined;
    i += 1;
    return c === '1';
  };
  let cx = 0;
  let cy = 0;
  let sx = 0;
  let sy = 0;
  // The last control point, for `S` and `T`, and the command that set it.
  let lastC: [number, number] | undefined;
  let lastQ: [number, number] | undefined;
  let cmd = '';
  let started = false;
  for (;;) {
    space();
    if (i >= n) break;
    if (COMMAND.test(d[i]!)) {
      cmd = d[i]!;
      i += 1;
    } else if (cmd === '' || cmd === 'z' || cmd === 'Z') break;
    const lower = cmd.toLowerCase();
    if (!started && lower !== 'm') break;
    const rel = cmd !== cmd.toUpperCase();
    if (lower === 'z') {
      out.push({ op: 'Z' });
      cx = sx;
      cy = sy;
      lastC = lastQ = undefined;
      continue;
    }
    const args: number[] = [];
    let ok = true;
    for (let k = 0; k < ARGS[lower]!; k += 1) {
      const v = lower === 'a' && (k === 3 || k === 4) ? readFlag() : readNumber();
      if (v === undefined) {
        ok = false;
        break;
      }
      args.push(typeof v === 'boolean' ? (v ? 1 : 0) : v);
    }
    if (!ok) break;
    const ox = rel ? cx : 0;
    const oy = rel ? cy : 0;
    let nextC: [number, number] | undefined;
    let nextQ: [number, number] | undefined;
    switch (lower) {
      case 'm':
        cx = ox + args[0]!;
        cy = oy + args[1]!;
        sx = cx;
        sy = cy;
        out.push({ op: 'M', x: cx, y: cy });
        started = true;
        // Further pairs after a move are lines.
        cmd = rel ? 'l' : 'L';
        break;
      case 'l':
        cx = ox + args[0]!;
        cy = oy + args[1]!;
        out.push({ op: 'L', x: cx, y: cy });
        break;
      case 'h':
        cx = ox + args[0]!;
        out.push({ op: 'L', x: cx, y: cy });
        break;
      case 'v':
        cy = oy + args[0]!;
        out.push({ op: 'L', x: cx, y: cy });
        break;
      case 'c':
      case 's': {
        const [x1, y1] =
          lower === 'c'
            ? [ox + args[0]!, oy + args[1]!]
            : lastC
              ? [2 * cx - lastC[0], 2 * cy - lastC[1]]
              : [cx, cy];
        const o = lower === 'c' ? 2 : 0;
        const x2 = ox + args[o]!;
        const y2 = oy + args[o + 1]!;
        cx = ox + args[o + 2]!;
        cy = oy + args[o + 3]!;
        out.push({ op: 'C', x1, y1, x2, y2, x: cx, y: cy });
        nextC = [x2, y2];
        break;
      }
      case 'q':
      case 't': {
        const [qx, qy] =
          lower === 'q'
            ? [ox + args[0]!, oy + args[1]!]
            : lastQ
              ? [2 * cx - lastQ[0], 2 * cy - lastQ[1]]
              : [cx, cy];
        const o = lower === 'q' ? 2 : 0;
        const x = ox + args[o]!;
        const y = oy + args[o + 1]!;
        out.push({
          op: 'C',
          x1: cx + ((qx - cx) * 2) / 3,
          y1: cy + ((qy - cy) * 2) / 3,
          x2: x + ((qx - x) * 2) / 3,
          y2: y + ((qy - y) * 2) / 3,
          x,
          y,
        });
        cx = x;
        cy = y;
        nextQ = [qx, qy];
        break;
      }
      case 'a': {
        const x = ox + args[5]!;
        const y = oy + args[6]!;
        arcCurves(cx, cy, args[0]!, args[1]!, args[2]!, args[3] === 1, args[4] === 1, x, y, out);
        cx = x;
        cy = y;
        break;
      }
    }
    lastC = nextC;
    lastQ = nextQ;
    if (out.length > MAX_DRAWING_COMMANDS) throw new SvgUnsupported('path too long');
  }
  return out;
}

// --- Basic shapes -----------------------------------------------------------------------

const K = 0.5522847498307936;

function ellipsePath(cx: number, cy: number, rx: number, ry: number): PathCommand[] {
  const kx = rx * K;
  const ky = ry * K;
  return [
    { op: 'M', x: cx + rx, y: cy },
    { op: 'C', x1: cx + rx, y1: cy + ky, x2: cx + kx, y2: cy + ry, x: cx, y: cy + ry },
    { op: 'C', x1: cx - kx, y1: cy + ry, x2: cx - rx, y2: cy + ky, x: cx - rx, y: cy },
    { op: 'C', x1: cx - rx, y1: cy - ky, x2: cx - kx, y2: cy - ry, x: cx, y: cy - ry },
    { op: 'C', x1: cx + kx, y1: cy - ry, x2: cx + rx, y2: cy - ky, x: cx + rx, y: cy },
    { op: 'Z' },
  ];
}

export function rectPath(x: number, y: number, w: number, h: number, rxIn?: number, ryIn?: number) {
  let rx = rxIn !== undefined && rxIn > 0 ? rxIn : undefined;
  let ry = ryIn !== undefined && ryIn > 0 ? ryIn : undefined;
  rx ??= ry ?? 0;
  ry ??= rx;
  rx = Math.min(rx, w / 2);
  ry = Math.min(ry, h / 2);
  if (rx === 0 || ry === 0)
    return [
      { op: 'M', x, y },
      { op: 'L', x: x + w, y },
      { op: 'L', x: x + w, y: y + h },
      { op: 'L', x, y: y + h },
      { op: 'Z' },
    ] satisfies PathCommand[];
  const kx = rx * K;
  const ky = ry * K;
  const r = x + w;
  const b = y + h;
  return [
    { op: 'M', x: x + rx, y },
    { op: 'L', x: r - rx, y },
    { op: 'C', x1: r - rx + kx, y1: y, x2: r, y2: y + ry - ky, x: r, y: y + ry },
    { op: 'L', x: r, y: b - ry },
    { op: 'C', x1: r, y1: b - ry + ky, x2: r - rx + kx, y2: b, x: r - rx, y: b },
    { op: 'L', x: x + rx, y: b },
    { op: 'C', x1: x + rx - kx, y1: b, x2: x, y2: b - ry + ky, x, y: b - ry },
    { op: 'L', x, y: y + ry },
    { op: 'C', x1: x, y1: y + ry - ky, x2: x + rx - kx, y2: y, x: x + rx, y },
    { op: 'Z' },
  ] satisfies PathCommand[];
}

/** The path of a basic shape element, or undefined when it renders nothing. */
export function shapePath(
  local: string,
  a: Readonly<Record<string, string>>
): PathCommand[] | undefined {
  const len = (name: string, fallback = 0) => num(a[name]) ?? fallback;
  switch (local) {
    case 'path':
      return parseSvgPath(a.d ?? '');
    case 'rect': {
      const w = len('width');
      const h = len('height');
      if (!(w > 0 && h > 0)) return undefined;
      return rectPath(len('x'), len('y'), w, h, num(a.rx), num(a.ry));
    }
    case 'circle': {
      const r = len('r');
      return r > 0 ? ellipsePath(len('cx'), len('cy'), r, r) : undefined;
    }
    case 'ellipse': {
      const rx = num(a.rx) ?? num(a.ry) ?? 0;
      const ry = num(a.ry) ?? num(a.rx) ?? 0;
      return rx > 0 && ry > 0 ? ellipsePath(len('cx'), len('cy'), rx, ry) : undefined;
    }
    case 'line':
      return [
        { op: 'M', x: len('x1'), y: len('y1') },
        { op: 'L', x: len('x2'), y: len('y2') },
      ];
    case 'polyline':
    case 'polygon': {
      const pts = numberList(a.points) ?? [];
      const out: PathCommand[] = [];
      for (let k = 0; k + 1 < pts.length; k += 2)
        out.push({ op: k === 0 ? 'M' : 'L', x: pts[k]!, y: pts[k + 1]! });
      if (out.length === 0) return undefined;
      if (local === 'polygon') out.push({ op: 'Z' });
      return out;
    }
  }
  return undefined;
}

export const transformPath = (path: readonly PathCommand[], m: Matrix): PathCommand[] =>
  path.map((c) => {
    const at = (x: number, y: number): [number, number] => [
      m[0] * x + m[2] * y + m[4],
      m[1] * x + m[3] * y + m[5],
    ];
    if (c.op === 'Z') return c;
    const [x, y] = at(c.x, c.y);
    if (c.op === 'C') {
      const [x1, y1] = at(c.x1, c.y1);
      const [x2, y2] = at(c.x2, c.y2);
      return { op: 'C', x1, y1, x2, y2, x, y };
    }
    if (c.op === 'Q') {
      const [x1, y1] = at(c.x1, c.y1);
      return { op: 'Q', x1, y1, x, y };
    }
    return { op: c.op, x, y };
  });
