// The shapes that Office Math draws itself, as filled outlines: fraction and radical rules, the
// radical sign, stretched delimiters, horizontal braces, and arrows. Lengths are EMU, with y up
// from the math baseline. The proportions are measured on the reference PDF of the `equations`
// fixture at 18 pt, where the delimiters and the radical sign are paths, not glyphs. Internal to
// the layout lane.

import type { PathCommand } from '@eigenpal/ooxml-core';

type Cmd = PathCommand;

const M = (x: number, y: number): Cmd => ({ op: 'M', x, y });
const L = (x: number, y: number): Cmd => ({ op: 'L', x, y });
const Q = (x1: number, y1: number, x: number, y: number): Cmd => ({ op: 'Q', x1, y1, x, y });
const Z: Cmd = { op: 'Z' };

/** A filled rectangle. */
export function rect(x: number, y: number, w: number, h: number): Cmd[] {
  return [M(x, y), L(x + w, y), L(x + w, y + h), L(x, y + h), Z];
}

/**
 * The radical sign from `x` (its left edge) between `bottom` and `top` (the top of its stroke),
 * `width` wide, for a font size of `em`. The tick on the left keeps its size; the long stroke
 * rises to `top`, `width` from the left edge.
 */
export function radicalSign(
  x: number,
  bottom: number,
  top: number,
  width: number,
  em: number
): Cmd[] {
  const t = 0.055 * em;
  const px = (v: number): number => x + v * em;
  const py = (v: number): number => bottom + v * em;
  return [
    M(px(0), py(0.434)),
    L(px(0.015), py(0.398)),
    L(px(0.098), py(0.434)),
    L(px(0.3), py(0)),
    L(px(0.347), py(0)),
    L(x + width - t * 1.6, top - t * 0.45),
    L(x + width, top - t * 0.45),
    L(x + width, top),
    L(x + width - t * 0.82, top),
    L(px(0.332), py(0.129)),
    L(px(0.16), py(0.508)),
    Z,
  ];
}

/** The width that a stretched delimiter takes for `ch`, at font size `em`, or undefined. */
export function delimiterWidth(ch: string, em: number, height: number): number | undefined {
  switch (ch) {
    case '(':
    case ')':
      return Math.min(0.224 * em + 0.02 * height, 0.5 * em);
    case '[':
    case ']':
    case '⌊':
    case '⌋':
    case '⌈':
    case '⌉':
    case '⟦':
    case '⟧':
      return 0.218 * em;
    case '{':
    case '}':
      return Math.min(0.3 * em + 0.02 * height, 0.55 * em);
    case '⟨':
    case '⟩':
      return Math.min(0.24 * em + 0.04 * height, 0.5 * em);
    case '|':
      return 0.075 * em;
    case '‖':
      return 0.25 * em;
    default:
      return undefined;
  }
}

/**
 * A stretched delimiter (`m:d`), from `x` to `x + w`, between `bottom` and `top`. `ch` must be one
 * that `delimiterWidth` knows.
 */
export function delimiter(
  ch: string,
  x: number,
  bottom: number,
  top: number,
  w: number,
  em: number
): Cmd[] {
  const h = top - bottom;
  const mid = (top + bottom) / 2;
  const stem = 0.08 * em;
  const serif = 0.038 * em;
  const flip = (cmds: Cmd[]): Cmd[] =>
    cmds.map((c) => {
      if (c.op === 'Z') return c;
      if (c.op === 'Q') return { ...c, x1: 2 * x + w - c.x1, x: 2 * x + w - c.x };
      if (c.op === 'C')
        return { ...c, x1: 2 * x + w - c.x1, x2: 2 * x + w - c.x2, x: 2 * x + w - c.x };
      return { ...c, x: 2 * x + w - c.x };
    });
  switch (ch) {
    case '(':
    case ')': {
      // A crescent: the outer curve reaches the left edge at the middle, the inner one is a stem
      // thinner there, and the tips meet at the right edge.
      const tip = x + w;
      const cmds = [
        M(tip, top),
        Q(x - w * 0.02, mid + h * 0.25, x, mid),
        Q(x - w * 0.02, mid - h * 0.25, tip, bottom),
        L(tip - stem * 0.2, bottom + stem * 0.45),
        Q(x + stem * 1.3, mid - h * 0.22, x + stem, mid),
        Q(x + stem * 1.3, mid + h * 0.22, tip - stem * 0.2, top - stem * 0.45),
        Z,
      ];
      return ch === '(' ? cmds : flip(cmds);
    }
    case '[':
    case ']':
    case '⌊':
    case '⌋':
    case '⌈':
    case '⌉': {
      const topSerif = ch === '[' || ch === ']' || ch === '⌈' || ch === '⌉';
      const bottomSerif = ch === '[' || ch === ']' || ch === '⌊' || ch === '⌋';
      const cmds: Cmd[] = [M(x, bottom), L(x, top)];
      if (topSerif) cmds.push(L(x + w, top), L(x + w, top - serif), L(x + stem, top - serif));
      else cmds.push(L(x + stem, top));
      if (bottomSerif)
        cmds.push(L(x + stem, bottom + serif), L(x + w, bottom + serif), L(x + w, bottom));
      else cmds.push(L(x + stem, bottom));
      cmds.push(Z);
      return ch === '[' || ch === '⌊' || ch === '⌈' ? cmds : flip(cmds);
    }
    case '⟦':
    case '⟧': {
      const cmds = [
        ...delimiter('[', x, bottom, top, w, em),
        ...rect(x + stem * 2.2, bottom + serif, stem * 0.7, h - 2 * serif),
      ];
      return ch === '⟦' ? cmds : flip(cmds);
    }
    case '{':
    case '}': {
      // Two arcs that meet in a point at the middle of the left edge.
      const right = x + w;
      const neck = x + w * 0.5;
      const cmds = [
        M(right, top),
        Q(neck - stem * 0.5, top, neck - stem * 0.5, top - h * 0.12),
        L(neck - stem * 0.5, mid + h * 0.1),
        Q(neck - stem * 0.5, mid, x, mid),
        Q(neck - stem * 0.5, mid, neck - stem * 0.5, mid - h * 0.1),
        L(neck - stem * 0.5, bottom + h * 0.12),
        Q(neck - stem * 0.5, bottom, right, bottom),
        L(right, bottom + serif),
        Q(neck + stem * 0.5, bottom + serif, neck + stem * 0.5, bottom + h * 0.12),
        L(neck + stem * 0.5, mid - h * 0.1),
        Q(neck + stem * 0.5, mid - stem, x + stem, mid),
        Q(neck + stem * 0.5, mid + stem, neck + stem * 0.5, mid + h * 0.1),
        L(neck + stem * 0.5, top - h * 0.12),
        Q(neck + stem * 0.5, top - serif, right, top - serif),
        Z,
      ];
      return ch === '{' ? cmds : flip(cmds);
    }
    case '⟨':
    case '⟩': {
      const cmds = [
        M(x + w, top),
        L(x, mid),
        L(x + w, bottom),
        L(x + w + stem * 0.2, bottom + stem * 0.6),
        L(x + stem * 1.1, mid),
        L(x + w + stem * 0.2, top - stem * 0.6),
        Z,
      ];
      return ch === '⟨' ? cmds : flip(cmds);
    }
    case '|':
      return rect(x, bottom, w, h);
    case '‖':
      return [
        ...rect(x, bottom, stem * 0.9, h),
        ...rect(x + w - stem * 0.9, bottom, stem * 0.9, h),
      ];
    default:
      return [];
  }
}

/**
 * A horizontal brace or bracket (`m:groupChr`) from `x0` to `x1`: `⏟` and `⎵` open upward under
 * the base, `⏞` and `⎴` open downward over it. `y` is the edge next to the base.
 */
export function horizontalGroup(ch: string, x0: number, x1: number, y: number, em: number): Cmd[] {
  const below = ch === '⏟' || ch === '⎵' || ch === '⏡';
  const s = below ? -1 : 1;
  const depth = 0.22 * em;
  const t = 0.055 * em;
  const far = y + s * depth;
  if (ch === '⎴' || ch === '⎵' || ch === '⏠' || ch === '⏡')
    return [
      M(x0, y),
      L(x0, far),
      L(x1, far),
      L(x1, y),
      L(x1 - t, y),
      L(x1 - t, far - s * t),
      L(x0 + t, far - s * t),
      L(x0 + t, y),
      Z,
    ];
  // A brace: two arcs from the ends to a point in the middle that points away from the base.
  const midX = (x0 + x1) / 2;
  const shoulder = y + s * depth * 0.55;
  const tip = far;
  const r = Math.min((x1 - x0) / 6, 0.3 * em);
  return [
    M(x0, y),
    Q(x0, shoulder, x0 + r, shoulder),
    L(midX - r, shoulder),
    Q(midX, shoulder, midX, tip),
    Q(midX, shoulder, midX + r, shoulder),
    L(x1 - r, shoulder),
    Q(x1, shoulder, x1, y),
    L(x1 - t, y),
    Q(x1 - t, shoulder - s * t, x1 - r, shoulder - s * t),
    L(midX + r, shoulder - s * t),
    Q(midX + t, shoulder - s * t, midX, tip - s * t * 1.5),
    Q(midX - t, shoulder - s * t, midX - r, shoulder - s * t),
    L(x0 + r, shoulder - s * t),
    Q(x0 + t, shoulder - s * t, x0 + t, y),
    Z,
  ];
}

/** A right (or left) arrow from `x0` to `x1` on `y`, for a vector accent or a group arrow. */
export function arrow(x0: number, x1: number, y: number, em: number, left = false): Cmd[] {
  const t = 0.05 * em;
  const head = Math.min(0.2 * em, (x1 - x0) / 2);
  const hh = 0.11 * em;
  if (left)
    return [
      M(x0, y),
      L(x0 + head, y + hh),
      L(x0 + head, y + t / 2),
      L(x1, y + t / 2),
      L(x1, y - t / 2),
      L(x0 + head, y - t / 2),
      L(x0 + head, y - hh),
      Z,
    ];
  return [
    M(x1, y),
    L(x1 - head, y + hh),
    L(x1 - head, y + t / 2),
    L(x0, y + t / 2),
    L(x0, y - t / 2),
    L(x1 - head, y - t / 2),
    L(x1 - head, y - hh),
    Z,
  ];
}

/** A stroke between two points, `t` wide, as a filled outline (the slash of a skewed fraction). */
export function stroke(x0: number, y0: number, x1: number, y1: number, t: number): Cmd[] {
  const len = Math.hypot(x1 - x0, y1 - y0) || 1;
  const nx = (-(y1 - y0) / len) * (t / 2);
  const ny = ((x1 - x0) / len) * (t / 2);
  return [M(x0 + nx, y0 + ny), L(x1 + nx, y1 + ny), L(x1 - nx, y1 - ny), L(x0 - nx, y0 - ny), Z];
}

/** Commands moved by `dx`, `dy`. */
export function moveCommands(cmds: readonly Cmd[], dx: number, dy: number): Cmd[] {
  if (dx === 0 && dy === 0) return [...cmds];
  return cmds.map((c) => {
    if (c.op === 'Z') return c;
    if (c.op === 'Q') return { ...c, x1: c.x1 + dx, y1: c.y1 + dy, x: c.x + dx, y: c.y + dy };
    if (c.op === 'C')
      return {
        ...c,
        x1: c.x1 + dx,
        y1: c.y1 + dy,
        x2: c.x2 + dx,
        y2: c.y2 + dy,
        x: c.x + dx,
        y: c.y + dy,
      };
    return { ...c, x: c.x + dx, y: c.y + dy };
  });
}
