// Smart guides and snapping, shown while a shape moves or resizes: the
// moving box snaps to the edges and centers of the other shapes in its scope, to the slide's
// edges and center, and to equal spacing between shapes. Pure geometry in slide EMU, DOM-free.
// Each axis snaps on its own, to the nearest candidate within the tolerance.

/** An axis-aligned box in slide EMU: left, top, right, bottom. */
export interface SnapBounds {
  readonly x: number;
  readonly y: number;
  readonly r: number;
  readonly b: number;
}

/** A line of a box on one axis: its start edge, its center, or its end edge. */
export type SnapLine = 'start' | 'center' | 'end';

/** The axis a value is measured on: `x` (left to right) or `y` (top to bottom). */
export type SnapAxis = 'x' | 'y';

/** The lines of the moving box that may snap, per axis. */
export interface SnapLines {
  readonly x: readonly SnapLine[];
  readonly y: readonly SnapLine[];
}

/** Every line on both axes: a move. */
export const ALL_LINES: SnapLines = {
  x: ['start', 'center', 'end'],
  y: ['start', 'center', 'end'],
};

export interface SnapRequest {
  /** The moving box at the pointer position, before snapping. */
  readonly moving: SnapBounds;
  /** The bounding boxes of the other visible shapes in the same scope. */
  readonly targets: readonly SnapBounds[];
  readonly slide: { readonly width: number; readonly height: number };
  /** The snap distance in EMU (a screen-pixel tolerance divided by the zoom). */
  readonly tolerance: number;
  /**
   * The lines that move. A move moves all of them (the default) and also snaps to equal
   * spacing; a resize moves only the dragged edges.
   */
  readonly lines?: SnapLines;
}

/**
 * A guide to draw. `align`: a line across the aligned boxes, perpendicular to `axis` at `at`,
 * from `from` to `to` on the other axis (`slide` when it aligns with the slide). `gap`: one
 * of two or more equal gaps, from `from` to `to` on `axis`, drawn at `at` on the other axis.
 */
export type SnapGuide =
  | {
      readonly kind: 'align';
      readonly axis: SnapAxis;
      readonly at: number;
      readonly from: number;
      readonly to: number;
      readonly slide: boolean;
    }
  | {
      readonly kind: 'gap';
      readonly axis: SnapAxis;
      readonly from: number;
      readonly to: number;
      readonly at: number;
    };

/** The snapped displacement to add to the moving box, and the guides at the snapped place. */
export interface SnapResult {
  readonly dx: number;
  readonly dy: number;
  readonly guides: readonly SnapGuide[];
}

/** Lines closer than this are aligned: it absorbs float error, never a visible distance. */
const EPS = 1;

interface Span {
  readonly lo: number;
  readonly hi: number;
}

const along = (b: SnapBounds, axis: SnapAxis): Span =>
  axis === 'x' ? { lo: b.x, hi: b.r } : { lo: b.y, hi: b.b };
const across = (b: SnapBounds, axis: SnapAxis): Span => along(b, axis === 'x' ? 'y' : 'x');
const lineOf = (s: Span, line: SnapLine): number =>
  line === 'start' ? s.lo : line === 'end' ? s.hi : (s.lo + s.hi) / 2;
const overlaps = (a: Span, b: Span): boolean => a.lo < b.hi && b.lo < a.hi;

/** Edges align with edges (left with left or right), centers with centers. */
const pairs = (line: SnapLine): readonly SnapLine[] =>
  line === 'center' ? ['center'] : ['start', 'end'];

/**
 * Equal-spacing places on one axis for a box of `size` whose cross span is `cross`: between two
 * shapes with equal gaps, or beside a pair at the pair's gap. Each place is the box start with
 * the gaps it makes equal. Only shapes in the same row (overlapping `cross`) count, and no
 * other shape of the row may sit inside a gap.
 */
function spacingPlaces(
  rows: readonly Span[],
  size: number
): { start: number; gaps: [Span, Span][] }[] {
  const clear = (lo: number, hi: number): boolean => !rows.some((s) => s.lo < hi && s.hi > lo);
  const out: { start: number; gaps: [Span, Span][] }[] = [];
  for (const a of rows)
    for (const b of rows) {
      if (a === b || a.hi > b.lo) continue;
      // Between `a` and `b`, with equal gaps on both sides.
      const start = (a.hi + b.lo - size) / 2;
      if (start >= a.hi && start + size <= b.lo && clear(a.hi, start) && clear(start + size, b.lo))
        out.push({
          start,
          gaps: [
            [
              { lo: a.hi, hi: start },
              { lo: start + size, hi: b.lo },
            ],
          ],
        });
      // Beside the pair `a`, `b`, at their gap.
      const gap = b.lo - a.hi;
      if (gap <= 0 || !clear(a.hi, b.lo)) continue;
      const pair: Span = { lo: a.hi, hi: b.lo };
      const after = b.hi + gap;
      if (clear(b.hi, after)) out.push({ start: after, gaps: [[pair, { lo: b.hi, hi: after }]] });
      const before = a.lo - gap - size;
      if (clear(a.lo - gap, a.lo))
        out.push({ start: before, gaps: [[pair, { lo: a.lo - gap, hi: a.lo }]] });
    }
  return out;
}

/** The nearest snap on one axis within the tolerance, or 0. */
function snapAxis(req: SnapRequest, axis: SnapAxis, lines: readonly SnapLine[]): number {
  const m = along(req.moving, axis);
  const extent = axis === 'x' ? req.slide.width : req.slide.height;
  let best = 0;
  let bestDist = Infinity;
  const offer = (d: number): void => {
    if (Math.abs(d) <= req.tolerance && Math.abs(d) < bestDist) {
      best = d;
      bestDist = Math.abs(d);
    }
  };
  for (const line of lines) {
    const at = lineOf(m, line);
    for (const t of req.targets)
      for (const other of pairs(line)) offer(lineOf(along(t, axis), other) - at);
    // The slide's edges and center: each line of the box snaps to its own kind.
    offer(lineOf({ lo: 0, hi: extent }, line) - at);
  }
  if (moves(lines)) {
    const cross = across(req.moving, axis);
    const rows = req.targets
      .filter((t) => overlaps(across(t, axis), cross))
      .map((t) => along(t, axis));
    for (const place of spacingPlaces(rows, m.hi - m.lo)) offer(place.start - m.lo);
  }
  return best;
}

/** Whether the lines are a whole move (equal spacing applies only to a move). */
const moves = (lines: readonly SnapLine[]): boolean =>
  lines.includes('start') && lines.includes('end');

/** The guides on one axis for the moving box at its final place. */
function guidesOn(req: SnapRequest, box: SnapBounds, axis: SnapAxis, lines: readonly SnapLine[]) {
  const m = along(box, axis);
  const cross = across(box, axis);
  const extent = axis === 'x' ? req.slide.width : req.slide.height;
  const crossExtent = axis === 'x' ? req.slide.height : req.slide.width;
  const aligned: { at: number; from: number; to: number; slide: boolean }[] = [];
  const add = (at: number, from: number, to: number, slide: boolean): void => {
    const same = aligned.find((g) => Math.abs(g.at - at) <= EPS);
    if (!same) aligned.push({ at, from, to, slide });
    else {
      same.from = Math.min(same.from, from);
      same.to = Math.max(same.to, to);
      same.slide ||= slide;
    }
  };
  for (const line of lines) {
    const at = lineOf(m, line);
    for (const t of req.targets) {
      const c = across(t, axis);
      for (const other of pairs(line))
        if (Math.abs(lineOf(along(t, axis), other) - at) <= EPS)
          add(at, Math.min(cross.lo, c.lo), Math.max(cross.hi, c.hi), false);
    }
    if (Math.abs(lineOf({ lo: 0, hi: extent }, line) - at) <= EPS) add(at, 0, crossExtent, true);
  }
  const guides: SnapGuide[] = aligned.map((g) => ({ kind: 'align', axis, ...g }));
  if (!moves(lines)) return guides;
  const row = req.targets.filter((t) => overlaps(across(t, axis), cross));
  const rows = row.map((t) => along(t, axis));
  // A gap is drawn across the middle of the boxes that bound it.
  const middle = (gap: Span): number => {
    const sides = [box, ...row].filter((t) => {
      const s = along(t, axis);
      return Math.abs(s.hi - gap.lo) <= EPS || Math.abs(s.lo - gap.hi) <= EPS;
    });
    let lo = -Infinity;
    let hi = Infinity;
    for (const t of sides) {
      lo = Math.max(lo, across(t, axis).lo);
      hi = Math.min(hi, across(t, axis).hi);
    }
    return lo <= hi ? (lo + hi) / 2 : (cross.lo + cross.hi) / 2;
  };
  const seen = new Set<string>();
  for (const place of spacingPlaces(rows, m.hi - m.lo)) {
    if (Math.abs(place.start - m.lo) > EPS) continue;
    for (const pair of place.gaps)
      for (const gap of pair) {
        const key = `${Math.round(gap.lo)}:${Math.round(gap.hi)}`;
        if (seen.has(key)) continue;
        seen.add(key);
        guides.push({ kind: 'gap', axis, from: gap.lo, to: gap.hi, at: middle(gap) });
      }
  }
  return guides;
}

/**
 * Snap a moving box: the nearest alignment or equal-spacing place within the tolerance, on
 * each axis independently, and the guides to draw at the snapped place.
 */
export function snapBox(req: SnapRequest): SnapResult {
  const lines = req.lines ?? ALL_LINES;
  const dx = lines.x.length ? snapAxis(req, 'x', lines.x) : 0;
  const dy = lines.y.length ? snapAxis(req, 'y', lines.y) : 0;
  const m = req.moving;
  const box = { x: m.x + dx, y: m.y + dy, r: m.r + dx, b: m.b + dy };
  return {
    dx,
    dy,
    guides: [...guidesOn(req, box, 'x', lines.x), ...guidesOn(req, box, 'y', lines.y)],
  };
}

/** A box moved by a displacement. */
export function moveBounds(b: SnapBounds, dx: number, dy: number): SnapBounds {
  return { x: b.x + dx, y: b.y + dy, r: b.r + dx, b: b.b + dy };
}

/** A box whose `lines` (start or end edges) moved by a displacement: a resize. */
export function resizeBounds(b: SnapBounds, lines: SnapLines, dx: number, dy: number): SnapBounds {
  return {
    x: b.x + (lines.x.includes('start') ? dx : 0),
    r: b.r + (lines.x.includes('end') ? dx : 0),
    y: b.y + (lines.y.includes('start') ? dy : 0),
    b: b.b + (lines.y.includes('end') ? dy : 0),
  };
}
