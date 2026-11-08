// After DOCX Editor `packages/core/src/layout/visual-caret-navigation.ts` and
// `packages/core/src/editor/surface-selection-collapse.ts` at commit adf0e178.
//
// The caret on a line in visual order (`Line.bidi`). Offsets stay logical; a stop is one place
// on screen: an offset and the side of a direction boundary it shows on. Every visual edge of
// a line gets one stop, so the arrow keys step by one place, and every offset that a cluster
// boundary allows gets one where the edges suffice. Lines without `bidi` read in source order
// and never come here. DOM-free.

import type { CaretAffinity, TextPoint } from '../contracts/index.ts';
import type { LayoutItem, Line } from '../layout/index.ts';
import { paragraphUnits } from '../automation/commands/text-units.ts';

type ShapeItem = Extract<LayoutItem, { kind: 'shape' }>;

/** One caret place on a line, in visual order. */
export interface VisualCaretStop {
  readonly offset: number;
  /** Set when the place is the trailing edge of the character before the offset. */
  readonly affinity?: 'upstream';
  readonly x: number;
}

/** Edges closer than this (EMU) are one place. */
const SAME = 1;

/** The x of a caret at `offset` on `line`, on the side `affinity` names. */
export function caretXOn(line: Line, offset: number, affinity?: CaretAffinity): number | undefined {
  const i = offset - line.start;
  const trailing = line.bidi?.trailingX;
  if (affinity === 'upstream' && trailing && i > 0 && i <= trailing.length) return trailing[i - 1];
  return line.caretX[Math.max(0, Math.min(i, line.caretX.length - 1))];
}

/**
 * The caret stops of a bidi line from left to right. `allowed` says which offsets a caret may
 * rest at (cluster boundaries). `wraps` is true when the next line continues the paragraph: its
 * start offset then belongs to that line.
 *
 * Each allowed offset has up to two places: the leading edge of its unit (downstream) and the
 * trailing edge of the unit before it (upstream). The places are matched to the visual edges
 * in source order, the downstream place first, with an augmenting path where an offset finds
 * no free place, so each edge shows one offset and each offset shows once where it can. An edge that no unmatched offset reaches
 * repeats an offset there.
 */
export function bidiLineStops(
  line: Line,
  allowed: (offset: number) => boolean,
  wraps: boolean
): VisualCaretStop[] {
  const n = line.end - line.start;
  const trailing = line.bidi?.trailingX ?? [];
  const options: VisualCaretStop[][] = [];
  for (let k = 0; k <= n; k += 1) {
    const offset = line.start + k;
    if (!allowed(offset) || (k === n && wraps && n > 0)) continue;
    const down: VisualCaretStop = { offset, x: line.caretX[k] ?? 0 };
    const up = k > 0 && k < n ? trailing[k - 1] : undefined;
    options.push(
      up !== undefined && Math.abs(up - down.x) > SAME
        ? [down, { offset, affinity: 'upstream', x: up }]
        : [down]
    );
  }
  // The visual edges: the distinct places, from left to right.
  const xs = options.flat().map((c) => c.x);
  xs.sort((a, b) => a - b);
  const edges: number[] = [];
  for (const x of xs) if (edges.length === 0 || x - edges[edges.length - 1]! > SAME) edges.push(x);
  const edgeOf = (x: number): number => {
    let lo = 0;
    let hi = edges.length - 1;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (edges[mid]! < x - SAME) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  };
  const places = options.map((list) => list.map((c) => edgeOf(c.x)));
  const owner = new Array<number>(edges.length).fill(-1);
  const take = (o: number, seen: Uint8Array): boolean => {
    for (const e of places[o]!) {
      if (seen[e]) continue;
      seen[e] = 1;
      if (owner[e] === -1 || take(owner[e]!, seen)) {
        owner[e] = o;
        return true;
      }
    }
    return false;
  };
  // A free place first, in source order; an augmenting path only for an offset with none, so
  // an offset keeps its downstream place unless a later offset has no other.
  for (let o = 0; o < options.length; o += 1) {
    const free = places[o]!.find((e) => owner[e] === -1);
    if (free !== undefined) owner[free] = o;
    else take(o, new Uint8Array(edges.length));
  }
  const stops: VisualCaretStop[] = [];
  edges.forEach((_, e) => {
    const own = owner[e]!;
    const at: VisualCaretStop[] = [];
    options.forEach((list, o) => {
      if (own !== -1 && o !== own) return;
      list.forEach((c, j) => places[o]![j] === e && at.push(c));
    });
    const stop = at.find((c) => !c.affinity) ?? at[0];
    if (stop) stops.push(stop);
  });
  return stops;
}

/** The laid-out line that holds the caret: a line's end offset belongs to the next line. */
export function lineIndexAt(lines: readonly Line[], paragraph: number, offset: number): number {
  return lines.findIndex(
    (l, i) =>
      l.paragraph === paragraph &&
      offset >= l.start &&
      (offset < l.end || lines[i + 1]?.paragraph !== paragraph)
  );
}

const stopsOfLine = new WeakMap<Line, VisualCaretStop[]>();

/**
 * The caret stops of a bidi line of a shape's text, from left to right: at the cluster
 * boundaries of its paragraph, without the start of the next line of the paragraph.
 */
export function lineStopsOf(item: ShapeItem, line: Line): VisualCaretStop[] {
  let stops = stopsOfLine.get(line);
  if (!stops) {
    const lines = item.text?.lines ?? [];
    const next = lines[lines.indexOf(line) + 1];
    const units = paragraphUnits(item.element.text?.paragraphs[line.paragraph]);
    const allowed = new Set(units.stops);
    stops = bidiLineStops(line, (o) => allowed.has(o), next?.paragraph === line.paragraph);
    stopsOfLine.set(line, stops);
  }
  return stops;
}

/**
 * The index of the stop where the caret of `point` shows at `x`: the stop of its offset there,
 * else any stop there (an offset two places share), else the stop of its offset, else the
 * nearest stop.
 */
export function stopIndex(
  stops: readonly VisualCaretStop[],
  point: Pick<TextPoint, 'offset'>,
  x: number
): number {
  let best = -1;
  let bestRank = Infinity;
  let bestDist = Infinity;
  stops.forEach((s, i) => {
    const d = Math.abs(s.x - x);
    const own = s.offset === point.offset;
    const rank = d <= SAME ? (own ? 0 : 1) : own ? 2 : 3;
    if (rank < bestRank || (rank === bestRank && d < bestDist)) {
      best = i;
      bestRank = rank;
      bestDist = d;
    }
  });
  return best;
}

/** The stop nearest `x`. */
export function nearestStop(
  stops: readonly VisualCaretStop[],
  x: number
): VisualCaretStop | undefined {
  let best: VisualCaretStop | undefined;
  for (const s of stops) if (!best || Math.abs(s.x - x) < Math.abs(best.x - x)) best = s;
  return best;
}

/** A stop as a text point of `paragraph`. */
export function stopPoint(paragraph: number, stop: VisualCaretStop): TextPoint {
  return stop.affinity
    ? { paragraph, offset: stop.offset, affinity: stop.affinity }
    : { paragraph, offset: stop.offset };
}

/**
 * The x ranges that the units `[from, to)` of a bidi line cover, from left to right: one band
 * for each stretch that stands together on screen. A logical range that crosses a direction
 * boundary covers several.
 */
export function bidiSelectionBands(line: Line, from: number, to: number): [number, number][] {
  const trailing = line.bidi?.trailingX ?? [];
  const cells: [number, number][] = [];
  for (let i = Math.max(from, line.start); i < Math.min(to, line.end); i += 1) {
    const a = line.caretX[i - line.start];
    const b = trailing[i - line.start];
    if (a === undefined || b === undefined || Math.abs(a - b) <= 0) continue;
    cells.push([Math.min(a, b), Math.max(a, b)]);
  }
  cells.sort((p, q) => p[0] - q[0]);
  const bands: [number, number][] = [];
  for (const c of cells) {
    const last = bands[bands.length - 1];
    if (last && c[0] <= last[1] + SAME) last[1] = Math.max(last[1], c[1]);
    else bands.push([c[0], c[1]]);
  }
  return bands;
}

/**
 * Where a plain Left (-1) or Right (1) collapses the range `[from, to]` of one bidi line: the
 * outermost edge of its selected bands on the key's side, as the offset and side of the
 * selected unit there. Undefined when the range has no band on the line.
 */
export function collapseOnLine(
  line: Line,
  from: number,
  to: number,
  direction: -1 | 1
): Pick<TextPoint, 'offset' | 'affinity'> | undefined {
  const bands = bidiSelectionBands(line, from, to);
  const band = direction === -1 ? bands[0] : bands[bands.length - 1];
  if (!band) return undefined;
  const x = direction === -1 ? band[0] : band[1];
  const near = (v: number | undefined): boolean => v !== undefined && Math.abs(v - x) <= SAME;
  const trailing = line.bidi?.trailingX ?? [];
  for (let o = Math.max(from, line.start); o < Math.min(to, line.end); o += 1) {
    const i = o - line.start;
    if (near(line.caretX[i])) return { offset: o };
    if (near(trailing[i]))
      return near(line.caretX[i + 1]) ? { offset: o + 1 } : { offset: o + 1, affinity: 'upstream' };
  }
  return undefined;
}
