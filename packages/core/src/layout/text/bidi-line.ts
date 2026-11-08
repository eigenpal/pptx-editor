// The positions of a bidi line: its units placed in visual order (UAX #9 L2), the pieces paint
// draws, and the caret position of each source offset. DOM-free; lengths are EMU.
//
// Trailing spaces follow DOCX Editor `packages/core/src/layout/bidi-trailing-spaces.ts` at
// commit adf0e178: spaces that reordering puts past the line's visible text hang outside it,
// and alignment ignores them; spaces between two pieces of visible text take room like text.
// UAX #9 L1 gives the spaces that end a line the paragraph level, so they hang at the line's
// trailing edge.

import { bidiAlgorithm } from './bidi.ts';
import { lineVisualOrder, type ParagraphBidi } from './rtl-paragraph.ts';

/** A run of text at its place: its source text and its cell edges (`GlyphRun`). */
interface BidiRun {
  readonly text: string;
  readonly x: readonly number[];
  readonly rtl?: true;
}

/** One stretch of a line that paint draws in one direction, from one segment. */
export interface BidiPiece {
  /** The segment of the units. */
  readonly seg: number;
  /** Line-relative source range of the units, `[from, to)`. */
  readonly from: number;
  readonly to: number;
  /** The units read right to left: unit `from + i` draws in cell `to - from - 1 - i`. */
  readonly rtl: boolean;
  /** Cell boundaries from left to right, `to - from + 1` values, relative to `contentLeft`. */
  readonly x: readonly number[];
}

/** A line laid out in visual order. */
export interface BidiLine {
  /** The left edge of each unit, line-relative, from the visual start of the line. */
  readonly left: Float64Array;
  /** Whether each unit reads right to left (an odd level after L1). */
  readonly rtl: Uint8Array;
  /** The visual extent of the line's visible content (the units before `contentEnd`). */
  readonly contentLeft: number;
  readonly contentRight: number;
  /** The pieces in visual order. */
  readonly pieces: readonly BidiPiece[];
}

/**
 * Place the units `[start, end)` of a paragraph in visual order. `advance(i)` and `segment(i)`
 * take a paragraph unit index; `codes` holds the paragraph's UTF-16 units. Units from
 * `contentEnd` on are the line's trailing spaces and breaks.
 */
export function placeBidiLine(
  bidi: ParagraphBidi,
  codes: ArrayLike<number>,
  start: number,
  end: number,
  contentEnd: number,
  advance: (i: number) => number,
  segment: (i: number) => number
): BidiLine {
  const count = end - start;
  const local = Array.from({ length: count }, (_, i) => codes[start + i]!);
  const { order, levels } = lineVisualOrder(bidi, local, start, end);
  const left = new Float64Array(count);
  const rtl = new Uint8Array(count);
  let x = 0;
  let contentLeft = Infinity;
  let contentRight = -Infinity;
  for (const index of order) {
    left[index] = x;
    rtl[index] = levels[index]! % 2;
    const width = advance(start + index);
    if (start + index < contentEnd) {
      contentLeft = Math.min(contentLeft, x);
      contentRight = Math.max(contentRight, x + width);
    }
    x += width;
  }
  if (!Number.isFinite(contentLeft)) {
    // Only spaces and breaks: the line starts at the paragraph's leading edge.
    contentLeft = bidi.rtl ? x : 0;
    contentRight = contentLeft;
  }
  const pieces: BidiPiece[] = [];
  let cur: { seg: number; rtl: boolean; lo: number; hi: number } | null = null;
  const flush = (): void => {
    if (!cur) return;
    const xs: number[] = [];
    if (cur.rtl) for (let i = cur.hi; i >= cur.lo; i--) xs.push(left[i]! - contentLeft);
    else for (let i = cur.lo; i <= cur.hi; i++) xs.push(left[i]! - contentLeft);
    const last = cur.rtl ? cur.lo : cur.hi;
    xs.push(left[last]! + advance(start + last) - contentLeft);
    pieces.push({ seg: cur.seg, from: cur.lo, to: cur.hi + 1, rtl: cur.rtl, x: xs });
    cur = null;
  };
  for (const index of order) {
    const seg = segment(start + index);
    const unitRtl = rtl[index] === 1;
    if (
      cur &&
      cur.seg === seg &&
      cur.rtl === unitRtl &&
      (unitRtl ? index === cur.lo - 1 : index === cur.hi + 1)
    ) {
      if (unitRtl) cur.lo = index;
      else cur.hi = index;
      continue;
    }
    flush();
    cur = { seg, rtl: unitRtl, lo: index, hi: index };
  }
  flush();
  return { left, rtl, contentLeft, contentRight, pieces };
}

/**
 * The caret position of each source offset of a placed line, `count + 1` values relative to
 * `contentLeft`: the leading edge of the unit at the offset (its left edge, or its right edge
 * when it reads right to left), and after the last unit, that unit's trailing edge. On a line
 * that mixes directions, one offset can stand at two places; this gives the place of the unit
 * that follows it.
 */
export function bidiCaretX(
  line: BidiLine,
  start: number,
  count: number,
  advance: (i: number) => number
): number[] {
  const out: number[] = [];
  const edge = (i: number, trailing: boolean): number => {
    const leftEdge = line.left[i]! - line.contentLeft;
    const rightEdge = leftEdge + advance(start + i);
    return (line.rtl[i] === 1) !== trailing ? rightEdge : leftEdge;
  };
  for (let i = 0; i < count; i++) out.push(edge(i, false));
  out.push(count > 0 ? edge(count - 1, true) : 0);
  return out;
}

/**
 * A bullet or number of a right-to-left paragraph, laid out left to right at its place, in the
 * paragraph's direction: one run for each stretch of one direction. In the `rtl-text` reference
 * PDF (slide 3), `1.` reads `.1`, as UAX #9 orders it in a right-to-left paragraph.
 */
export function markerRuns<T extends BidiRun>(marker: T): T[] {
  const n = marker.text.length;
  if (n === 0) return [marker];
  const codes = Array.from({ length: n }, (_, i) => marker.text.charCodeAt(i));
  const levels = bidiAlgorithm.getEmbeddingLevels(marker.text, 'rtl').levels;
  const advance = (i: number): number => marker.x[i + 1] - marker.x[i];
  const line = placeBidiLine({ rtl: true, levels }, codes, 0, n, n, advance, () => 0);
  return line.pieces.map((piece) => ({
    ...marker,
    text: marker.text.slice(piece.from, piece.to),
    x: piece.x.map((v) => v + marker.x[0]),
    ...(piece.rtl ? { rtl: true as const } : {}),
  }));
}

/**
 * The pending bullet as one run (`Line.pendingBullet`): the editor draws it dimmed, as a hint.
 * A bullet that reads in two directions keeps its left-to-right order there.
 */
export function wholeMarker<T extends BidiRun>(placed: T, rtl: boolean): T {
  if (!rtl) return placed;
  const runs = markerRuns(placed);
  return runs.length === 1 ? runs[0] : placed;
}
