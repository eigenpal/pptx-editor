// The visual order of a laid-out line for the editor (`Line.bidi`): where each unit stands and
// which way it reads, so the caret, the arrow keys, and the selection work on screen while
// offsets stay logical. DOM-free; lengths are EMU.

import type { BidiLine } from './bidi-line.ts';

/** The visual order of a line's units (`Line.bidi`). @public */
export interface LineBidi {
  /** The paragraph reads right to left (`a:pPr/@rtl`). */
  readonly paragraphRtl: boolean;
  /** Whether each unit, `end - start` values from `start`, reads right to left. */
  readonly rtl: readonly boolean[];
  /**
   * The trailing edge of each unit: its right edge, or its left edge when it reads right to
   * left. `Line.caretX` holds the leading edges.
   */
  readonly trailingX: readonly number[];
}

/**
 * The trailing edges and directions of the units of a placed bidi line whose visible content
 * starts at `left`. `advance(i)` takes a paragraph unit index.
 */
export function lineBidiOf(
  line: BidiLine,
  left: number,
  start: number,
  advance: (i: number) => number,
  paragraphRtl: boolean
): LineBidi {
  const rtl: boolean[] = [];
  const trailingX: number[] = [];
  for (let i = 0; i < line.left.length; i += 1) {
    const unitRtl = line.rtl[i] === 1;
    const x = line.left[i]! - line.contentLeft + left;
    rtl.push(unitRtl);
    trailingX.push(unitRtl ? x : x + advance(start + i));
  }
  return { paragraphRtl, rtl, trailingX };
}

/** A line's visual order moved right by `dx` (a later column). */
export function shiftLineBidi(bidi: LineBidi, dx: number): LineBidi {
  return { ...bidi, trailingX: bidi.trailingX.map((v) => v + dx) };
}
