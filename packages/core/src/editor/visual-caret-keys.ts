// Adapted from DOCX Editor `packages/core/src/layout/visual-caret-navigation.ts`
// (`horizontalCaretStep`, `visualLineEdge`, `visualWordBoundary`) and
// `packages/core/src/editor/surface-selection-collapse.ts` at commit adf0e178.
//
// The caret keys on a line in visual order (`Line.bidi`). Left and Right step one place on
// screen; past the line's edge they go to the next line in reading order: a line of the same
// paragraph is entered at its edge on the key's side, another paragraph at its start or end.
// With Ctrl or Alt they walk the places by word, and with Cmd they go to the line's visual edge on
// their side. Home and End stay at the logical line edges (`text-input.ts`); `lineEdgePlace`
// only picks the side the caret shows on. Every result is a logical offset with the side it
// shows on. A line without `bidi` keeps the source-order keys of `text-input.ts`.

import type { LayoutItem, Line } from '../layout/index.ts';
import type { TextPoint } from '../contracts/index.ts';
import type { ParagraphUnits } from '../automation/commands/text-units.ts';
import {
  caretXOn,
  collapseOnLine,
  lineIndexAt,
  lineStopsOf,
  nearestStop,
  stopIndex,
  stopPoint,
} from './visual-caret.ts';

type ShapeItem = Extract<LayoutItem, { kind: 'shape' }>;

const linesOf = (item: ShapeItem): readonly Line[] => item.text?.lines ?? [];

/** The same offset on the same side. */
export function samePlace(a: TextPoint, b: TextPoint): boolean {
  return (
    a.paragraph === b.paragraph &&
    a.offset === b.offset &&
    (a.affinity ?? 'downstream') === (b.affinity ?? 'downstream')
  );
}

/** One place left (-1) or right (1) from `point` on a bidi line; null past the text body edge. */
function visualStep(item: ShapeItem, point: TextPoint, direction: -1 | 1): TextPoint | null {
  const lines = linesOf(item);
  const li = lineIndexAt(lines, point.paragraph, point.offset);
  const line = lines[li];
  if (!line?.bidi) return null;
  const stops = lineStopsOf(item, line);
  const i = stopIndex(stops, point, caretXOn(line, point.offset, point.affinity) ?? 0);
  const target = stops[i + direction];
  if (target) return stopPoint(line.paragraph, target);
  // Past the line's edge, in reading order: a right-to-left line reads on to its left.
  const logical = (direction * (line.bidi.paragraphRtl ? -1 : 1)) as -1 | 1;
  const next = lines[li + logical];
  if (!next) return null;
  if (next.paragraph === line.paragraph && next.bidi) {
    const into = lineStopsOf(item, next);
    const entry = direction === 1 ? into[0] : into[into.length - 1];
    return entry ? stopPoint(next.paragraph, entry) : null;
  }
  return { paragraph: next.paragraph, offset: logical === 1 ? next.start : next.end };
}

/** Whether `a` and `b` are neighboring caret stops of the paragraph (no jump between them). */
function adjacent(u: ParagraphUnits, a: number, b: number): boolean {
  return Math.abs(u.stops.indexOf(a) - u.stops.indexOf(b)) === 1;
}

/** Whether the unit from `offset` belongs to a word. */
function inWord(u: ParagraphUnits, offset: number): boolean {
  return u.words.some((w) => w.word && w.start <= offset && offset < w.end);
}

/**
 * A word step on screen (`visualWordBoundary`): past the gap, then through one word, stopping
 * where the word ends or where the walk jumps across a direction boundary after a word. Within
 * the caret's paragraph; null when the walk does not move.
 */
function visualWordStep(
  item: ShapeItem,
  start: TextPoint,
  direction: -1 | 1,
  u: ParagraphUnits
): TextPoint | null {
  let current = start;
  let consumedWord = false;
  for (let count = 0; count <= 2 * u.length + 2; count += 1) {
    const next = visualStep(item, current, direction);
    if (!next || next.paragraph !== start.paragraph) break;
    if (next.offset !== current.offset) {
      const from = Math.min(next.offset, current.offset);
      if (!adjacent(u, next.offset, current.offset)) {
        if (consumedWord) break;
      } else {
        const word = inWord(u, from);
        if (consumedWord && !word) break;
        consumedWord ||= word;
      }
    }
    current = next;
  }
  return current.offset === start.offset ? null : current;
}

/**
 * The caret after Left or Right on a line in visual order, or null when the caret is not on one
 * (the source-order keys apply). `word` is Ctrl or Alt with the arrow. `edge` is Cmd alone with
 * it: the caret goes to the line's visual edge on the arrow's side, whatever the paragraph's
 * direction. `logicalEnd` gives the offset where End stops on the line (before a break or the
 * spaces of a wrap), so an edge step does not land in the spaces a wrap hangs.
 */
export function bidiKeyTarget(
  item: ShapeItem,
  focus: TextPoint,
  pressed: string,
  mods: { readonly word: boolean; readonly edge: boolean },
  u: ParagraphUnits,
  logicalEnd: () => number | undefined
): TextPoint | null {
  if (pressed !== 'ArrowLeft' && pressed !== 'ArrowRight') return null;
  const lines = linesOf(item);
  const line = lines[lineIndexAt(lines, focus.paragraph, focus.offset)];
  if (!line?.bidi) return null;
  const direction = pressed === 'ArrowLeft' ? -1 : 1;
  if (mods.edge) {
    const stops = lineStopsOf(item, line);
    const limit = logicalEnd();
    const kept = limit === undefined ? stops : stops.filter((s) => s.offset <= limit);
    const pool = kept.length > 0 ? kept : stops;
    const stop = direction === -1 ? pool[0] : pool[pool.length - 1];
    return stop ? stopPoint(line.paragraph, stop) : null;
  }
  // At a paragraph edge a word step goes on like a plain arrow, so the key is never dead.
  return (
    (mods.word ? visualWordStep(item, focus, direction, u) : null) ??
    visualStep(item, focus, direction) ??
    focus
  );
}

/**
 * Home or End at the logical line edge `point` (DOCX Editor `lineEdgeTarget`): on a line in
 * visual order, End shows the caret after the line's last character (upstream), where a
 * direction boundary puts the next character elsewhere. Home shows it before the first.
 */
export function lineEdgePlace(item: ShapeItem, point: TextPoint, end: boolean): TextPoint {
  if (!end || point.offset === 0) return point;
  const lines = linesOf(item);
  const line = lines[lineIndexAt(lines, point.paragraph, point.offset)];
  if (!line?.bidi || point.offset <= line.start) return point;
  const up = caretXOn(line, point.offset, 'upstream');
  return up !== caretXOn(line, point.offset) ? { ...point, affinity: 'upstream' } : point;
}

/** Up or Down onto bidi line `target`: the place nearest `x`; null for another line. */
export function bidiVerticalTarget(item: ShapeItem, target: Line, x: number): TextPoint | null {
  if (!target.bidi) return null;
  const stop = nearestStop(lineStopsOf(item, target), x);
  return stop ? stopPoint(target.paragraph, stop) : null;
}

/**
 * Where a plain Left (-1) or Right (1) collapses the range `from`–`to` when its start is on a
 * line in visual order, or null when it is not (the logical end applies). On one line: the
 * outermost edge of the selected bands on the key's side. Over several lines: the end that the
 * paragraph's reading direction puts on that side.
 */
export function bidiCollapse(
  item: ShapeItem,
  from: TextPoint,
  to: TextPoint,
  direction: -1 | 1
): TextPoint | null {
  const lines = linesOf(item);
  const li = lineIndexAt(lines, from.paragraph, from.offset);
  const line = lines[li];
  if (!line?.bidi) return null;
  const toLine = lineIndexAt(lines, to.paragraph, to.offset);
  const sameLine =
    from.paragraph === to.paragraph &&
    (toLine === li || (to.offset === line.end && lines[li + 1]?.paragraph === line.paragraph));
  if (sameLine) {
    const edge = collapseOnLine(line, from.offset, to.offset, direction);
    if (edge) return { paragraph: line.paragraph, ...edge };
  }
  return direction * (line.bidi.paragraphRtl ? -1 : 1) === -1 ? from : to;
}
