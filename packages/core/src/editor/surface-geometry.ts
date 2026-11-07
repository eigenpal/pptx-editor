// Text geometry on a laid-out shape: the text frame matrix, the caret point under a position,
// and the caret line of a point. Shared by the surface, the selection painters, and the slide
// show. Moved out of `surface.ts` (the 1000-line cap).

import type { TextPoint } from '../contracts/index.ts';
import {
  apply,
  invert,
  multiply,
  type LayoutItem,
  type Line,
  type Matrix,
} from '../layout/index.ts';

/** The text frame → slide matrix of a shape item, including the body rotation. */
export function textFrameMatrix(item: Extract<LayoutItem, { kind: 'shape' }>): Matrix | undefined {
  if (!item.text || !item.textMatrix) return undefined;
  const b = item.text.box;
  const r = (item.text.rotation * Math.PI) / 180;
  const c = Math.cos(r);
  const s = Math.sin(r);
  const cx = b.width / 2;
  const cy = b.height / 2;
  // translate(box) · rotate(rotation around center)
  const rot: Matrix = [c, s, -s, c, cx - c * cx + s * cy, cy - s * cx - c * cy];
  return multiply(multiply(item.textMatrix, [1, 0, 0, 1, b.x, b.y]), rot);
}

/** The caret point nearest to a slide-space position inside a shape's text. */
export function pointAt(item: LayoutItem, sx: number, sy: number): TextPoint | undefined {
  if (item.kind !== 'shape' || !item.text) return undefined;
  const m = textFrameMatrix(item);
  if (!m) return undefined;
  const p = apply(invert(m), sx, sy);
  const lines = item.text.lines;
  if (lines.length === 0) return { paragraph: 0, offset: 0 };
  let line = lines[lines.length - 1];
  for (const l of lines)
    if (p.y < l.top + l.height) {
      line = l;
      break;
    }
  let best = 0;
  let dist = Infinity;
  line.caretX.forEach((x, i) => {
    const d = Math.abs(x - p.x);
    if (d < dist) {
      dist = d;
      best = i;
    }
  });
  let offset = line.start + best;
  // A line that ends at a break or a wrap: the caret may not sit after its last unit.
  if (
    best === line.caretX.length - 1 &&
    line.end > line.start &&
    line.paragraph === lines[lines.indexOf(line) + 1]?.paragraph
  )
    offset = Math.max(line.start, line.end - 1);
  return { paragraph: line.paragraph, offset };
}

/** Caret geometry in slide EMU: top point, bottom point. */
/** The caret's line in slide EMU, and the color of the text it sits in, when there is one. */
export interface CaretLine {
  readonly x1: number;
  readonly y1: number;
  readonly x2: number;
  readonly y2: number;
  /** A CSS color: the text's own color, so the caret shows wherever the text does. */
  readonly color?: string;
}

/**
 * The color of the text at `offset` in `line`: the run that holds the character before the
 * caret, else the first run after it, as typing there takes that run's format. A solid fill
 * gives its color, a gradient its first stop. No color for a line with no text run, for no
 * fill, or for a fill that is nearly transparent.
 */
function textColorAt(line: Line, offset: number): string | undefined {
  const runs = line.runs.filter((r) => r.kind === 'text' && r.text.length > 0);
  const before = runs.filter((r) => r.start < offset);
  const run = before[before.length - 1] ?? runs[0];
  if (!run) return undefined;
  const fill = run.props.fill;
  const color =
    fill.type === 'solid'
      ? fill.color
      : fill.type === 'linear' || fill.type === 'path'
        ? fill.stops[0]?.color
        : undefined;
  if (!color || color.a < 0.35) return undefined;
  return `rgb(${color.r} ${color.g} ${color.b})`;
}

export function caretAt(item: LayoutItem, point: TextPoint): CaretLine | undefined {
  if (item.kind !== 'shape' || !item.text) return undefined;
  const m = textFrameMatrix(item);
  if (!m) return undefined;
  const lines = item.text.lines.filter((l) => l.paragraph === point.paragraph);
  let x = item.text.box.width / 2;
  let top = 0;
  let h = 12 * 12_700;
  let color: string | undefined;
  const line =
    lines.find(
      (l, i) => point.offset >= l.start && (point.offset < l.end || i === lines.length - 1)
    ) ?? lines[0];
  if (line) {
    x = line.caretX[Math.max(0, Math.min(point.offset - line.start, line.caretX.length - 1))] ?? x;
    h = Math.max(line.fontHeight, 1);
    top = line.baseline - h * 0.8;
    color = textColorAt(line, point.offset);
  } else if (item.text.lines.length === 0) {
    // An empty body: center vertically per anchor, at the left inset.
    top = item.text.box.height / 2 - h / 2;
  }
  const a = apply(m, x, top);
  const b = apply(m, x, top + h);
  return color
    ? { x1: a.x, y1: a.y, x2: b.x, y2: b.y, color }
    : { x1: a.x, y1: a.y, x2: b.x, y2: b.y };
}

/**
 * Whether a caret point sits in linked text: the character before it (the first one at the
 * start of a paragraph) carries `a:hlinkClick`, the rule the selection's format reads.
 */
export function linkedAt(item: LayoutItem, point: TextPoint): boolean {
  if (item.kind !== 'shape') return false;
  const p = item.element.text?.paragraphs[point.paragraph];
  if (!p) return false;
  const at = Math.max(0, point.offset - 1);
  let pos = 0;
  for (const r of p.runs) {
    const len = r.kind === 'break' ? 1 : r.text.length;
    if (r.kind !== 'break' && at >= pos && at < pos + Math.max(1, len))
      return !!(r.props.hyperlink || r.props.hyperlinkAction);
    pos += len;
  }
  return false;
}

/** The slide point `(sx, sy)` is on linked text of `item` (`linkedAt` at the point there). */
export function linkedAtPoint(item: LayoutItem, sx: number, sy: number): boolean {
  const at = pointAt(item, sx, sy);
  return !!at && linkedAt(item, at);
}
