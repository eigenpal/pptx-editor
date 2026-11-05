// Hit tests of the slide surface: which element is under a slide point, whether a point is on
// a shape's border band (a press there moves the shape), and whether it is on laid-out text (a
// press there places the caret). Slide EMU throughout. Moved out of `surface.ts` (the
// 1000-line cap).

import { apply, invert, type LayoutItem, type SlideLayout } from '../layout/index.ts';
import { textFrameMatrix } from './surface-geometry.ts';
import { isLineItem, scopeItems } from '../automation/commands/placement.ts';

const EMU_PER_PX = 9_525;

type Point = { readonly x: number; readonly y: number };

/** Whether a slide point is within `tolerance` of a shape's first path (curves sampled). */
function nearPath(item: LayoutItem, p: Point, tolerance: number): boolean {
  if (item.kind !== 'shape') return false;
  const local: { x: number; y: number }[] = [];
  let at = { x: 0, y: 0 };
  for (const c of item.geometry.paths[0]?.commands ?? []) {
    if (c.op === 'Z') continue;
    if (c.op === 'C')
      for (let i = 1; i < 8; i += 1) {
        const t = i / 8;
        const u = 1 - t;
        local.push({
          x: u * u * u * at.x + 3 * u * u * t * c.x1 + 3 * u * t * t * c.x2 + t * t * t * c.x,
          y: u * u * u * at.y + 3 * u * u * t * c.y1 + 3 * u * t * t * c.y2 + t * t * t * c.y,
        });
      }
    local.push({ x: c.x, y: c.y });
    at = { x: c.x, y: c.y };
  }
  const pts = local.map((q) => apply(item.matrix, q.x, q.y));
  for (let i = 1; i < pts.length; i += 1) {
    const a = pts[i - 1];
    const b = pts[i];
    const dx = b.x - a.x;
    const dy = b.y - a.y;
    const len2 = dx * dx + dy * dy;
    const t = len2 ? Math.max(0, Math.min(1, ((p.x - a.x) * dx + (p.y - a.y) * dy) / len2)) : 0;
    if (Math.hypot(p.x - (a.x + t * dx), p.y - (a.y + t * dy)) <= tolerance) return true;
  }
  return false;
}

/**
 * The editable element under a slide point among a scope's items (the slide's top level, or
 * an entered group's children). Geometric: a click anywhere inside a shape's bounds selects
 * it, not only on painted ink. Topmost wins. @internal
 */
export function hitItem(
  layout: SlideLayout | undefined,
  p: Point,
  zoom: number,
  scope?: string
): string | undefined {
  if (!layout) return undefined;
  const tolerance = (6 * EMU_PER_PX) / zoom;
  const items = scopeItems(layout.items, scope);
  for (let i = items.length - 1; i >= 0; i -= 1) {
    const item = items[i];
    if (item.element.layer !== 'slide') continue;
    // A line is hit near its path, not anywhere in its bounding box.
    if (isLineItem(item)) {
      if (nearPath(item, p, tolerance)) return item.element.nodeId;
      continue;
    }
    const local = apply(invert(item.matrix), p.x, p.y);
    const { cx, cy } = item.element.xfrm;
    const tx = cx < tolerance ? tolerance : 0;
    const ty = cy < tolerance ? tolerance : 0;
    if (local.x >= -tx && local.x <= cx + tx && local.y >= -ty && local.y <= cy + ty)
      return item.element.nodeId;
  }
  return undefined;
}

/**
 * Placeholder types whose content is not text (ECMA-376 Part 1, 19.7.10 ST_PlaceholderType):
 * their prompt asks for a picture, table, chart, diagram, or media object.
 */
const NON_TEXT_PLACEHOLDERS: ReadonlySet<string> = new Set([
  'pic',
  'clipArt',
  'tbl',
  'chart',
  'dgm',
  'media',
  'sldImg',
]);

/**
 * Whether an empty placeholder takes a press in its box as a caret placement. A text
 * placeholder (19.3.1.36 ph) does: it shows a text prompt. A text box does too (`inText`).
 * Another empty shape does not, so a press anywhere in it moves it.
 */
const takesText = (type: string | undefined): boolean =>
  type !== undefined && !NON_TEXT_PLACEHOLDERS.has(type);

/** The border band width: 8 screen pixels, in slide EMU. */
const bandOf = (zoom: number): number => (8 * EMU_PER_PX) / zoom;

/**
 * Whether a slide point is on laid-out text of a shape: inside a line box, between the line's
 * first and last caret stops. A press there edits the text, even inside the border band of a
 * text box with no inset. @internal
 */
export function onGlyph(item: LayoutItem, p: Point): boolean {
  if (item.kind !== 'shape' || !item.text) return false;
  const m = textFrameMatrix(item);
  if (!m) return false;
  const q = apply(invert(m), p.x, p.y);
  return item.text.lines.some((line) => {
    if (line.end <= line.start || q.y < line.top || q.y > line.top + line.height) return false;
    let lo = Infinity;
    let hi = -Infinity;
    for (const x of line.caretX) {
      lo = Math.min(lo, x);
      hi = Math.max(hi, x);
    }
    return q.x >= lo && q.x <= hi;
  });
}

/**
 * Whether a slide point is on a shape's text: the shape has text (or is an empty text box or
 * text placeholder) and the point is on a glyph, or inside its box away from the border band where
 * a drag moves the shape. @internal
 */
export function inText(item: LayoutItem, p: Point, zoom: number): boolean {
  if (item.kind !== 'shape' || !item.text) return false;
  const el = item.element;
  const hasRuns = el.text?.paragraphs.some((para) => para.runs.length > 0);
  if (!hasRuns && !el.textBox && !takesText(el.placeholder?.type)) return false;
  if (onGlyph(item, p)) return true;
  const band = bandOf(zoom);
  const local = apply(invert(item.matrix), p.x, p.y);
  const { cx, cy } = item.element.xfrm;
  return local.x > band && local.x < cx - band && local.y > band && local.y < cy - band;
}

/**
 * Whether a slide point is on a shape's border band, inside or just outside its box. A point on
 * laid-out text is not: the text takes the press. @internal
 */
export function onBorder(item: LayoutItem, p: Point, zoom: number): boolean {
  const band = bandOf(zoom);
  const local = apply(invert(item.matrix), p.x, p.y);
  const { cx, cy } = item.element.xfrm;
  const inOuter =
    local.x >= -band && local.x <= cx + band && local.y >= -band && local.y <= cy + band;
  const inInner = local.x > band && local.x < cx - band && local.y > band && local.y < cy - band;
  return inOuter && !inInner && !onGlyph(item, p);
}
