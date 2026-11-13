// The text under later bodies in a soft-edged group. Under the soft edge of a group, the bodies
// of its shapes paint together through the soft edge, and the text of its shapes paints once over
// them. So that a later body still covers the text of an earlier shape, that text paints clipped
// to the outside of each later body over it. The SVG paint and the PDF export clip with the same
// paths.

import type { LayoutItem, PathCommand } from '../layout/index.ts';
import { boundsOf, cornersUnder, grow, type Bounds } from './effect-transforms.ts';
import { boxCorners, textCorners } from './item-corners.ts';
import { mapCommands } from './path-data.ts';

/** The most clip paths that the text of one soft-edged group takes (`groupTextCovers`). */
export const MAX_GROUP_TEXT_CLIPS = 256;

/** The clip paths of the text of a soft-edged group's shapes. @internal */
export interface GroupTextCovers {
  /**
   * For each shape whose text a later body covers, the paths that clip its text, in slide space.
   * Each path is a rectangle around the text and the outline of one later body, filled even-odd:
   * inside the rectangle and outside the body. The text paints inside all of them.
   */
  readonly clips: ReadonlyMap<LayoutItem, readonly (readonly PathCommand[])[]>;
  /**
   * False when the text of some shape needs more clips than `MAX_GROUP_TEXT_CLIPS` leaves: that
   * text paints over the later bodies.
   */
  readonly complete: boolean;
}

type BoxItem = Extract<LayoutItem, { kind: 'shape' } | { kind: 'picture' }>;

/** A body that can cover text: its slide-space bounds and outlines. */
interface Body {
  readonly bounds: Bounds;
  readonly outlines: () => PathCommand[][];
}

/** A closed path around bounds. */
function boundsPath(b: Bounds): PathCommand[] {
  return [
    { op: 'M', x: b.x, y: b.y },
    { op: 'L', x: b.r, y: b.y },
    { op: 'L', x: b.r, y: b.b },
    { op: 'L', x: b.x, y: b.b },
    { op: 'Z' },
  ];
}

/**
 * The body of a shape or picture: the filled paths of its geometry, in slide space. A shape with
 * no fill covers nothing: its outline alone is left out.
 */
function bodyOf(item: BoxItem): Body | undefined {
  if (item.kind === 'shape' && item.element.fill.type === 'none') return undefined;
  const filled = item.geometry.paths.filter((p) => p.fill !== 'none');
  if (filled.length === 0) return undefined;
  return {
    bounds: boundsOf(boxCorners(item)),
    outlines: () => filled.map((p) => mapCommands(p.commands, item.matrix)),
  };
}

const overlaps = (a: Bounds, b: Bounds): boolean =>
  a.x < b.r && b.x < a.r && a.y < b.b && b.y < a.b;

/**
 * The clips of the text of `group`'s shapes that a later body in the group covers: a shape,
 * picture, or table that paints after the text and whose box meets its text frame. The text of a
 * 3-D shape paints in its face's turned space and takes no clip. A group with effects holds at
 * most `MAX_GROUP_EFFECT_ELEMENTS` items, so the check of each text against each later body stays
 * bounded, and `MAX_GROUP_TEXT_CLIPS` bounds the clips. @internal
 */
export function groupTextCovers(group: Extract<LayoutItem, { kind: 'group' }>): GroupTextCovers {
  // The bodies and texts in paint order: a shape's text paints over its own body.
  const order: { body?: Body; text?: { item: LayoutItem; bounds: Bounds } }[] = [];
  const walk = (item: LayoutItem): void => {
    if (item.element.hidden) return;
    switch (item.kind) {
      case 'group':
        for (const child of item.children) walk(child);
        return;
      case 'frame': {
        if (item.poster) walk(item.poster);
        for (const child of item.drawing ?? item.chart ?? []) walk(child);
        if (!item.table) return;
        const w = Math.max(item.element.xfrm.cx, item.table.width);
        const h = Math.max(item.element.xfrm.cy, item.table.height);
        const bounds = boundsOf(cornersUnder(item.matrix, w, h));
        order.push({ body: { bounds, outlines: () => [boundsPath(bounds)] } });
        return;
      }
      default: {
        const body = bodyOf(item);
        if (body) order.push({ body });
        const text = textCorners(item);
        if (text.length > 0 && !item.element.effects?.threeD)
          order.push({ text: { item, bounds: boundsOf(text) } });
      }
    }
  };
  for (const child of group.children) walk(child);
  const clips = new Map<LayoutItem, PathCommand[][]>();
  let left = MAX_GROUP_TEXT_CLIPS;
  let complete = true;
  const later: Body[] = [];
  for (let i = order.length - 1; i >= 0; i -= 1) {
    const { body, text } = order[i]!;
    if (body) later.push(body);
    if (!text) continue;
    const over = later.filter((b) => overlaps(b.bounds, text.bounds));
    if (over.length === 0) continue;
    const outlines = over.flatMap((b) => b.outlines());
    if (outlines.length > left) {
      complete = false;
      continue;
    }
    left -= outlines.length;
    // Around the text, grown by its larger side: text that overflows its frame stays inside.
    const b = text.bounds;
    const around = boundsPath(grow(b, Math.max(b.r - b.x, b.b - b.y)));
    clips.set(
      text.item,
      outlines.map((o) => [...around, ...o])
    );
  }
  return { clips, complete };
}
