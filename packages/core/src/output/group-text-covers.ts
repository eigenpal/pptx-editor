// The text under later bodies in a soft-edged group. Under the soft edge of a group, the bodies
// of its shapes paint together through the soft edge, and the text of its shapes paints once over
// them. So that a later body still covers the text of an earlier shape, that text paints clipped
// to the outside of each later body over it. The SVG paint and the PDF export clip with the same
// paths.

import type { ResolvedEffects, ResolvedFill } from '../cascade/index.ts';
import type { LayoutItem, PathCommand } from '../layout/index.ts';
import { blurSigma, boundsOf, grow, type Bounds } from './effect-transforms.ts';
import { boxCorners, textCorners } from './item-corners.ts';
import { mapCommands } from './path-data.ts';

/**
 * The most path commands in the clips of the text of one soft-edged group (`groupTextCovers`):
 * each text that a body covers takes its own copy of the body's outline, so the clips of a group
 * stay this size whatever the outlines and the number of texts.
 */
export const MAX_GROUP_TEXT_CLIP_COMMANDS = 65_536;

/** The clip paths of the text of a soft-edged group's shapes. @internal */
export interface GroupTextCovers {
  /**
   * For each shape whose text a later body covers, the paths that clip its text, in slide space.
   * Each path is a rectangle around the text and the outline of one later body, filled even-odd:
   * inside the rectangle and outside the body. The text paints inside all of them.
   */
  readonly clips: ReadonlyMap<LayoutItem, readonly (readonly PathCommand[])[]>;
  /**
   * False when the clips of the text of some shape need more commands than
   * `MAX_GROUP_TEXT_CLIP_COMMANDS` leaves: that text paints over the later bodies.
   */
  readonly complete: boolean;
}

type BoxItem = Extract<LayoutItem, { kind: 'shape' } | { kind: 'picture' }>;

/** A body that can cover text: its slide-space bounds, its outlines, and their commands. */
interface Body {
  readonly bounds: Bounds;
  readonly outlines: () => readonly PathCommand[][];
  /** The number of outlines, and of commands in them. */
  readonly count: number;
  readonly commands: number;
}

/** The commands of `boundsPath`. */
const AROUND = 5;

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
 * Whether a fill surely hides what lies under it: an opaque color, a gradient of opaque stops, or
 * a pattern of two opaque colors. A picture can hold clear pixels, and a group fill is the group's.
 */
function opaque(fill: ResolvedFill): boolean {
  switch (fill.type) {
    case 'solid':
      return fill.color.a >= 1;
    case 'linear':
    case 'path':
      return fill.stops.length > 0 && fill.stops.every((s) => s.color.a >= 1);
    case 'pattern':
      return fill.fg.a >= 1 && fill.bg.a >= 1;
    default:
      return false;
  }
}

/**
 * Whether effects fade a body: a soft edge or a blur with a radius, as the paints draw them (the
 * SVG blur filters any positive `blurSigma`). A radius of 0 leaves the body opaque.
 */
function fades(fx: ResolvedEffects | undefined): boolean {
  return (fx?.softEdge?.radius ?? 0) > 0 || (fx?.blur ? blurSigma(fx.blur.radius) > 0 : false);
}

/**
 * The body of a shape that surely hides the text under it: the filled paths of its geometry, in
 * slide space, under an opaque fill. A picture, a table, a shape whose fill lets text show
 * through, a shape whose soft edge or blur fades its body, and a 3-D shape, whose face can paint
 * away from its flat outline, cover nothing: the text paints over them unclipped.
 */
function bodyOf(item: BoxItem): Body | undefined {
  if (item.kind !== 'shape' || !opaque(item.element.fill)) return undefined;
  const fx = item.element.effects;
  if (fades(fx) || fx?.threeD) return undefined;
  const filled = item.geometry.paths.filter((p) => p.fill !== 'none');
  if (filled.length === 0) return undefined;
  let outlines: PathCommand[][] | undefined;
  return {
    bounds: boundsOf(boxCorners(item)),
    // Mapped once, when a text first takes it.
    outlines: () => (outlines ??= filled.map((p) => mapCommands(p.commands, item.matrix))),
    count: filled.length,
    commands: filled.reduce((n, p) => n + p.commands.length, 0),
  };
}

const overlaps = (a: Bounds, b: Bounds): boolean =>
  a.x < b.r && b.x < a.r && a.y < b.b && b.y < a.b;

/**
 * The clips of the text of `group`'s shapes that a later body in the group covers: an opaque
 * shape (`bodyOf`), in no nested group with a soft edge or a blur, that paints after the text
 * and whose box meets its text frame. The text of a
 * 3-D shape paints in its face's turned space and takes no clip. A group with effects holds at
 * most `MAX_GROUP_EFFECT_ELEMENTS` items, so the check of each text against each later body stays
 * bounded, and `MAX_GROUP_TEXT_CLIP_COMMANDS` bounds the clips. @internal
 */
export function groupTextCovers(group: Extract<LayoutItem, { kind: 'group' }>): GroupTextCovers {
  // The bodies and texts in paint order: a shape's text paints over its own body.
  const order: { body?: Body; text?: { item: LayoutItem; bounds: Bounds } }[] = [];
  // `faded`: a group between `group` and the item fades it with a soft edge or a blur, so its
  // bodies cover nothing.
  const walk = (item: LayoutItem, faded: boolean): void => {
    if (item.element.hidden) return;
    switch (item.kind) {
      case 'group': {
        const fading = faded || fades(item.element.effects);
        for (const child of item.children) walk(child, fading);
        return;
      }
      case 'frame': {
        // A table's cells can be clear: it covers nothing.
        if (item.poster) walk(item.poster, faded);
        for (const child of item.drawing ?? item.chart ?? []) walk(child, faded);
        return;
      }
      default: {
        const body = faded ? undefined : bodyOf(item);
        if (body) order.push({ body });
        const text = textCorners(item);
        if (text.length > 0 && !item.element.effects?.threeD)
          order.push({ text: { item, bounds: boundsOf(text) } });
      }
    }
  };
  for (const child of group.children) walk(child, false);
  const clips = new Map<LayoutItem, PathCommand[][]>();
  let left = MAX_GROUP_TEXT_CLIP_COMMANDS;
  let complete = true;
  const later: Body[] = [];
  for (let i = order.length - 1; i >= 0; i -= 1) {
    const { body, text } = order[i]!;
    if (body) later.push(body);
    if (!text) continue;
    const over = later.filter((b) => overlaps(b.bounds, text.bounds));
    if (over.length === 0) continue;
    // Each clip is the rectangle around the text and one outline: counted before any is mapped.
    const outlineCount = over.reduce((n, b) => n + b.count, 0);
    const commands = over.reduce((n, b) => n + b.commands, 0) + AROUND * outlineCount;
    if (commands > left) {
      complete = false;
      continue;
    }
    left -= commands;
    const outlines = over.flatMap((b) => b.outlines());
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
