// Vertical anchoring of laid-out text, shared by text layout and table layout. Internal to the
// layout lane (not exported from `layout/index.ts`).

import type { BodyProps } from '../../cascade/index.ts';
import type { Line, TextLayout } from './text-layout.ts';

/** How far anchoring moves text down inside the inner box (insets excluded). */
export function anchorOffset(
  anchor: BodyProps['anchor'],
  innerHeight: number,
  contentHeight: number
): number {
  if (anchor === 'ctr') return (innerHeight - contentHeight) / 2;
  if (anchor === 'b') return innerHeight - contentHeight;
  return 0;
}

/** Lines moved down by `dy`. Position arrays are shared, not copied. */
export function shiftLines(lines: readonly Line[], dy: number): Line[] {
  return lines.map((l) => ({ ...l, top: l.top + dy, baseline: l.baseline + dy }));
}

/** `vert` values that lay text out in the swapped frame (its width is the box height). */
export const ROTATED_VERT: ReadonlySet<string> = new Set([
  'vert',
  'vert270',
  'eaVert',
  'wordArtVert',
  'mongolianVert',
]);

/**
 * Whether a body's lines do not depend on its box height: not laid out in the swapped frame,
 * and one column (columns flow by height).
 */
export function heightIndependent(body: BodyProps): boolean {
  return body.numCol <= 1 && !ROTATED_VERT.has(body.vert);
}

/**
 * A layout made in a box of one height, moved to a box of the same width at another place and
 * height. Valid only for `heightIndependent` bodies: then only the box and the anchoring
 * change, and the result equals laying the text out again in the new box.
 */
export function placeTextLayout(
  text: TextLayout,
  body: BodyProps,
  rect: { x: number; y: number; width: number; height: number }
): TextLayout {
  const inner = (h: number): number => Math.max(0, h - body.tIns - body.bIns);
  const content = text.contentHeight - body.tIns - body.bIns;
  const dy =
    anchorOffset(body.anchor, inner(rect.height), content) -
    anchorOffset(body.anchor, inner(text.box.height), content);
  return {
    ...text,
    box: { x: rect.x, y: rect.y, width: rect.width, height: rect.height },
    lines: dy === 0 ? text.lines : shiftLines(text.lines, dy),
  };
}
