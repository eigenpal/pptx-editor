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
  'wordArtVertRtl',
  'mongolianVert',
]);

/**
 * A body with its insets in the frame that its text lays out in. Each inset stays on its side
 * of the shape: in `vert` text, whose frame turns 90° clockwise, the line starts at `tIns` and
 * the text top is at `rIns`; `vert270` turns the other way. Reference PDFs of
 * `text-direction` and `vertical-columns`: the lines start 3.6 pt (`tIns`, `bIns`) from the
 * box edge, and the first baseline is 7.2 pt (`lIns`, `rIns`) plus the line's own offset in.
 * DOCX Editor's `cellFlowBox` (`packages/core/src/layout/table-cell-text-direction.ts`) maps
 * the margins of a `btLr` cell, the `vert270` of a table, the same way.
 * @internal
 */
export function frameInsets(body: BodyProps): BodyProps {
  if (!ROTATED_VERT.has(body.vert)) return body;
  const { lIns, tIns, rIns, bIns } = body;
  if (body.vert === 'vert270') return { ...body, lIns: bIns, tIns: lIns, rIns: tIns, bIns: rIns };
  // `wordArtVert` and `mongolianVert` lay out as `wordArtVertRtl` and `vert` with their lines
  // mirrored (`columnLines`), so the first line takes `lIns`.
  if (body.vert === 'wordArtVert' || body.vert === 'mongolianVert')
    return { ...body, lIns: tIns, tIns: lIns, rIns: bIns, bIns: rIns };
  return { ...body, lIns: tIns, tIns: rIns, rIns: bIns, bIns: lIns };
}

/**
 * The advance of each stacked letter (`wordArtVert`, `wordArtVertRtl`), in ems. Measured on the
 * reference PDF of `text-direction` (16 pt) and on `STACK` of `vertical-east-asian` (32 pt
 * bold): 22.76 and 45.6 pt from one letter to the next.
 * @internal
 */
export const STACKED_LETTER_EM = 1.425;

/**
 * The lines of a frame turned 90° clockwise, in the order that `vert` places them. In
 * `wordArtVert` and `mongolianVert` the lines run from the left edge of the shape, so their
 * places mirror in the frame; the glyphs keep their turn. Stacked text (`wordArtVert`,
 * `wordArtVertRtl`) centers each letter in a square cell of its advance at the line's start
 * edge; the line's `baseline` is then the axis through those centers, which the paint reads
 * (ECMA-376 Part 1, 20.1.10.83). Reference PDFs of `text-direction` and
 * `vertical-east-asian`: stacked letters are within 1.4 pt of the reference.
 * @internal
 */
export function columnLines(lines: readonly Line[], vert: string, frameHeight: number): Line[] {
  const stacked = vert === 'wordArtVert' || vert === 'wordArtVertRtl';
  const mirror = vert === 'wordArtVert' || vert === 'mongolianVert';
  if (!stacked && !mirror) return [...lines];
  return lines.map((l) => {
    const top = mirror ? frameHeight - l.top - l.height : l.top;
    if (!stacked) return { ...l, top, baseline: top + l.baseline - l.top };
    let size = 0;
    for (const r of l.runs) size = Math.max(size, r.sizePt);
    const half = (STACKED_LETTER_EM * size * 12_700) / 2;
    return { ...l, top, baseline: mirror ? top + l.height - half : top + half };
  });
}

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
