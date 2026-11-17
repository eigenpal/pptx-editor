// The fill of a shape that uses the slide background (`p:sp/@useBgFill`, ECMA-376 Part 1,
// 19.3.1.43). The shape shows the slide background as if the slide showed through it, so a
// gradient or a picture keeps its place on the slide instead of stretching over the shape.

import { boolAttr, child, NS, type Element } from '../store/index.ts';
import { fillElementOf, type RelRect, type ResolvedFill } from './paint-props.ts';
import type { Xfrm } from './types.ts';

/** The resolved slide background and the part whose relationships its picture names. */
export interface SlideBackground {
  readonly fill: ResolvedFill;
  readonly part: string;
  readonly width: number;
  readonly height: number;
}

/**
 * Whether a `p:sp` fills with the slide background. The attribute replaces the style fill and
 * the fills that a placeholder inherits. A fill in the shape's own `p:spPr` still applies, so a
 * fill set in the editor shows without removing the attribute.
 */
export function fillsWithBackground(el: Element): boolean {
  return (
    el.ns === NS.p && boolAttr(el, 'useBgFill') === true && !fillElementOf(child(el, 'p:spPr'))
  );
}

const NO_INSET: RelRect = { l: 0, t: 0, r: 0, b: 0 };

/**
 * `inset` (fractions of the slide) as insets from the shape box `xfrm`. Negative insets extend
 * the box, as `a:tileRect` and `a:fillRect` allow. Undefined when the shape box is empty.
 */
function slideInsetToShape(
  inset: RelRect,
  xfrm: Xfrm,
  width: number,
  height: number
): RelRect | undefined {
  if (!(xfrm.cx > 0 && xfrm.cy > 0)) return undefined;
  const left = inset.l * width;
  const top = inset.t * height;
  const right = width - inset.r * width;
  const bottom = height - inset.b * height;
  return {
    l: (left - xfrm.x) / xfrm.cx,
    t: (top - xfrm.y) / xfrm.cy,
    r: (xfrm.x + xfrm.cx - right) / xfrm.cx,
    b: (xfrm.y + xfrm.cy - bottom) / xfrm.cy,
  };
}

/**
 * The fill of a shape with `useBgFill`. A solid color or a pattern applies as it is. A gradient
 * or a stretched picture applies over the slide box, as insets from the shape box. That needs
 * the shape box in slide space, unturned and unflipped (`slideSpace`), and, for a picture, a
 * background in the shape's own part, whose relationships the picture names. Otherwise the
 * shape has no fill, so the background shows through it where nothing sits under the shape.
 */
export function backgroundFillOf(
  bg: SlideBackground,
  xfrm: Xfrm,
  partName: string,
  slideSpace: boolean
): ResolvedFill {
  const fill = bg.fill;
  if (fill.type === 'solid' || fill.type === 'pattern' || fill.type === 'none') return fill;
  const mappable = slideSpace && xfrm.rot === 0 && !xfrm.flipH && !xfrm.flipV;
  if (mappable && (fill.type === 'linear' || fill.type === 'path')) {
    const tileRect = slideInsetToShape(fill.tileRect ?? NO_INSET, xfrm, bg.width, bg.height);
    if (tileRect) return { ...fill, tileRect };
  }
  if (mappable && fill.type === 'picture' && fill.mode === 'stretch' && bg.part === partName) {
    const fillRect = slideInsetToShape(fill.fillRect, xfrm, bg.width, bg.height);
    if (fillRect) return { ...fill, fillRect };
  }
  return { type: 'none' };
}
