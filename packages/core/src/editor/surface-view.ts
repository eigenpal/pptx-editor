// The view of the slide surface: the veil that dims what lies outside the slide, and the scroll
// that keeps the caret in view. Editing chrome only: thumbnails, the slide show, and export
// paint the slide alone, clipped to its box.

import { SVG_NS } from '../output/index.ts';

/** How far the veil reaches past the slide, in slide sizes: farther than any zoom shows. */
const REACH = 50;

/**
 * The veil over the area outside the slide (`.pptx-surface__outside`): the parts of shapes
 * dragged partly off the slide stay visible there, dimmed, as slide editors show them. The
 * slide box itself is a hole in the veil. Slide EMU. @internal
 */
export function outsideVeil(doc: Document, width: number, height: number): Element {
  const path = doc.createElementNS(SVG_NS, 'path');
  const rx = width * REACH;
  const ry = height * REACH;
  path.setAttribute('class', 'pptx-surface__outside');
  path.setAttribute('fill-rule', 'evenodd');
  path.setAttribute(
    'd',
    `M${-rx} ${-ry}H${width + rx}V${height + ry}H${-rx}Z` + `M0 0V${height}H${width}V0Z`
  );
  return path;
}

/** A box in client pixels. */
export interface ClientBox {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

/**
 * Scroll `scroller` the least distance that brings `box` inside its visible area, `margin`
 * pixels in from each edge. Nothing scrolls while the box is in view. @internal
 */
export function scrollBoxIntoView(scroller: HTMLElement, box: ClientBox, margin: number): void {
  const view = scroller.getBoundingClientRect();
  // The margin shrinks in a view too small to keep it on both sides.
  const mx = Math.max(0, Math.min(margin, (view.width - (box.right - box.left)) / 2));
  const my = Math.max(0, Math.min(margin, (view.height - (box.bottom - box.top)) / 2));
  if (box.left < view.left + mx) scroller.scrollLeft -= view.left + mx - box.left;
  else if (box.right > view.right - mx) scroller.scrollLeft += box.right - (view.right - mx);
  if (box.top < view.top + my) scroller.scrollTop -= view.top + my - box.top;
  else if (box.bottom > view.bottom - my) scroller.scrollTop += box.bottom - (view.bottom - my);
}

/** The surface padding in CSS pixels, including the no-stylesheet default. */
export function surfacePadding(root: HTMLElement): { x: number; y: number } {
  const own = root.ownerDocument.defaultView?.getComputedStyle(root);
  const side = (value: string | undefined): number => {
    const px = parseFloat(value ?? '');
    return Number.isFinite(px) ? px : 24;
  };
  return {
    x: side(own?.paddingLeft) + side(own?.paddingRight),
    y: side(own?.paddingTop) + side(own?.paddingBottom),
  };
}
