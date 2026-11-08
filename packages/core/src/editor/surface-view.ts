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

/**
 * The zoom that fits slide `size` (EMU) in `host`, inside the host's own padding (an open find
 * bar reserves its top) and the surface padding of `root` (`--doc-slide-padding`, 24px a side).
 * @internal
 */
export function fitZoom(
  host: HTMLElement,
  root: HTMLElement,
  size: { readonly width: number; readonly height: number }
): number {
  const pad = host.ownerDocument.defaultView?.getComputedStyle(host);
  const sum = (a?: string, b?: string): number =>
    (parseFloat(a ?? '') || 0) + (parseFloat(b ?? '') || 0);
  const own = surfacePadding(root);
  const w = Math.max(100, host.clientWidth - sum(pad?.paddingLeft, pad?.paddingRight) - own.x);
  const h = Math.max(100, host.clientHeight - sum(pad?.paddingTop, pad?.paddingBottom) - own.y);
  return Math.max(0.1, Math.min(w / (size.width / 9_525), h / (size.height / 9_525)));
}

/** A caret line in slide EMU: the caret's two ends. */
interface CaretLine {
  readonly x1: number;
  readonly y1: number;
  readonly x2: number;
  readonly y2: number;
}

/**
 * Scrolls the caret (the focus end of a range) into view when it moves by a key: typing at a
 * high zoom, or an arrow key. A pointer press does not scroll: the caret is under it.
 */
export class CaretScroll {
  /** The caret the view last followed. */
  #key = '';
  /** The caret scroll that waits for the next frame. */
  #pending: { caret: CaretLine; scale: number } | null = null;
  readonly #host: HTMLElement;
  readonly #stage: HTMLElement;
  readonly #root: HTMLElement;

  constructor(host: HTMLElement, stage: HTMLElement, root: HTMLElement) {
    this.#host = host;
    this.#stage = stage;
    this.#root = root;
  }

  /** Follow the caret `key` at `caret`; `pressed` is true while a pointer press is down. */
  follow(key: string, caret: CaretLine, scale: number, pressed: boolean): void {
    if (key === this.#key) return;
    this.#key = key;
    if (pressed) return;
    // Read the stage position at the next frame, where the browser lays out the page anyway. A
    // read here would force a layout of the whole slide on every keystroke.
    const waiting = !!this.#pending;
    this.#pending = { caret, scale };
    if (waiting) return;
    const view = this.#host.ownerDocument.defaultView;
    const scroll = () => {
      const next = this.#pending;
      this.#pending = null;
      if (!next || !this.#root.isConnected) return;
      const c = next.caret;
      const st = this.#stage.getBoundingClientRect();
      const box = {
        left: st.left + Math.min(c.x1, c.x2) * next.scale,
        right: st.left + Math.max(c.x1, c.x2) * next.scale,
        top: st.top + Math.min(c.y1, c.y2) * next.scale,
        bottom: st.top + Math.max(c.y1, c.y2) * next.scale,
      };
      scrollBoxIntoView(this.#host, box, 32);
    };
    if (view?.requestAnimationFrame) view.requestAnimationFrame(scroll);
    else scroll();
  }
}
