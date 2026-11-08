// The cost of a large slide on the editing surface (surface.ts): the slide paints over several
// frames, the pointer's cursor sits on a layer over the slide, and the slide's size and box are
// written only when they change. Each one keeps an interaction on a slide of thousands of shapes
// within a frame or two.

import type { ViewRect } from '../output/index.ts';

type SlideSize = { readonly width: number; readonly height: number };

/**
 * The part of slide space the editing surface can show (`PaintOptions.view`): the slide, and
 * past each edge as far as a window of `windowPx` CSS pixels (its larger side) reaches at
 * `zoom`. The surface does not clip the slide, but the window does: nothing past that reach
 * is ever on screen. The margin is a power of two times the slide's larger side, so zooming
 * changes the view, and repaints the text, only when the reach passes the next power. A
 * slide whose text overflows by thousands of lines then paints only the lines that can show.
 * @internal
 */
export function surfaceView(
  slide: SlideSize,
  windowPx: number,
  zoom: number,
  emuPerPx: number
): ViewRect {
  const { width: slideWidth, height: slideHeight } = slide;
  const side = Math.max(slideWidth, slideHeight, 1);
  const reach = zoom > 0 && windowPx > 0 ? (windowPx * emuPerPx) / zoom : side;
  const margin = side * 2 ** Math.max(0, Math.ceil(Math.log2(reach / side)));
  return {
    x: -margin,
    y: -margin,
    width: slideWidth + 2 * margin,
    height: slideHeight + 2 * margin,
  };
}

/**
 * The surface's view (`surfaceView`) for the window it is in. A window resize at a fixed zoom
 * renders nothing else, so a window that grows past the next power of two repaints here: the
 * overflowing text that it can now show paints on the next frame. @internal
 */
export class WindowReach {
  readonly #frame: NextFrame;
  /** The window size and zoom of the last view, or null before one. */
  #last: { readonly px: number; readonly zoom: number } | null = null;

  constructor(
    private readonly win: Window | null,
    private readonly emuPerPx: number,
    /** The slide on the surface and its zoom, or undefined for none. */
    private readonly shown: () => { readonly slide: SlideSize; readonly zoom: number } | undefined,
    private readonly repaint: () => void
  ) {
    this.#frame = new NextFrame(win);
    win?.addEventListener('resize', this.#onResize);
  }

  #px(): number {
    return Math.max(this.win?.innerWidth ?? 0, this.win?.innerHeight ?? 0);
  }

  /** The view of `slide` at `zoom` in this window, which the next resize compares with. */
  view(slide: SlideSize, zoom: number): ViewRect {
    const px = this.#px();
    this.#last = { px, zoom };
    return surfaceView(slide, px, zoom, this.emuPerPx);
  }

  #onResize = (): void => {
    const shown = this.shown();
    const last = this.#last;
    if (!shown || !last) return;
    const before = surfaceView(shown.slide, last.px, last.zoom, this.emuPerPx);
    const now = surfaceView(shown.slide, this.#px(), shown.zoom, this.emuPerPx);
    // A smaller view can wait for the next render: what is painted past it never shows.
    if (now.width > before.width) this.#frame.request(this.repaint);
  };

  /** Stop listening to the window. */
  destroy(): void {
    this.win?.removeEventListener('resize', this.#onResize);
    this.#frame.stop();
  }
}

/**
 * The new items that one frame paints. A slide with more goes on screen over several frames
 * (`SlidePainter.paint`), so opening a slide of 24,000 shapes never blocks for seconds. Until a
 * shape paints again, its earlier node stays on screen. @internal
 */
export const PAINT_BUDGET = 2_000;

/**
 * Runs a callback on the next frame, once, until it runs or stops. The surface uses it to paint
 * the items that a paint budget left out. @internal
 */
export class NextFrame {
  #cancel: (() => void) | null = null;

  constructor(private readonly view: Window | null) {}

  /** Run `callback` on the next frame, unless a run already waits. */
  request(callback: () => void): void {
    if (this.#cancel) return;
    const run = (): void => {
      this.#cancel = null;
      callback();
    };
    const view = this.view;
    if (view?.requestAnimationFrame) {
      const frame = view.requestAnimationFrame(run);
      this.#cancel = () => view.cancelAnimationFrame(frame);
    } else {
      const timer = setTimeout(run, 0);
      this.#cancel = () => clearTimeout(timer);
    }
  }

  /** Stop the run that waits, if any. */
  stop(): void {
    this.#cancel?.();
    this.#cancel = null;
  }
}

/**
 * The layer over the slide that shows the pointer's cursor. The cursor is not set on an
 * ancestor of the slide: a change there restyles every node of the slide (90 ms for a text of
 * 50,000 runs), and here it restyles one element. @internal
 */
export class CursorLayer {
  readonly element: HTMLDivElement;

  constructor(doc: Document) {
    this.element = doc.createElement('div');
    this.element.className = 'pptx-surface__hit';
  }

  /**
   * Show `cursor`. With `slide`, the painted slide, the point is past the slide's edge, where a
   * shape can reach and the slide itself is under the pointer: the slide shows it there, which
   * restyles the slide, but only when the cursor changes.
   */
  show(cursor: string, slide: SVGElement | null, outside: boolean): void {
    setStyle(this.element, 'cursor', cursor);
    if (slide) setStyle(slide, 'cursor', outside ? cursor : '');
  }
}

/** Set an attribute only when its value changes. @internal */
export function setAttr(el: Element, name: string, value: string): void {
  if (el.getAttribute(name) !== value) el.setAttribute(name, value);
}

/**
 * Set an inline style property only when its value changes: a write of the same value still
 * marks the element for style, and a slide of 10,000 shapes then takes 90 ms to lay out. @internal
 */
export function setStyle(
  el: HTMLElement | SVGElement,
  name: 'width' | 'height' | 'cursor',
  value: string
): void {
  if (el.style[name] !== value) el.style[name] = value;
}

/** Put the slide's picture `svg` first in `stage`, in place of an earlier one. */
export function placeSlide(stage: Element, svg: Element): void {
  setAttr(svg, 'class', 'pptx-surface__slide');
  const old = stage.querySelector('.pptx-surface__slide');
  if (old === svg) return;
  if (old) old.replaceWith(svg);
  else stage.prepend(svg);
}

/** Set the width and height of each of `svgs` in pixels. */
export function sizeAll(svgs: readonly (Element | null)[], w: number, h: number): void {
  for (const svg of svgs) {
    if (!svg) continue;
    setAttr(svg, 'width', String(w));
    setAttr(svg, 'height', String(h));
  }
}
