// The cost of a large slide on the editing surface (surface.ts): the slide paints over several
// frames, the pointer's cursor sits on a layer over the slide, and the slide's size and box are
// written only when they change. Each one keeps an interaction on a slide of thousands of shapes
// within a frame or two.

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
