// The slide rail's orientation: a column beside the slide (`vertical`, the default) or a row
// under it (`horizontal`), for a tall, narrow host such as a panel beside a chat. The React and
// Vue rails read the keys, the icons, and the labels from here, so both act the same. The
// scroll helper keeps the current slide in view inside the rail only.

import { warnInvalidRailOrientation } from './dev-warnings.ts';

/** Where the slide rail runs: a column (`vertical`) or a row (`horizontal`). @public */
export type SlideRailOrientation = 'vertical' | 'horizontal';

/** The rail's icons and the translation keys of its labels for an orientation. @internal */
export interface RailOrientationChrome {
  readonly moveEarlierIcon: string;
  readonly moveLaterIcon: string;
  readonly moveEarlierLabel: 'rail.moveUp' | 'rail.moveLeft';
  readonly moveLaterLabel: 'rail.moveDown' | 'rail.moveRight';
  /** The Hide thumbnails button: it points to where the rail goes. */
  readonly collapseIcon: string;
  /** The Show thumbnails button of the closed strip: it points to where the rail comes from. */
  readonly expandIcon: string;
}

const VERTICAL: RailOrientationChrome = {
  moveEarlierIcon: 'keyboard_arrow_up',
  moveLaterIcon: 'keyboard_arrow_down',
  moveEarlierLabel: 'rail.moveUp',
  moveLaterLabel: 'rail.moveDown',
  collapseIcon: 'keyboard_arrow_left',
  expandIcon: 'keyboard_arrow_right',
};

const HORIZONTAL: RailOrientationChrome = {
  moveEarlierIcon: 'keyboard_arrow_left',
  moveLaterIcon: 'keyboard_arrow_right',
  moveEarlierLabel: 'rail.moveLeft',
  moveLaterLabel: 'rail.moveRight',
  collapseIcon: 'keyboard_arrow_down',
  expandIcon: 'keyboard_arrow_up',
};

/** The rail's icons and labels for `orientation`. @internal */
export function railOrientationChrome(
  orientation: SlideRailOrientation | undefined
): RailOrientationChrome {
  return orientation === 'horizontal' ? HORIZONTAL : VERTICAL;
}

/**
 * The orientation of the rail that holds `el`, from its `data-orientation` attribute: the
 * section headers read it, so they need no prop of their own. @internal
 */
export function railOrientationOf(el: Element | null): SlideRailOrientation {
  return el?.closest('.pptx-rail')?.getAttribute('data-orientation') === 'horizontal'
    ? 'horizontal'
    : 'vertical';
}

/** The drag axis of a rail: the direction in which its slides read. @internal */
export function railDragAxis(orientation: SlideRailOrientation | undefined): 'x' | 'y' {
  return orientation === 'horizontal' ? 'x' : 'y';
}

/** A right-to-left list reads its row from the right. */
function rightToLeft(list: HTMLElement | null): boolean {
  if (!list) return false;
  const view = list.ownerDocument.defaultView;
  return view?.getComputedStyle(list).direction === 'rtl';
}

/**
 * The slide step of an arrow key in a rail: -1 to the slide before, 1 to the slide after, or 0
 * for another key. A vertical rail steps with Up and Down. A horizontal rail steps with Left
 * and Right, mirrored in a right-to-left list. @internal
 */
export function railStepOfKey(
  key: string,
  orientation: SlideRailOrientation | undefined,
  list: HTMLElement | null
): -1 | 0 | 1 {
  if (orientation !== 'horizontal') return key === 'ArrowUp' ? -1 : key === 'ArrowDown' ? 1 : 0;
  const step = key === 'ArrowLeft' ? -1 : key === 'ArrowRight' ? 1 : 0;
  return rightToLeft(list) ? ((-step || 0) as -1 | 0 | 1) : step;
}

/**
 * The key that moves from a slide to its section header in a rail with sections: Left in a
 * vertical rail (the header is the tree parent), Up in a horizontal rail (the header is above
 * its slides). @internal
 */
export function railParentKey(orientation: SlideRailOrientation | undefined): string {
  return orientation === 'horizontal' ? 'ArrowUp' : 'ArrowLeft';
}

/**
 * The rail's `orientation` prop as a value the rail can use: `horizontal`, or `vertical` for
 * anything else. A value that is neither prints a development warning naming `part`. @internal
 */
export function resolveRailOrientation(value: unknown, part: string): SlideRailOrientation {
  if (value === 'horizontal') return 'horizontal';
  if (value !== undefined && value !== 'vertical') warnInvalidRailOrientation(part, value);
  return 'vertical';
}

const SCROLLS = /(auto|scroll)/;

type Axes = { readonly el: HTMLElement; readonly x: boolean; readonly y: boolean };

/**
 * The nearest ancestor of `el` that scrolls, and the axes on which it does, or null. The search
 * never goes past `within`: the editor's own pane, so a host container is never scrolled.
 */
function scrollParent(el: HTMLElement, within: Element | null): Axes | null {
  const view = el.ownerDocument.defaultView;
  for (let p = el.parentElement; p; p = p.parentElement) {
    const style = view?.getComputedStyle(p);
    if (!style) return null;
    const x = SCROLLS.test(style.overflowX) && p.scrollWidth > p.clientWidth;
    const y = SCROLLS.test(style.overflowY) && p.scrollHeight > p.clientHeight;
    if (x || y) return { el: p, x, y };
    if (p === within) return null;
  }
  return null;
}

/** A `scroll-padding-*` length in pixels; `auto` and anything else read as 0. */
const px = (value: string | undefined): number => parseFloat(value ?? '') || 0;

/** The least distance that brings `[start, start + size]` into `[pad0, length - pad1]`. */
function nearest(start: number, size: number, length: number, pad0: number, pad1: number): number {
  if (start < pad0) return start - pad0;
  const end = start + size - (length - pad1);
  return end > 0 ? Math.min(end, start - pad0) : 0;
}

/**
 * Scroll `target` into view in its nearest scrolling ancestor only, by the least distance on
 * each axis that scrolls, and inside its `scroll-padding`. Unlike `scrollIntoView`, it leaves
 * the outer page alone: a host that slides the editor in from off screen, or keeps it in a
 * clipped panel, does not see its layout move. `within` bounds the search (the rail or the
 * sorter): an ancestor outside it is never scrolled. @internal
 */
export function scrollIntoNearestScroller(
  target: HTMLElement,
  within: Element | null = null
): void {
  const found = scrollParent(target, within);
  if (!found) return;
  const { el: scroller } = found;
  const style = scroller.ownerDocument.defaultView?.getComputedStyle(scroller);
  const box = target.getBoundingClientRect();
  const view = scroller.getBoundingClientRect();
  const top = box.top - view.top - scroller.clientTop;
  const left = box.left - view.left - scroller.clientLeft;
  const dy = found.y
    ? nearest(
        top,
        box.height,
        scroller.clientHeight,
        px(style?.scrollPaddingTop),
        px(style?.scrollPaddingBottom)
      )
    : 0;
  const dx = found.x
    ? nearest(
        left,
        box.width,
        scroller.clientWidth,
        px(style?.scrollPaddingLeft),
        px(style?.scrollPaddingRight)
      )
    : 0;
  if (dx || dy) scroller.scrollBy?.({ left: dx, top: dy });
}
