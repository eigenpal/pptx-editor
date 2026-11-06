// How many items the slide rail mounts while a long deck opens. Shared by the React and Vue
// rails, so both mount the same items at the same time.
//
// Mounting one rail item for each of thousands of slides took longer than opening the first
// slide, and it ran in the same task, so the first slide showed only after the whole rail
// mounted. While a deck opens (`EditorSnapshot.pendingSlides`), the rail mounts its first
// screen with the first slide, then more after each painted frame, as DOCX Editor shows its
// first pages while the rest lay out. Each step mounts what fits in about 8 ms
// (`MOUNT_FRAME_MS`): the rail times the step before and sizes the next one from it, so a
// keystroke between two frames waits for one short task, not for a thousand items. Items up to the current slide always mount, so
// the current slide's item is there to focus and to scroll into view. A deck that opens whole
// mounts every item at once, as before. A rail grouped in sections mounts whole sections
// (`mountedSections`), so it never shows a section header with no slides under it.

import type { SlideSection } from '../contracts/index.ts';

/** Items the rail mounts with the first slide: more than one screen of thumbnails. @internal */
export const MOUNT_FIRST_ITEMS = 40;
/** The most items one step adds, however cheap the items were. @internal */
export const MOUNT_CHUNK_ITEMS = 1000;
/** The items the first step adds, before the rail has timed a step. @internal */
export const MOUNT_FIRST_STEP_ITEMS = 100;
/** The fewest items one step adds, so the rail always moves. @internal */
export const MOUNT_MIN_STEP_ITEMS = 10;
/** The time one step may take to mount its items, in ms. @internal */
export const MOUNT_FRAME_MS = 8;

/** The mount limit of one rail. @internal */
export interface ProgressiveMount {
  /**
   * How many of `total` items to mount now. `needed` (the current slide) is always within the
   * limit. `opening` says whether the deck is still opening: an open that starts goes back to
   * the first screen, and a deck that opened whole mounts every item.
   */
  limit(total: number, needed: number, opening: boolean): number;
  /**
   * Call after each render of the rail. Grows the limit after the next painted frame, if
   * items are still left out; the render that the last growth caused sizes the next one.
   */
  follow(total: number): void;
  dispose(): void;
}

/** The sections a grouped rail mounts, and the item limit that fills them. @internal */
export interface MountedSections {
  readonly sections: readonly SlideSection[];
  readonly limit: number;
}

/**
 * The sections of `grouped` to mount for the item limit `limit`: each section that starts
 * below the limit. The limit grows to the end of the last one, so no section shows only part
 * of its slides; when that adds more than one first step ({@link MOUNT_FIRST_STEP_ITEMS}),
 * the section mounts step by step under its header instead, so a step stays short. @internal
 */
export function mountedSections(grouped: readonly SlideSection[], limit: number): MountedSections {
  let end = limit;
  for (const s of grouped)
    if (s.firstSlide < limit) end = Math.max(end, s.firstSlide + s.slideCount);
  const filled = end - limit <= MOUNT_FIRST_STEP_ITEMS ? end : limit;
  // An empty section shows when the slides before it do.
  const sections = grouped.filter(
    (s) => s.firstSlide < filled || (s.slideCount === 0 && s.firstSlide <= filled)
  );
  return { sections, limit: filled };
}

/**
 * The `aria-setsize` and `aria-posinset` of rail item `index` while the rail mounts only
 * `limit` of its `count` items, so a screen reader still hears "slide 3 of 5,000" (or the
 * place in its section, in a grouped rail). Null once every item is mounted: the list then
 * says it itself. @internal
 */
export function railItemPosition(
  index: number,
  count: number,
  limit: number,
  grouped: readonly SlideSection[] | null
): { readonly setSize: number; readonly posInSet: number } | null {
  if (limit >= count) return null;
  const section = grouped?.find(
    (s) => index >= s.firstSlide && index < s.firstSlide + s.slideCount
  );
  return section
    ? { setSize: section.slideCount, posInSet: index - section.firstSlide + 1 }
    : { setSize: count, posInSet: index + 1 };
}

/** After the next frame paints: `requestAnimationFrame`, then a task. Outside a browser, a task. */
function afterPaint(run: () => void): () => void {
  let timer: ReturnType<typeof setTimeout> | null = null;
  let frame: number | null = null;
  if (typeof requestAnimationFrame === 'function')
    frame = requestAnimationFrame(() => {
      frame = null;
      timer = setTimeout(run, 0);
    });
  else timer = setTimeout(run, 0);
  return () => {
    if (frame !== null && typeof cancelAnimationFrame === 'function') cancelAnimationFrame(frame);
    if (timer !== null) clearTimeout(timer);
  };
}

/**
 * A rail's mount limit; `changed` re-renders the rail when the limit grows. `now` is the clock
 * that times each step. @internal
 */
export function createProgressiveMount(
  changed: () => void,
  now: () => number = () => (globalThis.performance ?? Date).now()
): ProgressiveMount {
  let limit = Number.POSITIVE_INFINITY;
  let wasOpening = false;
  let cancel: (() => void) | null = null;
  let step = MOUNT_FIRST_STEP_ITEMS;
  /** The last growth, until the render it caused is timed. */
  let grew: { at: number; items: number } | null = null;
  /** The last timed step, and the cost of one more item. */
  let timed: { items: number; spent: number } | null = null;
  let each = 0;
  return {
    limit(total, needed, opening) {
      if (opening && !wasOpening) {
        cancel?.();
        cancel = null;
        grew = null;
        timed = null;
        step = MOUNT_FIRST_STEP_ITEMS;
        limit = MOUNT_FIRST_ITEMS;
      }
      wasOpening = opening;
      if (needed >= limit) limit = needed + 1;
      return Math.min(total, limit);
    },
    follow(total) {
      if (grew) {
        // The render after a growth. A render also costs the same for every step (the list
        // itself), so the cost of one more item is the slope between two steps of other sizes.
        const spent = now() - grew.at;
        if (!timed) each = spent / grew.items;
        else if (timed.items !== grew.items)
          each = Math.max(0, (spent - timed.spent) / (grew.items - timed.items));
        timed = { items: grew.items, spent };
        step =
          each > 0
            ? Math.max(
                MOUNT_MIN_STEP_ITEMS,
                Math.min(MOUNT_CHUNK_ITEMS, Math.floor(MOUNT_FRAME_MS / each))
              )
            : MOUNT_CHUNK_ITEMS;
        grew = null;
      }
      if (cancel || limit >= total) return;
      cancel = afterPaint(() => {
        cancel = null;
        grew = { at: now(), items: step };
        limit += step;
        changed();
      });
    },
    dispose() {
      cancel?.();
      cancel = null;
    },
  };
}
