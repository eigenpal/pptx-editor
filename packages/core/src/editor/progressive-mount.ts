// How many items the slide rail mounts while a long deck opens. Shared by the React and Vue
// rails, so both mount the same items at the same time.
//
// Mounting one rail item for each of thousands of slides took longer than opening the first
// slide, and it ran in the same task, so the first slide showed only after the whole rail
// mounted. While a deck opens (`EditorSnapshot.pendingSlides`), the rail mounts its first
// screen with the first slide, then a larger chunk after each painted frame, as DOCX Editor
// shows its first pages while the rest lay out. Items up to the current slide always mount, so
// the current slide's item is there to focus and to scroll into view. A deck that opens whole
// mounts every item at once, as before. A rail grouped in sections mounts whole sections
// (`mountedSections`), so it never shows a section header with no slides under it.

import type { SlideSection } from '../contracts/index.ts';

/** Items the rail mounts with the first slide: more than one screen of thumbnails. @internal */
export const MOUNT_FIRST_ITEMS = 40;
/** Items added after each painted frame. @internal */
export const MOUNT_CHUNK_ITEMS = 1000;

/** The mount limit of one rail. @internal */
export interface ProgressiveMount {
  /**
   * How many of `total` items to mount now. `needed` (the current slide) is always within the
   * limit. `opening` says whether the deck is still opening: an open that starts goes back to
   * the first screen, and a deck that opened whole mounts every item.
   */
  limit(total: number, needed: number, opening: boolean): number;
  /** Grow the limit after the next painted frame, if items are still left out. */
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
 * of its slides; a section longer than one chunk ({@link MOUNT_CHUNK_ITEMS}) still mounts
 * chunk by chunk under its header. @internal
 */
export function mountedSections(grouped: readonly SlideSection[], limit: number): MountedSections {
  let end = limit;
  for (const s of grouped)
    if (s.firstSlide < limit) end = Math.max(end, s.firstSlide + s.slideCount);
  const filled = end - limit <= MOUNT_CHUNK_ITEMS ? end : limit;
  // An empty section shows when the slides before it do.
  const sections = grouped.filter(
    (s) => s.firstSlide < filled || (s.slideCount === 0 && s.firstSlide <= filled)
  );
  return { sections, limit: filled };
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

/** A rail's mount limit; `changed` re-renders the rail when the limit grows. @internal */
export function createProgressiveMount(changed: () => void): ProgressiveMount {
  let limit = Number.POSITIVE_INFINITY;
  let wasOpening = false;
  let cancel: (() => void) | null = null;
  return {
    limit(total, needed, opening) {
      if (opening && !wasOpening) {
        cancel?.();
        cancel = null;
        limit = MOUNT_FIRST_ITEMS;
      }
      wasOpening = opening;
      if (needed >= limit) limit = needed + 1;
      return Math.min(total, limit);
    },
    follow(total) {
      if (cancel || limit >= total) return;
      cancel = afterPaint(() => {
        cancel = null;
        limit += MOUNT_CHUNK_ITEMS;
        changed();
      });
    },
    dispose() {
      cancel?.();
      cancel = null;
    },
  };
}
