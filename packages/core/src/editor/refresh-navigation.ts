// Bring a slide's rail item into view for `navigateToChange`, after DOCX Editor's reveal scroll
// options (docx-editor/packages/core/src/editor/reveal-scroll-options.ts at b63cbea4a).
//
// The rail and the slide sorter mark each item with `data-index`. The scroll moves only the
// rail's own scroll box: the page does not jump. Reduced motion makes smooth scrolling instant.

import type { ScrollPlacement } from './document-refresh-types.ts';
import { surfaceRootOf } from './surface-registry.ts';

/** The nearest ancestor of `el` that scrolls vertically or horizontally, or null. */
function scrollBox(el: HTMLElement): HTMLElement | null {
  const win = el.ownerDocument.defaultView;
  for (let p = el.parentElement; p; p = p.parentElement) {
    const style = win?.getComputedStyle(p);
    const y = style?.overflowY;
    const x = style?.overflowX;
    if (y === 'auto' || y === 'scroll' || x === 'auto' || x === 'scroll') return p;
  }
  return null;
}

/**
 * Scroll the rail so slide `index` shows: `start`, `center`, `nearest`, or `centerIfNeeded`
 * (no scroll when it shows already). @internal
 */
export function revealRailItem(
  editor: object,
  index: number,
  options: { block: ScrollPlacement; behavior: 'instant' | 'smooth'; offsetPx: number }
): void {
  const root = surfaceRootOf(editor);
  const scope = (root?.closest('.pptx-editor') as HTMLElement | null) ?? null;
  const item = scope?.querySelector<HTMLElement>(`.pptx-rail [data-index="${index}"]`);
  if (!item) return;
  const box = scrollBox(item);
  if (!box) return;
  const win = item.ownerDocument.defaultView;
  let reduced = false;
  try {
    reduced = !!win?.matchMedia?.('(prefers-reduced-motion: reduce)').matches;
  } catch {
    reduced = false;
  }
  const behavior: ScrollBehavior = options.behavior === 'smooth' && !reduced ? 'smooth' : 'auto';
  const a = item.getBoundingClientRect();
  const b = box.getBoundingClientRect();
  const vertical = box.scrollHeight > box.clientHeight;
  const [itemStart, itemSize, boxStart, boxSize] = vertical
    ? [a.top, a.height, b.top, box.clientHeight]
    : [a.left, a.width, b.left, box.clientWidth];
  const visible = itemStart >= boxStart && itemStart + itemSize <= boxStart + boxSize;
  if (options.block === 'centerIfNeeded' && visible) return;
  const pad = options.offsetPx;
  let delta: number;
  if (options.block === 'start') delta = itemStart - boxStart - pad;
  else if (options.block === 'nearest') {
    if (visible) return;
    delta =
      itemStart < boxStart
        ? itemStart - boxStart - pad
        : itemStart + itemSize - (boxStart + boxSize) + pad;
  } else delta = itemStart + itemSize / 2 - (boxStart + boxSize / 2);
  if (Math.abs(delta) < 0.5) return;
  box.scrollBy(vertical ? { top: delta, behavior } : { left: delta, behavior });
}
