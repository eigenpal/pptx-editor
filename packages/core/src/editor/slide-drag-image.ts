// The drag image and drag data of a slide drag in the slide rail and the slide sorter. Without
// a drag image, the browser takes a snapshot of the dragged item. That snapshot can show other
// content of the list: a thumbnail skips paint while out of view (`content-visibility: auto`),
// and the list scrolls and repaints when the dragged slide becomes current. The drag image here
// is a copy of the dragged slide's thumbnail alone, with a count when the drag moves more than
// one slide.
//
// Adapter glue, not a stable API: `@internal`.

/** The class of the drag image element. @internal */
export const SLIDE_DRAG_IMAGE_CLASS = 'pptx-slide-drag-image';

/**
 * The drag data type of a slide drag. Some browsers start a drag only when it has data. A type
 * of the editor's own keeps a text field from taking the drop as text. @internal
 */
export const SLIDE_DRAG_TYPE = 'application/x-pptx-editor-slides';

/** The width of the drag image in CSS pixels, at most. */
const MAX_WIDTH = 200;

/**
 * The theme values that the drag image reads. The image goes on the document or shadow root,
 * out of the editor's scope, so it takes them, the chrome font, and the direction inline from
 * the dragged item.
 */
const THEME = ['--doc-primary', '--doc-on-primary', '--doc-page-bg', '--doc-bg-hover'] as const;

/**
 * Start the drag data and the drag image of a slide drag on `item` (a rail or sorter item).
 * Call it in `dragstart`, before anything moves the list. The image is a copy of the item's
 * thumbnail (`.pptx-rail__thumb`), with the number of slides that move (`count`) when more than
 * one moves. The pointer keeps its place on the thumbnail. Without a thumbnail, a measured box,
 * or `setDragImage`, the browser's own drag image stays. @internal
 */
export function setSlideDragImage(
  e: { readonly dataTransfer: DataTransfer | null; clientX: number; clientY: number },
  item: Element,
  count: number
): void {
  const transfer = e.dataTransfer;
  if (!transfer) return;
  if (typeof transfer.setData === 'function') transfer.setData(SLIDE_DRAG_TYPE, '');
  const thumb = item.querySelector<HTMLElement>('.pptx-rail__thumb');
  const doc = item.ownerDocument;
  const view = doc.defaultView;
  if (!thumb || !view || typeof transfer.setDragImage !== 'function') return;
  const box = thumb.getBoundingClientRect();
  if (box.width <= 0 || box.height <= 0) return;
  const scale = Math.min(1, MAX_WIDTH / box.width);
  const width = box.width * scale;
  const height = box.height * scale;

  const image = doc.createElement('div');
  image.className = SLIDE_DRAG_IMAGE_CLASS;
  image.setAttribute('aria-hidden', 'true');
  // A hidden slide's image is dim, as its thumbnail.
  if (item.hasAttribute('data-hidden')) image.setAttribute('data-hidden', '');
  const theme = view.getComputedStyle(item);
  for (const name of THEME) {
    const value = theme.getPropertyValue(name).trim();
    if (value) image.style.setProperty(name, value);
  }
  image.style.fontFamily = theme.fontFamily;
  image.style.direction = theme.direction;
  image.style.width = `${width}px`;
  image.style.height = `${height}px`;
  const copy = thumb.cloneNode(true) as HTMLElement;
  // The copy lives for one task. Its references (`url(#…)`) find the thumbnail's definitions,
  // and its own ids would repeat them.
  for (const node of copy.querySelectorAll('[id]')) node.removeAttribute('id');
  copy.removeAttribute('style');
  copy.style.width = `${box.width}px`;
  copy.style.height = `${box.height}px`;
  if (scale < 1) copy.style.transform = `scale(${scale})`;
  image.append(copy);
  if (count > 1) {
    const badge = doc.createElement('span');
    badge.className = `${SLIDE_DRAG_IMAGE_CLASS}__count`;
    badge.textContent = String(count);
    image.append(badge);
  }
  // On the root of the item's tree: no transformed or clipping host can move or show it. In a
  // shadow root, the editor's styles and the thumbnail's definitions are in that tree.
  const root = item.getRootNode();
  if (root.nodeType === 11) root.appendChild(image);
  else (doc.body ?? doc.documentElement).append(image);

  const x = clamp((e.clientX - box.left) * scale, 0, width);
  const y = clamp((e.clientY - box.top) * scale, 0, height);
  transfer.setDragImage(image, x, y);
  // The browser takes its snapshot when `dragstart` returns; the element can go after that.
  view.setTimeout(() => image.remove(), 0);
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}
