// The drag image of a slide drag: a copy of the dragged thumbnail alone, never a snapshot of
// the list around it, with the editor's colors and the slide drag data.
import { afterAll, afterEach, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { SLIDE_DRAG_IMAGE_CLASS, SLIDE_DRAG_TYPE, setSlideDragImage } from '../slide-drag-image.ts';

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());
afterEach(() => doc.body.replaceChildren());

interface Drag {
  readonly images: { image: HTMLElement; x: number; y: number }[];
  readonly types: string[];
  readonly transfer: DataTransfer;
}

function drag(): Drag {
  const images: Drag['images'] = [];
  const types: string[] = [];
  const transfer = {
    setData: (type: string) => types.push(type),
    setDragImage: (image: HTMLElement, x: number, y: number) => images.push({ image, x, y }),
  } as unknown as DataTransfer;
  return { images, types, transfer };
}

/** A rail item in an editor root, with a thumbnail box at (10, 20). */
function railItem(width = 160, height = 90, hidden = false): HTMLElement {
  const root = doc.createElement('div');
  root.className = 'pptx-editor';
  const item = doc.createElement('li');
  item.className = 'pptx-rail__item';
  // The theme the editor root gives the item.
  item.style.setProperty('--doc-primary', 'rgb(1, 2, 3)');
  if (hidden) item.setAttribute('data-hidden', '');
  const thumb = doc.createElement('div');
  thumb.className = 'pptx-rail__thumb';
  thumb.style.aspectRatio = '1.7778';
  thumb.innerHTML =
    '<svg><defs><clipPath id="t0n1c"></clipPath></defs>' +
    '<image href="blob:picture" clip-path="url(#t0n1c)"></image></svg>';
  thumb.getBoundingClientRect = () =>
    ({ left: 10, top: 20, width, height, right: 10 + width, bottom: 20 + height }) as DOMRect;
  item.append(thumb);
  root.append(item);
  doc.body.append(root);
  return item;
}

describe('setSlideDragImage', () => {
  test('uses a copy of the dragged thumbnail, held where the pointer took it', async () => {
    const d = drag();
    const item = railItem();
    setSlideDragImage({ dataTransfer: d.transfer, clientX: 50, clientY: 40 }, item, 1);
    expect(d.types).toEqual([SLIDE_DRAG_TYPE]);
    expect(d.images).toHaveLength(1);
    const { image, x, y } = d.images[0]!;
    expect(image.className).toBe(SLIDE_DRAG_IMAGE_CLASS);
    expect([x, y]).toEqual([40, 20]);
    // Only the thumbnail, with its picture: the item and the list stay out of the image.
    expect(image.contains(item)).toBe(false);
    expect(image.querySelector('.pptx-rail__thumb image')?.getAttribute('href')).toBe(
      'blob:picture'
    );
    expect(image.querySelector(`.${SLIDE_DRAG_IMAGE_CLASS}__count`)).toBeNull();
    expect(image.hasAttribute('data-hidden')).toBe(false);
    // On the document root, with the editor's colors; no id repeats the thumbnail's.
    expect(image.parentElement).toBe(doc.body);
    expect(image.style.getPropertyValue('--doc-primary')).toBe('rgb(1, 2, 3)');
    expect(image.querySelectorAll('[id]')).toHaveLength(0);
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(image.isConnected).toBe(false);
  });

  test('shows the number of slides that move, and dims a hidden slide', () => {
    const d = drag();
    setSlideDragImage(
      { dataTransfer: d.transfer, clientX: 0, clientY: 0 },
      railItem(160, 90, true),
      3
    );
    const { image } = d.images[0]!;
    expect(image.querySelector(`.${SLIDE_DRAG_IMAGE_CLASS}__count`)?.textContent).toBe('3');
    expect(image.hasAttribute('data-hidden')).toBe(true);
  });

  test('scales a large sorter thumbnail down, and keeps the offset inside the image', () => {
    const d = drag();
    setSlideDragImage(
      { dataTransfer: d.transfer, clientX: 1000, clientY: 120 },
      railItem(400, 225),
      1
    );
    const { image, x, y } = d.images[0]!;
    expect(image.style.width).toBe('200px');
    expect(image.style.height).toBe('112.5px');
    expect([x, y]).toEqual([200, 50]);
  });

  test('goes into the shadow root of an editor in one', () => {
    const d = drag();
    const host = doc.createElement('div');
    doc.body.append(host);
    const shadow = host.attachShadow({ mode: 'open' });
    const item = railItem();
    shadow.append(item.parentElement!);
    setSlideDragImage({ dataTransfer: d.transfer, clientX: 0, clientY: 0 }, item, 1);
    expect(d.images[0]!.image.parentNode).toBe(shadow);
  });

  test('keeps the browser image without a thumbnail, and does nothing without a transfer', () => {
    const d = drag();
    const bare = doc.createElement('li');
    doc.body.append(bare);
    setSlideDragImage({ dataTransfer: d.transfer, clientX: 0, clientY: 0 }, bare, 1);
    expect(d.types).toEqual([SLIDE_DRAG_TYPE]);
    setSlideDragImage({ dataTransfer: null, clientX: 0, clientY: 0 }, railItem(), 1);
    expect(d.images).toHaveLength(0);
    expect(doc.querySelector(`.${SLIDE_DRAG_IMAGE_CLASS}`)).toBeNull();
  });
});
