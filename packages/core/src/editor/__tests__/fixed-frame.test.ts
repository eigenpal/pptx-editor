// Fixed chrome in a transformed, scaled, or clipping host: a place in the visible area becomes
// the `left`/`top`/`maxHeight` that put the panel there, measured by a probe in the panel's
// parent. The browser check is in `e2e/browser.spec.ts` ("embedding: in a transformed").

import { describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { positionPopover } from '../chrome-popover.ts';
import { inFixedView, placeInFixedFrame, type FixedSize, type FixedView } from '../index.ts';

// A DOM of this file's own: the core run has no global DOM (`bun run test`).
const win = new GlobalWindow({ width: 1000, height: 800 });
const document = win.document as unknown as Document;

const box = (left: number, top: number, width: number, height: number): DOMRect =>
  ({ left, top, width, height, right: left + width, bottom: top + height }) as DOMRect;

/** A host whose fixed children measure as if its frame starts at `x`, `y` with `scale`. */
function host(x: number, y: number, scale = 1, parent: Element = document.body): HTMLElement {
  const el = document.createElement('div');
  parent.appendChild(el);
  const append = el.appendChild.bind(el);
  el.appendChild = <T extends Node>(child: T): T => {
    (child as unknown as HTMLElement).getBoundingClientRect = () =>
      box(x, y, 100 * scale, 100 * scale);
    return append(child);
  };
  return el;
}

const SIZE = { width: 200, height: 120 };
const WINDOW: FixedView = { left: 0, top: 0, width: 1000, height: 800 };

describe('placeInFixedFrame', () => {
  test('a host with no frame of its own places in the whole window, unchanged', () => {
    const seen: [FixedSize, FixedView][] = [];
    const at = { left: 40, top: 60, maxHeight: 300 };
    const out = placeInFixedFrame(host(0, 0), SIZE, (size, view) => (seen.push([size, view]), at));
    expect(out).toBe(at);
    expect(seen).toEqual([[SIZE, WINDOW]]);
  });

  test('a translated frame moves the place back by its origin', () => {
    const el = host(37, 23);
    el.style.transform = 'translate(37px, 23px)';
    const out = placeInFixedFrame(el, SIZE, () => ({ left: 100, top: 50, maxHeight: 300 }));
    expect(out).toEqual({ left: 63, top: 27, maxHeight: 300 });
  });

  test('a scaled frame places the scaled panel, then divides by the scale', () => {
    const el = host(10, 20, 2);
    el.style.transform = 'scale(2)';
    const seen: FixedSize[] = [];
    const out = placeInFixedFrame(el, SIZE, (size) => {
      seen.push(size);
      return { left: 110, top: 220, maxHeight: 400, scroll: false };
    });
    // The window sees the panel at twice its layout size, for the edge clamp.
    expect(seen).toEqual([{ width: 400, height: 240 }]);
    expect(out).toEqual({ left: 50, top: 100, maxHeight: 200, scroll: false });
  });

  test('a clipping container around the frame limits the area the panel can show in', () => {
    // A drawer: clips its overflow and slides in with a transform.
    const drawer = document.createElement('div');
    drawer.style.overflow = 'hidden';
    drawer.style.transform = 'translateX(0)';
    drawer.getBoundingClientRect = () => box(600, 100, 300, 400);
    document.body.appendChild(drawer);
    const el = host(600, 100, 1, drawer);
    const views: FixedView[] = [];
    const out = placeInFixedFrame(el, SIZE, (_, view) => {
      views.push(view);
      // A place 10 px into the area, as a placement rule returns it.
      return { left: 10, top: 10, maxHeight: view.height - 20 };
    });
    expect(views).toEqual([{ left: 600, top: 100, width: 300, height: 400 }]);
    // 10 px into the drawer, in the drawer's own coordinates.
    expect(out).toEqual({ left: 10, top: 10, maxHeight: 380 });
  });

  test('a clipping drawer at the window origin with no scale still limits the area', () => {
    // A left sidebar: a container type makes it hold fixed panels, and it scrolls.
    const drawer = document.createElement('div');
    drawer.style.containerType = 'inline-size';
    drawer.style.overflow = 'auto';
    drawer.getBoundingClientRect = () => box(0, 0, 320, 800);
    document.body.appendChild(drawer);
    const el = host(0, 0, 1, drawer);
    const views: FixedView[] = [];
    placeInFixedFrame(el, SIZE, (_, view) => (views.push(view), { left: 0, top: 0 }));
    expect(views).toEqual([{ left: 0, top: 0, width: 320, height: 800 }]);
  });

  test('a fixed sheet is not clipped by the page around it or by a body scroll lock', () => {
    document.body.style.overflow = 'hidden';
    document.body.getBoundingClientRect = () => box(0, 0, 1000, 200);
    const main = document.createElement('main');
    main.style.overflow = 'auto';
    main.getBoundingClientRect = () => box(0, 0, 500, 300);
    document.body.appendChild(main);
    const sheet = document.createElement('div');
    sheet.style.position = 'fixed';
    sheet.style.transform = 'translateY(0)';
    sheet.style.overflow = 'hidden';
    sheet.getBoundingClientRect = () => box(400, 100, 600, 700);
    main.appendChild(sheet);
    const el = host(400, 100, 1, sheet);
    const views: FixedView[] = [];
    placeInFixedFrame(el, SIZE, (_, view) => (views.push(view), { left: 0, top: 0 }));
    // Only the sheet's own box: not `main`'s, and not the scroll-locked body's.
    expect(views).toEqual([{ left: 400, top: 100, width: 600, height: 700 }]);
    sheet.remove();
    main.remove();
    document.body.style.overflow = '';
  });

  test('inFixedView moves rectangles and points into the area', () => {
    const view = { left: 600, top: 100, width: 300, height: 400 };
    expect(inFixedView(box(650, 150, 20, 10), view)).toEqual({
      left: 50,
      top: 50,
      width: 20,
      height: 10,
      right: 70,
      bottom: 60,
    } as DOMRect);
    expect(inFixedView({ x: 610, y: 120 }, view)).toEqual({ x: 10, y: 20 });
  });

  test('the probe leaves the host as it was', () => {
    const el = host(5, 5);
    placeInFixedFrame(el, SIZE, () => ({ left: 0, top: 0 }));
    expect(el.childElementCount).toBe(0);
  });

  test('no host, or a DOM with no layout, keeps the place', () => {
    const at = { left: 1, top: 2 };
    expect(placeInFixedFrame(null, SIZE, () => at)).toBe(at);
    const plain = document.createElement('div');
    document.body.appendChild(plain);
    expect(placeInFixedFrame(plain, SIZE, () => at)).toBe(at);
  });
});

test('a selected-text popover uses its client rectangle inside a scaled host', () => {
  const frame = host(10, 20, 2);
  frame.style.transform = 'scale(2)';
  const anchor = document.createElement('button');
  const panel = document.createElement('div');
  frame.appendChild(anchor);
  frame.appendChild(panel);
  anchor.getBoundingClientRect = () => box(10, 600, 40, 30);
  Object.defineProperties(panel, {
    offsetWidth: { value: 200 },
    offsetHeight: { value: 100 },
  });
  const selectedText = box(110, 200, 80, 40);
  const at = positionPopover(panel, anchor, 'start', selectedText);
  expect(at.left).toBe(50);
  expect(at.top).toBe(113);
  expect(panel.style.left).toBe('50px');
  expect(panel.style.top).toBe('113px');
  frame.remove();
});

test('a tall popover placed twice still scrolls: the first cap does not hide its height', () => {
  const anchor = document.createElement('button');
  const panel = document.createElement('div');
  document.body.append(anchor, panel);
  anchor.getBoundingClientRect = () => box(10, 10, 40, 30);
  // The panel is as tall as its rows, less any height cap, as a browser measures it.
  Object.defineProperties(panel, {
    offsetWidth: { value: 200 },
    offsetHeight: { get: () => Math.min(5000, parseFloat(panel.style.maxHeight) || Infinity) },
  });
  const first = positionPopover(panel, anchor, 'start');
  const second = positionPopover(panel, anchor, 'start');
  expect(first.scroll).toBe(true);
  expect(second.scroll).toBe(true);
  expect(panel.style.overflowY).toBe('auto');
  anchor.remove();
  panel.remove();
});
