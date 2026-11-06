// A long deck opens without painting every thumbnail: the current slide's and the first of a
// task paint at once, the ones in view before the frame, and the others when they scroll into
// view or when the editor is idle. No container ever shows another slide than its own.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { textDeck } from '../../../../../scripts/bench/text-deck.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { ThumbnailPainter } from '../thumbnails.ts';

// A DOM for this file only: the engine test run registers no global DOM.
const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const wait = (ms: number): Promise<void> => new Promise((r) => setTimeout(r, ms));
/** The end of the task: the queue checks the containers that wait. */
const endOfTask = (): Promise<void> => Promise.resolve();

async function open(n: number): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(textDeck(n))).ok).toBe(true);
  return editor;
}

const ITEM_H = 100;
/** As the rail and the slide sorter ask. */
const DEFER = { defer: true } as const;

/**
 * One container for each slide in a scrolling list, as the rail mounts them. With `view`, the
 * list is laid out: `ITEM_H` pixels a container, the list 600 pixels tall, scrolled to `top`.
 */
function rail(n: number, view?: { top: number }) {
  const list = doc.createElement('ol') as unknown as HTMLElement;
  list.style.overflowY = 'auto';
  doc.body.append(list as unknown as Node);
  const rect = (top: number, h: number) =>
    ({ left: 0, right: 200, top, bottom: top + h, width: 200, height: h }) as DOMRect;
  if (view) list.getBoundingClientRect = () => rect(0, 600);
  const boxes = Array.from({ length: n }, (_, i) => {
    const li = doc.createElement('li');
    const box = doc.createElement('div') as unknown as HTMLElement;
    if (view) box.getBoundingClientRect = () => rect(i * ITEM_H - view.top, ITEM_H);
    li.append(box as unknown as Node);
    list.append(li);
    return box;
  });
  return { list, boxes };
}

const painted = (boxes: HTMLElement[]): number =>
  boxes.filter((b) => b.firstElementChild?.localName === 'svg').length;

async function until(check: () => boolean): Promise<void> {
  for (let i = 0; i < 400 && !check(); i += 1) await wait(5);
}

describe('thumbnails of a long deck', () => {
  test('the current slide and the first screen paint at once, the rest when idle', async () => {
    const editor = await open(40);
    editor.goToSlide(30);
    const { list, boxes } = rail(40);
    boxes.forEach((box, i) => editor.renderThumbnail(i, box, DEFER));
    // The first 12 of the task, and the current slide's.
    expect(painted(boxes)).toBe(13);
    expect(painted(boxes.slice(0, 12))).toBe(12);
    expect(painted([boxes[30]!])).toBe(1);
    await until(() => painted(boxes) === 40);
    expect(painted(boxes)).toBe(40);
    editor.destroy();
    list.remove();
  });

  test('a container that waits shows no slide, never another one', async () => {
    const editor = await open(30);
    const { list, boxes } = rail(30);
    boxes.forEach((box, i) => editor.renderThumbnail(i, box, DEFER));
    await until(() => painted(boxes) === 30);
    // Every container moves on by one slide (a slide inserted at the top).
    const before = boxes.map((b) => b.firstElementChild);
    await endOfTask();
    boxes.slice(0, 29).forEach((box, i) => editor.renderThumbnail(i + 1, box, DEFER));
    await endOfTask();
    expect(boxes.slice(0, 12).every((b, i) => b.firstElementChild !== before[i])).toBe(true);
    // No layout here: the others wait, empty.
    expect(boxes.slice(12, 29).every((b) => b.firstElementChild === null)).toBe(true);
    await until(() => painted(boxes) === 30);
    boxes
      .slice(0, 29)
      .forEach((box, i) => expect(box.textContent).toContain(`Slide ${i + 2} title`));
    editor.destroy();
    list.remove();
  });

  test('in a list scrolled down, a slide deleted above leaves no container in view empty', async () => {
    const editor = await open(60);
    // Containers 30 to 35 are in the list's box; 28 to 37 within the overscan.
    const { list, boxes } = rail(60, { top: 3000 });
    boxes.forEach((box, i) => editor.renderThumbnail(i, box, DEFER));
    await endOfTask();
    expect(painted(boxes.slice(28, 38))).toBe(10);
    await until(() => painted(boxes) === 60);
    // Slide 3 goes: each container from 3 on shows the next slide (a list keyed by index).
    await endOfTask();
    expect(editor.exec({ type: 'deleteSlides', indices: [2] }).ok).toBe(true);
    boxes.slice(0, 59).forEach((box, i) => editor.renderThumbnail(i, box, DEFER));
    await endOfTask();
    expect(boxes.slice(28, 38).every((b) => b.firstElementChild?.localName === 'svg')).toBe(true);
    boxes
      .slice(28, 38)
      .forEach((box, k) => expect(box.textContent).toContain(`Slide ${28 + k + 2} title`));
    editor.destroy();
    list.remove();
  });

  test('with no defer, every thumbnail paints at once (a handout captured in the same task)', async () => {
    const editor = await open(40);
    const { list, boxes } = rail(40);
    boxes.forEach((box, i) => editor.renderThumbnail(i, box));
    expect(painted(boxes)).toBe(40);
    editor.destroy();
    list.remove();
  });

  test('a disposer forgets its container, and does nothing after a later call', async () => {
    const editor = await open(40);
    const { list, boxes } = rail(40);
    const box = boxes[30]!;
    boxes.slice(0, 30).forEach((b, i) => editor.renderThumbnail(i, b, DEFER));
    // An old disposer, then a new ask for the same container: the new ask still paints.
    const old = editor.renderThumbnail(30, box, DEFER);
    editor.renderThumbnail(30, box, DEFER);
    old();
    // A disposed container that waits never paints.
    const next = boxes[31]!;
    editor.renderThumbnail(31, next, DEFER)();
    expect(painted([box, next])).toBe(0);
    await until(() => painted([box]) === 1);
    expect(painted([box])).toBe(1);
    await wait(100);
    expect(painted([next])).toBe(0);
    editor.destroy();
    list.remove();
  });

  test('a container that scrolls into view paints at once', async () => {
    type Entry = { target: Element; isIntersecting: boolean };
    const observers: { cb: (e: Entry[]) => void; root: Element | null; seen: Element[] }[] = [];
    class Observer {
      readonly seen: Element[] = [];
      constructor(cb: (e: Entry[]) => void, o: { root: Element | null }) {
        observers.push({ cb, root: o.root, seen: this.seen });
      }
      observe(el: Element): void {
        this.seen.push(el);
      }
      unobserve(): void {}
      disconnect(): void {}
    }
    const w = win as unknown as { IntersectionObserver: unknown };
    const had = w.IntersectionObserver;
    w.IntersectionObserver = Observer;
    try {
      const editor = await open(40);
      const { list, boxes } = rail(40);
      boxes.forEach((box, i) => editor.renderThumbnail(i, box, DEFER));
      await endOfTask();
      // One observer, on the scrolling list, for the containers that wait.
      expect(observers).toHaveLength(1);
      expect(observers[0]!.root).toBe(list as unknown as Element);
      expect(observers[0]!.seen).toEqual(boxes.slice(13) as unknown as Element[]);
      expect(painted(boxes)).toBe(13);
      // The list scrolls: 30 to 34 come into view. They paint at once, within a time slice, and
      // the rest of them in the next task, before the idle time reaches them.
      observers[0]!.cb(
        boxes.map((b, i) => ({
          target: b as unknown as Element,
          isIntersecting: i >= 30 && i < 35,
        }))
      );
      expect(painted(boxes.slice(30, 35))).toBeGreaterThan(0);
      await until(() => painted(boxes.slice(30, 35)) === 5);
      expect(painted(boxes.slice(30, 35))).toBe(5);
      expect(painted(boxes.slice(20, 30))).toBeLessThan(10);
      editor.destroy();
      list.remove();
    } finally {
      w.IntersectionObserver = had;
    }
  });

  test('a container not on the page paints at once (export, print, paint then append)', async () => {
    const editor = await open(20);
    const { list, boxes } = rail(20);
    list.remove();
    boxes.forEach((box, i) => editor.renderThumbnail(i, box, DEFER));
    expect(painted(boxes)).toBe(20);
    editor.destroy();
  });

  test('a waiting container that leaves the page paints when it is back', async () => {
    const editor = await open(20);
    const { list, boxes } = rail(20);
    boxes.forEach((box, i) => editor.renderThumbnail(i, box, DEFER));
    await endOfTask();
    list.remove();
    await wait(50);
    expect(painted(boxes)).toBe(13);
    doc.body.append(list as unknown as Node);
    await until(() => painted(boxes) === 20);
    expect(painted(boxes)).toBe(20);
    editor.destroy();
    list.remove();
  });

  test('released containers leave nothing waiting: no observer, no timer', async () => {
    type Entry = { target: Element; isIntersecting: boolean };
    let live = 0;
    class Observer {
      constructor(_cb: (e: Entry[]) => void, _o: unknown) {}
      #seen = new Set<Element>();
      observe(el: Element): void {
        if (this.#seen.size === 0) live += 1;
        this.#seen.add(el);
      }
      unobserve(el: Element): void {
        this.#seen.delete(el);
      }
      disconnect(): void {
        if (this.#seen.size > 0 || live > 0) live -= 1;
        this.#seen.clear();
      }
    }
    const w = win as unknown as { IntersectionObserver: unknown };
    const had = w.IntersectionObserver;
    w.IntersectionObserver = Observer;
    try {
      const editor = await open(40);
      const painter = new ThumbnailPainter({
        session: () => editor.session,
        idPrefix: 'q',
        frameLabel: (kind) => kind,
        layoutPreview: () => undefined,
      });
      // A rail that unmounts at once: its containers wait, and an observer watches them.
      const { list, boxes } = rail(40);
      const disposers = boxes.map((box, i) => painter.render(i, box, true));
      await endOfTask();
      expect(painter.waiting).toBe(true);
      expect(live).toBe(1);
      list.remove();
      for (const dispose of disposers) dispose();
      expect(painter.waiting).toBe(false);
      expect(live).toBe(0);

      // A rail that leaves the page with no release: its containers are held weakly and looked
      // for again; a release later stops that too.
      const other = rail(40);
      const later = other.boxes.map((box, i) => painter.render(i, box, true));
      await endOfTask();
      other.list.remove();
      await wait(50);
      expect(painter.waiting).toBe(true);
      for (const dispose of later) dispose();
      expect(painter.waiting).toBe(false);
      expect(live).toBe(0);
      painter.clear();
      editor.destroy();
    } finally {
      w.IntersectionObserver = had;
    }
  });

  test('a waiting thumbnail paints the slide it was asked for after slides move', async () => {
    const editor = await open(30);
    const { list, boxes } = rail(30);
    boxes.forEach((box, i) => editor.renderThumbnail(i, box, DEFER));
    // Box 20 waits for slide 21; slide 1 then moves to the end, with no new render asked for.
    expect(painted([boxes[20]!])).toBe(0);
    expect(editor.exec({ type: 'moveSlides', indices: [0], to: 29 }).ok).toBe(true);
    await until(() => painted([boxes[20]!]) === 1);
    expect(boxes[20]!.textContent).toContain('Slide 21 title');
    editor.destroy();
    list.remove();
  });

  test('a destroyed editor paints nothing that waited', async () => {
    const editor = await open(20);
    const { list, boxes } = rail(20);
    boxes.forEach((box, i) => editor.renderThumbnail(i, box, DEFER));
    expect(painted(boxes)).toBe(13);
    editor.destroy();
    await wait(50);
    expect(painted(boxes)).toBe(13);
    list.remove();
  });
});
