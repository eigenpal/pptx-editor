// The default editor keeps its floating corner parts above the notes card (`notes-inset.ts`).

import { describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { NOTES_INSET_PROPERTY, trackNotesInset } from '../index.ts';

// A DOM of this file's own: the core run has no global DOM (`bun run test`).
const win = new GlobalWindow();
const document = win.document as unknown as Document;

/** Give `el` a box from `top` to `bottom`. */
function place(el: HTMLElement, top: number, bottom: number): void {
  el.getBoundingClientRect = () =>
    ({ top, bottom, left: 0, right: 100, width: 100, height: bottom - top }) as DOMRect;
}

/** Run `body` with a ResizeObserver whose callback the test calls, and the observed elements. */
function withObserver(body: (resize: () => void, observed: Set<Element>) => void): void {
  const view = win as unknown as { ResizeObserver: unknown };
  const saved = view.ResizeObserver;
  let callback: (() => void) | null = null;
  const observed = new Set<Element>();
  view.ResizeObserver = class {
    constructor(run: () => void) {
      callback = run;
    }
    observe(el: Element): void {
      observed.add(el);
    }
    unobserve(el: Element): void {
      observed.delete(el);
    }
    disconnect(): void {
      observed.clear();
    }
  };
  try {
    body(() => callback!(), observed);
  } finally {
    view.ResizeObserver = saved;
  }
}

describe('trackNotesInset', () => {
  test('writes the height the notes card takes from the bottom, and follows its size', () => {
    withObserver((resize, observed) => {
      const box = document.createElement('div');
      const notes = document.createElement('section');
      notes.className = 'pptx-notes';
      box.append(notes);
      document.body.append(box);
      place(box, 0, 800);
      place(notes, 648, 788);
      const tracker = trackNotesInset(box);
      expect(box.style.getPropertyValue(NOTES_INSET_PROPERTY)).toBe('152px');
      expect(observed.has(box) && observed.has(notes)).toBe(true);
      // The notes grow: the parts move up with the card.
      place(notes, 500, 788);
      resize();
      expect(box.style.getPropertyValue(NOTES_INSET_PROPERTY)).toBe('300px');
      // The notes close to their head.
      place(notes, 752, 788);
      resize();
      expect(box.style.getPropertyValue(NOTES_INSET_PROPERTY)).toBe('48px');
      tracker.dispose();
      expect(box.style.getPropertyValue(NOTES_INSET_PROPERTY)).toBe('');
      expect(observed.size).toBe(0);
      box.remove();
    });
  });

  test('a view with no notes card, or a card that is not rendered, takes nothing', () => {
    withObserver((_resize, observed) => {
      const box = document.createElement('div');
      document.body.append(box);
      place(box, 0, 800);
      const tracker = trackNotesInset(box);
      expect(box.style.getPropertyValue(NOTES_INSET_PROPERTY)).toBe('0px');
      // The Normal view comes back with its notes: a refresh finds the card.
      const notes = document.createElement('section');
      notes.className = 'pptx-notes';
      place(notes, 700, 788);
      box.append(notes);
      tracker.refresh();
      expect(box.style.getPropertyValue(NOTES_INSET_PROPERTY)).toBe('100px');
      expect(observed.has(notes)).toBe(true);
      // `display: none` gives an empty box at the top: that is no card, not a tall one.
      notes.getBoundingClientRect = () =>
        ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 }) as DOMRect;
      tracker.refresh();
      expect(box.style.getPropertyValue(NOTES_INSET_PROPERTY)).toBe('0px');
      // The Slide Sorter view removes the card.
      notes.remove();
      tracker.refresh();
      expect(box.style.getPropertyValue(NOTES_INSET_PROPERTY)).toBe('0px');
      expect(observed.has(notes)).toBe(false);
      tracker.dispose();
      box.remove();
    });
  });

  test('in a scaled container the inset counts in the box pixels, not the viewport pixels', () => {
    withObserver(() => {
      const box = document.createElement('div');
      const notes = document.createElement('section');
      notes.className = 'pptx-notes';
      box.append(notes);
      document.body.append(box);
      // `transform: scale(0.75)`: the 800px box shows as 600 viewport pixels.
      Object.defineProperty(box, 'offsetHeight', { configurable: true, value: 800 });
      place(box, 0, 600);
      place(notes, 486, 591);
      const tracker = trackNotesInset(box);
      expect(box.style.getPropertyValue(NOTES_INSET_PROPERTY)).toBe('152px');
      tracker.dispose();
      box.remove();
    });
  });

  test('without ResizeObserver it still measures on refresh', () => {
    const view = win as unknown as { ResizeObserver: unknown };
    const saved = view.ResizeObserver;
    view.ResizeObserver = undefined;
    try {
      const box = document.createElement('div');
      const notes = document.createElement('section');
      notes.className = 'pptx-notes';
      box.append(notes);
      document.body.append(box);
      place(box, 0, 600);
      place(notes, 500, 590);
      const tracker = trackNotesInset(box);
      expect(box.style.getPropertyValue(NOTES_INSET_PROPERTY)).toBe('100px');
      tracker.dispose();
      box.remove();
    } finally {
      view.ResizeObserver = saved;
    }
  });
});
