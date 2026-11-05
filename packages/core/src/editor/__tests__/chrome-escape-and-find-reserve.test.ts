// Two chrome focus and placement rules the adapters share: Escape in a menu bar menu leaves
// focus on its menubar item (`menuBarClosedOnEscape`), and an open find bar keeps its height
// free above the slide (`reserveFindBarSpace`).

import { describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { menuBarClosedOnEscape, reserveFindBarSpace } from '../index.ts';

// A DOM of this file's own: the core run has no global DOM (`bun run test`).
const win = new GlobalWindow();
const document = win.document as unknown as Document;

describe('menuBarClosedOnEscape', () => {
  test('true only while the trigger holds focus, as the popover leaves it after Escape', () => {
    const trigger = document.createElement('button');
    const other = document.createElement('button');
    document.body.append(trigger, other);
    other.focus();
    expect(menuBarClosedOnEscape(trigger)).toBe(false);
    trigger.focus();
    expect(menuBarClosedOnEscape(trigger)).toBe(true);
    trigger.remove();
    expect(menuBarClosedOnEscape(trigger)).toBe(false);
    expect(menuBarClosedOnEscape(null)).toBe(false);
  });
});

describe('reserveFindBarSpace', () => {
  test("writes the bar's height on its parent, 0 while hidden, and removes it on cleanup", () => {
    let resized: (() => void) | null = null;
    const view = win as unknown as { ResizeObserver: unknown; requestAnimationFrame: unknown };
    const saved = view.ResizeObserver;
    const savedFrame = view.requestAnimationFrame;
    // The write waits for the next frame; here the frame comes at once.
    view.requestAnimationFrame = (run: (time: number) => void) => (run(0), 0);
    view.ResizeObserver = class {
      constructor(callback: () => void) {
        resized = callback;
      }
      observe(): void {}
      disconnect(): void {}
    };
    try {
      const parent = document.createElement('div');
      const bar = document.createElement('div');
      bar.hidden = true;
      parent.append(bar);
      document.body.append(parent);
      let height = 0;
      Object.defineProperty(bar, 'offsetHeight', { get: () => height });
      const stop = reserveFindBarSpace(bar);
      expect(parent.style.getPropertyValue('--pptx-find-reserve')).toBe('0px');
      bar.hidden = false;
      height = 71;
      resized!();
      expect(parent.style.getPropertyValue('--pptx-find-reserve')).toBe('71px');
      bar.hidden = true;
      resized!();
      expect(parent.style.getPropertyValue('--pptx-find-reserve')).toBe('0px');
      stop();
      expect(parent.style.getPropertyValue('--pptx-find-reserve')).toBe('');
    } finally {
      view.ResizeObserver = saved;
      view.requestAnimationFrame = savedFrame;
    }
  });

  test('a bar in the canvas margin does not move the slide; one over the slide does, once', () => {
    let resized: (() => void) | null = null;
    const view = win as unknown as { ResizeObserver: unknown; requestAnimationFrame: unknown };
    const saved = view.ResizeObserver;
    const savedFrame = view.requestAnimationFrame;
    view.requestAnimationFrame = (run: (time: number) => void) => (run(0), 0);
    view.ResizeObserver = class {
      constructor(callback: () => void) {
        resized = callback;
      }
      observe(): void {}
      unobserve(): void {}
      disconnect(): void {}
    };
    const box = (el: HTMLElement, rect: () => [number, number, number, number]): void => {
      el.getBoundingClientRect = () => {
        const [left, top, right, bottom] = rect();
        return { left, top, right, bottom, width: right - left, height: bottom - top } as DOMRect;
      };
    };
    try {
      const parent = document.createElement('div');
      const stage = document.createElement('div');
      stage.className = 'pptx-surface__stage';
      const bar = document.createElement('div');
      bar.hidden = true;
      parent.append(stage, bar);
      document.body.append(parent);
      Object.defineProperty(bar, 'offsetHeight', { get: () => 55 });
      // The bar sits at the top right; the fitted slide starts 80 px down.
      box(bar, () => [600, 12, 980, 67]);
      let slideTop = 80;
      box(stage, () => [100, slideTop, 900, slideTop + 450]);
      const stop = reserveFindBarSpace(bar);
      bar.hidden = false;
      resized!();
      expect(parent.style.getPropertyValue('--pptx-find-reserve')).toBe('0px');
      // A zoom brings the slide under the bar: the bar reserves its height.
      slideTop = 20;
      resized!();
      expect(parent.style.getPropertyValue('--pptx-find-reserve')).toBe('55px');
      // The reserve moved the slide clear; the bar keeps the room until it closes.
      slideTop = 80;
      resized!();
      expect(parent.style.getPropertyValue('--pptx-find-reserve')).toBe('55px');
      bar.hidden = true;
      resized!();
      expect(parent.style.getPropertyValue('--pptx-find-reserve')).toBe('0px');
      // Opened again over a slide in the clear: nothing moves.
      bar.hidden = false;
      resized!();
      expect(parent.style.getPropertyValue('--pptx-find-reserve')).toBe('0px');
      stop();
    } finally {
      view.ResizeObserver = saved;
      view.requestAnimationFrame = savedFrame;
    }
  });

  test('a bar with no parent reserves nothing', () => {
    expect(typeof reserveFindBarSpace(null)).toBe('function');
    expect(typeof reserveFindBarSpace(document.createElement('div'))).toBe('function');
  });
});
