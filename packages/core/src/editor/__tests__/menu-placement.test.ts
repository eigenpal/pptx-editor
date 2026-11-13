// Where a context menu opens (`placeInViewport`) and where a slide menu is anchored
// (`slideMenuPoint`, `canvasMenuPoint`). The React and Vue context menus share them.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { canvasMenuPoint, placeInViewport } from '../menu-arrangement.ts';
import { slideMenuPoint } from '../pane-support.ts';

describe('context menu placement', () => {
  test('opens below and to the right of the point when it fits', () => {
    expect(placeInViewport({ x: 100, y: 120 }, 240, 300, 1000, 760)).toEqual({
      left: 100,
      top: 120,
    });
  });

  test('flips above and to the left at the bottom and right edges, with a corner at the point', () => {
    expect(placeInViewport({ x: 990, y: 700 }, 240, 300, 1000, 760)).toEqual({
      left: 750,
      top: 400,
    });
  });

  test('a menu that fits neither above nor below stays at the point and scrolls', () => {
    // A rail at the bottom of a short window: more room above, so the menu ends at the point.
    expect(placeInViewport({ x: 300, y: 380 }, 240, 420, 1000, 560)).toEqual({
      left: 300,
      top: 8,
      maxHeight: 372,
    });
    // More room below: the menu starts at the point.
    expect(placeInViewport({ x: 300, y: 100 }, 240, 900, 1000, 760)).toEqual({
      left: 300,
      top: 100,
      maxHeight: 652,
    });
  });

  test('a keyboard open anchors at the slide, a pointer open at the pointer', () => {
    const item = {
      getBoundingClientRect: () => ({ left: 40, top: 500, right: 160, bottom: 580 }),
    } as unknown as HTMLElement;
    expect(slideMenuPoint(item, { x: 77, y: 530 })).toEqual({ x: 77, y: 530 });
    expect(slideMenuPoint(item, null)).toEqual({ x: 52, y: 572 });
  });
});

describe('canvas menu point', () => {
  const win = new GlobalWindow();
  const doc = win.document as unknown as Document;
  afterAll(() => win.happyDOM.close());
  const rect = (left: number, top: number, width: number, height: number) =>
    ({ left, top, width, height, right: left + width, bottom: top + height }) as DOMRect;

  test('a keyboard open in edited text opens below the caret; else at the selection', () => {
    const owner = doc.createElement('div');
    const outline = doc.createElement('div');
    outline.className = 'pptx-surface__outline';
    outline.getBoundingClientRect = () => rect(100, 100, 200, 80);
    const input = doc.createElement('textarea');
    input.className = 'pptx-surface__input';
    // The input sits on the caret line (`placeInput`).
    input.getBoundingClientRect = () => rect(240.4, 130, 1, 24);
    owner.append(outline, input);
    doc.body.append(owner);
    // A pointer open stays at the pointer.
    expect(canvasMenuPoint(owner, { x: 5, y: 6 })).toEqual({ x: 5, y: 6 });
    // No text caret: the middle of the selection outline.
    expect(canvasMenuPoint(owner, null)).toEqual({ x: 200, y: 140 });
    input.focus();
    expect(canvasMenuPoint(owner, null, true)).toEqual({ x: 240, y: 154 });
    // A menu for another selection keeps to the outline, wherever the input is.
    expect(canvasMenuPoint(owner, null)).toEqual({ x: 200, y: 140 });
  });
});
