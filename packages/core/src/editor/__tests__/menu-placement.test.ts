// Where a context menu opens (`placeInViewport`) and where a slide menu is anchored
// (`slideMenuPoint`). The React and Vue context menus share both.

import { describe, expect, test } from 'bun:test';
import { placeInViewport } from '../menu-arrangement.ts';
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
