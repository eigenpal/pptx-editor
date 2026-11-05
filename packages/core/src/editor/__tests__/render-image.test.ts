// The size of a slide or shape picture: fit into the asked box, keeping the aspect ratio.

import { describe, expect, test } from 'bun:test';
import { fitSize } from '../render-image.ts';

describe('picture size', () => {
  const W = 1280;
  const H = 720;
  test('no size gives the own size', () => {
    expect(fitSize(W, H, {})).toEqual([1280, 720]);
  });
  test('one size sets the scale and keeps the aspect ratio', () => {
    expect(fitSize(W, H, { width: 640 })).toEqual([640, 360]);
    expect(fitSize(W, H, { height: 180 })).toEqual([320, 180]);
  });
  test('both sizes fit the picture into the box', () => {
    expect(fitSize(W, H, { width: 640, height: 640 })).toEqual([640, 360]);
    expect(fitSize(W, H, { width: 2000, height: 360 })).toEqual([640, 360]);
  });
  test('a size larger than the renderer draws is clamped', () => {
    expect(fitSize(W, H, { width: 100_000 })).toEqual([8192, 4608]);
  });
  test('a tiny size keeps at least one pixel', () => {
    expect(fitSize(W, H, { width: 0.5 })).toEqual([1, 1]);
  });
});
