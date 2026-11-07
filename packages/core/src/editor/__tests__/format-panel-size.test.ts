// The Format panel's size rule: a locked aspect ratio sizes both sides from the one typed.

import { describe, expect, test } from 'bun:test';
import { formatPanelSize } from '../canvas-support.ts';

describe('formatPanelSize', () => {
  test('unlocked, only the typed side changes', () => {
    expect(formatPanelSize({ cx: 200, cy: 100 }, 'cx', 400)).toEqual({ cx: 400, cy: 100 });
    expect(formatPanelSize({ cx: 200, cy: 100, lockAspectRatio: false }, 'cy', 50)).toEqual({
      cx: 200,
      cy: 50,
    });
  });

  test('locked, the other side keeps the ratio, rounded to whole EMU', () => {
    const g = { cx: 300, cy: 100, lockAspectRatio: true };
    expect(formatPanelSize(g, 'cx', 600)).toEqual({ cx: 600, cy: 200 });
    expect(formatPanelSize(g, 'cy', 50)).toEqual({ cx: 150, cy: 50 });
    expect(formatPanelSize(g, 'cx', 100)).toEqual({ cx: 100, cy: 33 });
  });

  test('a side of 0 has no ratio, and a negative value is 0', () => {
    expect(formatPanelSize({ cx: 0, cy: 100, lockAspectRatio: true }, 'cx', 50)).toEqual({
      cx: 50,
      cy: 100,
    });
    expect(formatPanelSize({ cx: 200, cy: 100, lockAspectRatio: true }, 'cx', -5)).toEqual({
      cx: 0,
      cy: 0,
    });
  });
});
