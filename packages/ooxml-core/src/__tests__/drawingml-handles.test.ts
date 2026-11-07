import { describe, expect, test } from 'bun:test';
import { adjustForHandle, presetAdjustments, presetHandles, presetNames } from '../index.ts';

const W = 2_000_000;
const H = 1_000_000;

describe('presetHandles', () => {
  test('places a handle by its own formulas', () => {
    // roundRect: the handle is at x1 = ss * adj / 100000 on the top edge.
    expect(presetHandles('roundRect', W, H, new Map())).toEqual([
      { kind: 'xy', x: 166_670, y: 0, guides: ['adj'] },
    ]);
    expect(presetHandles('roundRect', W, H, new Map([['adj', 50_000]]))[0]!.x).toBe(500_000);
  });

  test('a mirrored preset mirrors its handles, and a shape of no size has none', () => {
    const down = presetHandles('downArrow', W, H, new Map());
    const up = presetHandles('upArrow', W, H, new Map());
    expect(up.map((h) => [h.x, H - h.y])).toEqual(down.map((h) => [h.x, h.y]));
    expect(presetHandles('roundRect', 0, H, new Map())).toEqual([]);
    expect(presetHandles('rect', W, H, new Map())).toEqual([]);
  });
});

describe('adjustForHandle', () => {
  test('every handle of every preset reaches a place inside its range', () => {
    const far: string[] = [];
    for (const preset of presetNames()) {
      presetHandles(preset, W, H, new Map()).forEach((handle, i) => {
        const slots = presetAdjustments(preset, W, H, new Map())!;
        const target = new Map<string, number>();
        for (const g of handle.guides) {
          const s = slots.find((a) => a.name === g)!;
          const lo = s.min ?? 0;
          const hi = s.max ?? (s.angle ? 21_599_999 : 100_000);
          target.set(g, Math.round(lo + (hi - lo) * 0.3));
        }
        const goal = presetHandles(preset, W, H, target)[i]!;
        const values = adjustForHandle(preset, W, H, new Map(), i, goal)!;
        const got = presetHandles(preset, W, H, values)[i]!;
        if (Math.hypot(got.x - goal.x, got.y - goal.y) > 2_000) far.push(`${preset} #${i}`);
      });
    }
    expect(far).toEqual([]);
  });

  test('a point past the range gives the end of the range', () => {
    const values = adjustForHandle('roundRect', W, H, new Map(), 0, { x: 3 * W, y: 0 });
    expect(values).toEqual(new Map([['adj', 50_000]]));
    expect(adjustForHandle('rect', W, H, new Map(), 0, { x: 0, y: 0 })).toBeUndefined();
  });
});
