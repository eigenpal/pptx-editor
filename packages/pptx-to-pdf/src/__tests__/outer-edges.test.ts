/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The outer edge of the union of a shape's filled paths, for soft-edge and blur fades.

import { describe, expect, test } from 'bun:test';
import type { PathCommand } from '@pptx-editor.dev/core/layout';
import { outerEdges } from '../outer-edges.ts';

const PT = 12_700;

const square = (x: number, y: number, side: number): PathCommand[] => [
  { op: 'M', x, y },
  { op: 'L', x: x + side, y },
  { op: 'L', x: x + side, y: y + side },
  { op: 'L', x, y: y + side },
  { op: 'Z' },
];

const points = (paths: PathCommand[][] | undefined): { x: number; y: number }[] =>
  (paths ?? []).flat().flatMap((c) => (c.op === 'Z' ? [] : [{ x: c.x, y: c.y }]));

describe('outer edges', () => {
  test('two overlapping squares stroke only the outer edge of their union', () => {
    const a = square(0, 0, 100 * PT);
    const b = square(50 * PT, 50 * PT, 100 * PT);
    const edges = outerEdges([a, b]);
    expect(edges).toHaveLength(2);
    const pts = points(edges);
    expect(pts.length).toBeGreaterThan(0);
    // No stroked point lies inside the other square, away from the outer edge.
    const inner = (q: { x: number; y: number }, x0: number, y0: number): boolean =>
      q.x > x0 + PT && q.x < x0 + 99 * PT && q.y > y0 + PT && q.y < y0 + 99 * PT;
    for (const q of pts) {
      expect(inner(q, 0, 0)).toBe(false);
      expect(inner(q, 50 * PT, 50 * PT)).toBe(false);
    }
    // The outer corners of the union are stroked.
    for (const [x, y] of [
      [0, 0],
      [150 * PT, 150 * PT],
    ])
      expect(pts.some((q) => Math.hypot(q.x - x, q.y - y) < 2 * PT)).toBe(true);
  });

  test('a square inside another strokes nothing', () => {
    const edges = outerEdges([square(0, 0, 100 * PT), square(20 * PT, 20 * PT, 10 * PT)]) ?? [];
    expect(edges[1]).toHaveLength(0);
    expect(edges[0].length).toBeGreaterThan(0);
  });

  test('40,000 separate small paths give up the search in bounded time', () => {
    const filled = Array.from({ length: 40_000 }, (_, i) =>
      square((i % 200) * 3 * PT, Math.floor(i / 200) * 3 * PT, PT)
    );
    const start = performance.now();
    const edges = outerEdges(filled);
    const elapsed = performance.now() - start;
    expect(elapsed).toBeLessThan(200);
    // Past the search, the caller draws a hard edge rather than fade edges inside the union.
    expect(edges).toBeUndefined();
  });
});
