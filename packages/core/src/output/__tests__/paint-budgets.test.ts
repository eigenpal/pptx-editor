// The slide paint budgets: a fill reads and spends in one step, and only a kept record's share
// spends without a read.

import { expect, test } from 'bun:test';
import {
  MAX_SLIDE_TILE_IMAGES,
  PaintBudgets,
  ShapeBandBudget,
  TileBudget,
} from '../paint-budgets.ts';

test('take reads the tiles left and spends one image for each tile that fits', () => {
  const budget = new TileBudget();
  let seen = -1;
  const tiles = budget.take((left) => {
    seen = left;
    return [1, 2, 3];
  });
  expect(seen).toBe(MAX_SLIDE_TILE_IMAGES);
  expect(tiles).toEqual([1, 2, 3]);
  expect(budget.asks).toBe(1);
  expect(budget.spent).toBe(3);
  expect(budget.left).toBe(MAX_SLIDE_TILE_IMAGES - 3);
});

test('take counts a read that does not fit, and spends nothing', () => {
  const budget = new TileBudget();
  expect(budget.take(() => undefined)).toBeUndefined();
  expect(budget.asks).toBe(1);
  expect(budget.spent).toBe(0);
});

test('a fill cannot spend a budget without reading it', () => {
  expect('spend' in new TileBudget()).toBe(false);
  expect('spend' in new ShapeBandBudget()).toBe(false);
});

test('PaintBudgets.spend spends a kept share of both budgets without a read', () => {
  const budgets = new PaintBudgets();
  budgets.spend({ tiles: 10, bands: 256 });
  expect(budgets.spent).toEqual({ tiles: 10, bands: 256 });
  expect(budgets.asks).toEqual({ tiles: 0, bands: 0 });
});

test('left stops at zero, so paints past the limit read the same left', () => {
  const a = new PaintBudgets();
  const b = new PaintBudgets();
  a.spend({ tiles: MAX_SLIDE_TILE_IMAGES + 1, bands: 0 });
  b.spend({ tiles: MAX_SLIDE_TILE_IMAGES + 50, bands: 0 });
  expect(a.left.tiles).toBe(0);
  expect(a.left).toEqual(b.left);
  // A fill that reads it fits nothing in either.
  expect(a.tiles.take((left) => (left >= 1 ? [0] : undefined))).toBeUndefined();
  expect(b.tiles.take((left) => (left >= 1 ? [0] : undefined))).toBeUndefined();
});
