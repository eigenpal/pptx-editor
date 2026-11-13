// The budgets of one slide paint: the tile images of tiled pictures (`TileBudget`) and the band
// work of shape gradients (`ShapeBandBudget`). Each bounds what one paint of a slide writes. A
// painter that keeps its `Defs` across paints (`SlidePainter`) starts new budgets for each paint,
// so the paints and edits before it do not use them up.

import { ShapeBandBudget } from './paint-gradient.ts';
import type { Defs } from './svg.ts';

/**
 * The most tile images that one paint of a slide draws (`TileBudget`). A slide with more tiles
 * fills the rest with the pattern of `tilePaint`, so a file cannot make the paint write millions
 * of elements, with many shapes or many filled paths.
 */
export const MAX_SLIDE_TILE_IMAGES = 4096;

/** The tile images left in one paint of a slide. */
export class TileBudget {
  #spent = 0;
  #asks = 0;
  /** How many tiled fills read the budget: a paint that reads none draws the same whatever is left. */
  get asks(): number {
    return this.#asks;
  }
  /** The images left, for a tiled fill that fits its tiles to them. */
  ask(): number {
    this.#asks += 1;
    return this.left;
  }
  /** The images still allowed, never below zero. */
  get left(): number {
    return Math.max(0, MAX_SLIDE_TILE_IMAGES - this.#spent);
  }
  /** The images spent so far, past the budget too. */
  get spent(): number {
    return this.#spent;
  }
  spend(n: number): void {
    this.#spent += n;
  }
}

/** What a paint spent of each budget. */
export interface BudgetSpend {
  readonly tiles: number;
  readonly bands: number;
}

/** The budgets of one paint of a slide. */
export class PaintBudgets {
  readonly tiles = new TileBudget();
  readonly bands = new ShapeBandBudget();
  /** How many fills read a budget. */
  get asks(): number {
    return this.tiles.asks + this.bands.asks;
  }
  /** What is left of each budget: a paint that reads them depends on nothing else of them. */
  get state(): string {
    return `${this.tiles.left} ${this.bands.left}`;
  }
  /** What is spent of each budget so far. */
  get spent(): BudgetSpend {
    return { tiles: this.tiles.spent, bands: this.bands.spent };
  }
  spend(s: BudgetSpend): void {
    this.tiles.spend(s.tiles);
    this.bands.spend(s.bands);
  }
}

/** The budgets of the slide paint in progress, by the painter's `Defs`. */
const work = new WeakMap<Defs, PaintBudgets>();

/**
 * The budgets of the paint in progress with `defs`: a fresh `paintSlide` makes new `Defs`, so its
 * first tiled picture or shape gradient makes them.
 */
export function paintBudgets(defs: Defs): PaintBudgets {
  let budgets = work.get(defs);
  if (!budgets) work.set(defs, (budgets = new PaintBudgets()));
  return budgets;
}

/**
 * Start a new slide paint with `defs`, with whole budgets. A painter that keeps its `Defs` across
 * paints (`SlidePainter`) calls this when each paint starts, so each paint of a slide spends the
 * same budgets and draws the same tiles and bands.
 */
export function startPaintBudgets(defs: Defs): PaintBudgets {
  const budgets = new PaintBudgets();
  work.set(defs, budgets);
  return budgets;
}
