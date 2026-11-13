// The budgets of one slide paint: the tile images of tiled pictures (`TileBudget`) and the band
// work of shape gradients (`ShapeBandBudget`). Each bounds what one paint of a slide writes. This
// module keeps the budgets of the paint in progress for each `Defs` (`paintBudgets`). A painter that
// keeps its `Defs` across paints (`SlidePainter`) starts new budgets for each paint
// (`startPaintBudgets`), so the paints and edits before it do not use them up.

import { MAX_SHAPE_BANDS } from './paint-gradient.ts';
import type { Defs } from './svg.ts';

/**
 * The most tile images that one paint of a slide draws (`TileBudget`). A slide with more tiles
 * fills the rest with the pattern of `tilePaint`, so a file cannot make the paint write millions
 * of elements, with many shapes or many filled paths.
 */
export const MAX_SLIDE_TILE_IMAGES = 4096;

/**
 * The private way to spend a budget again: only `PaintBudgets.spend` uses it, to spend the share
 * of a record that a later paint keeps. A fill spends only through `take` or `allow`, which read
 * the budget first, so no fill can spend without reading.
 */
const replay = Symbol('replay');

/**
 * Why `left` stops at zero, and why that keeps `PaintMeter.fits` safe: a fill reads only `left`,
 * so two paints with the same `left` draw the same, however far past the limit either spent.
 * Equal `left` then means an equal drawing, which is what `fits` tests.
 */
const leftOf = (limit: number, spent: number): number => Math.max(0, limit - spent);

/** The tile images left in one paint of a slide. */
export class TileBudget {
  #spent = 0;
  #asks = 0;
  /** How many tiled fills read the budget: a paint that reads none draws the same whatever is left. */
  get asks(): number {
    return this.#asks;
  }
  /**
   * Read the budget and spend from it in one step. `fit` gets the images left and answers the
   * tiles that fit in them, or undefined when they do not fit. This spends one image for each
   * tile, and answers what `fit` answered.
   */
  take<T extends { readonly length: number }>(fit: (left: number) => T | undefined): T | undefined {
    this.#asks += 1;
    const tiles = fit(this.left);
    if (tiles) this.#spent += tiles.length;
    return tiles;
  }
  /** The images still allowed, never below zero (`leftOf`). */
  get left(): number {
    return leftOf(MAX_SLIDE_TILE_IMAGES, this.#spent);
  }
  /** The images spent so far, past the limit too. */
  get spent(): number {
    return this.#spent;
  }
  /** Spend `n` again for a kept record (`PaintBudgets.spend`). */
  [replay](n: number): void {
    this.#spent += n;
  }
}

/**
 * The most paths of one shape that draw a `path="shape"` gradient (fill or overlay) as outline
 * bands; a shape with more paths uses the rectangle construction, so a file cannot multiply the
 * bands.
 */
const MAX_BANDED_PATHS = 16;
/** The most path commands of one shape that draw outline bands: bands repeat every command. */
const MAX_BANDED_COMMANDS = 2000;
/**
 * The band work (path commands × `MAX_SHAPE_BANDS`) one slide may spend. Later shapes use the
 * rectangle construction, so many shapes cannot multiply the bands either.
 */
const MAX_SLIDE_BAND_WORK = 512_000;

/**
 * The shape gradient band work of one slide. A shape whose paths take `paths` and `commands`
 * draws outline bands when `allow` returns true; otherwise it uses the rectangle construction.
 * The SVG paint and `pptx-to-pdf` share the limits.
 */
export class ShapeBandBudget {
  #spent = 0;
  #asks = 0;
  /** How many shapes read the budget (`allow`) where the answer depended on what is left. */
  get asks(): number {
    return this.#asks;
  }
  /** The band work still allowed, never below zero (`leftOf`). */
  get left(): number {
    return leftOf(MAX_SLIDE_BAND_WORK, this.#spent);
  }
  /** The band work spent so far, past the limit too. */
  get spent(): number {
    return this.#spent;
  }
  /** Spend `n` again for a kept record (`PaintBudgets.spend`). */
  [replay](n: number): void {
    this.#spent += n;
  }
  /**
   * Read the budget and spend from it in one step: whether a shape of `paths` and `commands`
   * draws outline bands. When it does, this spends its band work.
   */
  allow(paths: number, commands: number): boolean {
    // A shape past the limits of one shape uses the rectangle construction, whatever is left.
    if (paths > MAX_BANDED_PATHS || commands > MAX_BANDED_COMMANDS) return false;
    this.#asks += 1;
    const cost = commands * MAX_SHAPE_BANDS;
    const ok = cost <= this.left;
    if (ok) this.#spent += cost;
    return ok;
  }
}

/** An amount of each budget: spent, or left. */
export interface BudgetAmounts {
  readonly tiles: number;
  readonly bands: number;
}

/** The budgets of one paint of a slide. */
export class PaintBudgets {
  readonly tiles = new TileBudget();
  readonly bands = new ShapeBandBudget();
  /**
   * How many fills read each budget, where the answer depended on it: a paint that reads a
   * budget depends on nothing else of it than what is left.
   */
  get asks(): BudgetAmounts {
    return { tiles: this.tiles.asks, bands: this.bands.asks };
  }
  /** What is left of each budget. */
  get left(): BudgetAmounts {
    return { tiles: this.tiles.left, bands: this.bands.left };
  }
  /** What is spent of each budget so far. */
  get spent(): BudgetAmounts {
    return { tiles: this.tiles.spent, bands: this.bands.spent };
  }
  /**
   * Spend `s` again: the share of a record that a later paint keeps (`PaintMeter.spendShare`).
   * This is the only spend that does not read, because the record read when it painted.
   */
  spend(s: BudgetAmounts): void {
    this.tiles[replay](s.tiles);
    this.bands[replay](s.bands);
  }
}

/**
 * The budgets of the slide paint in progress, by the `Defs` of that paint. This module alone reads
 * and writes it, so `Defs` stays a plain SVG helper. A `WeakMap` lets the budgets go with their
 * `Defs`.
 */
const work = new WeakMap<Defs, PaintBudgets>();

/**
 * The budgets of the slide paint in progress with `defs`. The fills that spend a budget find it
 * here. A fresh `paintSlide` makes new `Defs`, so its first tiled picture or shape gradient makes
 * whole budgets.
 */
export function paintBudgets(defs: Defs): PaintBudgets {
  let budgets = work.get(defs);
  if (!budgets) work.set(defs, (budgets = new PaintBudgets()));
  return budgets;
}

/**
 * Replace the budgets of `defs` with whole ones, and answer them. A painter that keeps its `Defs`
 * across paints (`SlidePainter`, through `PaintMeter`) calls this when each paint starts, so each
 * paint of a slide spends the same budgets and draws the same tiles and bands.
 */
export function startPaintBudgets(defs: Defs): PaintBudgets {
  const budgets = new PaintBudgets();
  work.set(defs, budgets);
  return budgets;
}
