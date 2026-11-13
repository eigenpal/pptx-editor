// The record of one painted layout item (`PaintedItem`), and the budget accounting of one slide
// paint (`PaintMeter`). `SlidePainter` keeps these records between paints; this module decides
// what a record spent and read of the paint budgets, and whether it still fits them.

import type { ChildPaint } from './paint.ts';
import { startPaintBudgets, type BudgetAmounts, type PaintBudgets } from './paint-budgets.ts';
import type { Defs } from './svg.ts';

/** Which budgets a paint read: a fill read a budget when what it drew depended on what was left. */
export interface BudgetReads {
  readonly tiles: boolean;
  readonly bands: boolean;
}

/**
 * The nodes that one layout item painted, and what that paint used of the slide's budgets
 * (`PaintBudgets`). A record never changes. When a paint continues a partial group, it makes a new
 * record (`PaintMeter.extend`), so `spent`, `reads`, and `startLeft` always change together.
 */
export interface PaintedItem {
  /** The item's node in the slide's content group, or null when the item paints nothing. */
  readonly node: SVGElement | null;
  /** The definitions that the paint added to the slide's `<defs>`. */
  readonly defs: readonly SVGElement[];
  /** What the paint spent of each budget. A later paint that keeps the node spends the same. */
  readonly spent: BudgetAmounts;
  /** Which budgets the paint read. A budget that it did not read cannot change its nodes. */
  readonly reads: BudgetReads;
  /**
   * For each budget that the paint read: what was left of it when the paint began to read it.
   * The nodes stay correct while that much is left at the item's place in paint order
   * (`PaintMeter.fits`). For a budget that the paint did not read, the value is not used.
   */
  readonly startLeft: BudgetAmounts;
  /**
   * Set only for a partial paint of a group: the index of the first child that is not painted
   * yet. A later paint continues from this child.
   */
  readonly rest?: number;
}

/** A record and the paint weight (`paintWeight`) that making it used of the frame. */
export interface MeteredPaint {
  readonly painted: PaintedItem;
  readonly weight: number;
}

/** Remove the nodes and definitions of `painted` from the slide. */
export function dropPainted(painted: PaintedItem): void {
  painted.node?.remove();
  for (const d of painted.defs) d.remove();
}

const add = (a: BudgetAmounts, b: BudgetAmounts): BudgetAmounts => ({
  tiles: a.tiles + b.tiles,
  bands: a.bands + b.bands,
});

/** What one measured paint gives: its value, its definitions, and its budget use. */
interface Measured<T> {
  readonly value: T;
  readonly defs: readonly SVGElement[];
  readonly spent: BudgetAmounts;
  readonly startLeft: BudgetAmounts;
  readonly reads: BudgetReads;
}

/**
 * The budgets of one slide paint, and the records of what each item paint spends and reads of
 * them. The budgets start whole for each paint (`resetBudgets`) and are spent in paint order: the
 * background, then each item on screen. Each record on screen spends its share once, when the walk
 * reaches it (`spendShare`), so that `fits` compares with what is left at the item's own place.
 *
 * The accounting depends on fills that read a budget before they spend from it: a paint that read
 * nothing spent nothing. `TileBudget.take` and `ShapeBandBudget.allow` read and spend in one step,
 * so a fill cannot do otherwise.
 */
export class PaintMeter {
  readonly #defs: Defs;
  readonly #budgets: PaintBudgets;
  /**
   * The records whose share this paint has spent: the ones it painted or extended, which spent
   * while they painted, and the ones that `spendShare` kept. A record spends at most once.
   */
  readonly #counted = new WeakSet<PaintedItem>();

  private constructor(defs: Defs) {
    this.#defs = defs;
    this.#budgets = startPaintBudgets(defs);
  }

  /**
   * Start one slide paint with `defs`. Side effect: this replaces the budgets that
   * `paintBudgets(defs)` returns with whole ones, so the fills of this paint spend from them and
   * nothing that an earlier paint spent counts. It is the only way to make a meter, so a meter
   * always measures the budgets that its paint's fills spend.
   */
  static resetBudgets(defs: Defs): PaintMeter {
    return new PaintMeter(defs);
  }

  /**
   * Whether the nodes of `painted` are what a new paint draws at this place in paint order: each
   * budget that it read has as much left now as when it began to read it.
   */
  fits({ reads, startLeft }: PaintedItem): boolean {
    const left = this.#budgets.left;
    return (
      (!reads.tiles || startLeft.tiles === left.tiles) &&
      (!reads.bands || startLeft.bands === left.bands)
    );
  }

  /**
   * Spend the share of `painted` at this place in paint order, unless this paint already did: a
   * record that this paint made or extended spent while it painted.
   */
  spendShare(painted: PaintedItem): void {
    if (this.#counted.has(painted)) return;
    this.#budgets.spend(painted.spent);
    this.#counted.add(painted);
  }

  /** Paint a whole item with `paint`, and answer its record. */
  recordWhole(paint: () => SVGElement | null): PaintedItem {
    const fresh = this.#measure(paint);
    return this.#counts({
      node: fresh.value,
      defs: fresh.defs,
      spent: fresh.spent,
      startLeft: fresh.startLeft,
      reads: fresh.reads,
    });
  }

  /**
   * Begin a group of `count` children in its new `node`: `paint` paints its first children, as
   * many as fit the frame. The record is partial (`rest`) when some children are left. The paint
   * weight counts the group node as one.
   */
  recordGroupStart(node: SVGElement, count: number, paint: () => ChildPaint): MeteredPaint {
    const fresh = this.#measure(paint);
    const { next, weight } = fresh.value;
    const painted = this.#counts({
      node,
      defs: fresh.defs,
      spent: fresh.spent,
      startLeft: fresh.startLeft,
      reads: fresh.reads,
      ...(next < count ? { rest: next } : {}),
    });
    return { painted, weight: weight + 1 };
  }

  /**
   * Spend the share of `painted`, and when it is a partial group of `count` children, continue
   * it: `paint` paints its children from `rest` into its node. Answers the new record, which joins
   * the earlier and the new nodes, spends, and reads. A budget that the record reads for the
   * first time starts at what was left when this item's turn began. Nothing spent that budget
   * between that point and the new read, because only a fill that reads a budget spends it.
   */
  extend(
    painted: PaintedItem,
    count: number,
    paint: (node: SVGElement, from: number) => ChildPaint
  ): MeteredPaint {
    const atTurn = this.#budgets.left;
    this.spendShare(painted);
    const { node, rest } = painted;
    if (rest === undefined || !node) return { painted, weight: 0 };
    const more = this.#measure(() => paint(node, rest));
    const { next, weight } = more.value;
    const extended = this.#counts({
      node,
      defs: [...painted.defs, ...more.defs],
      spent: add(painted.spent, more.spent),
      startLeft: {
        tiles: painted.reads.tiles ? painted.startLeft.tiles : atTurn.tiles,
        bands: painted.reads.bands ? painted.startLeft.bands : atTurn.bands,
      },
      reads: {
        tiles: painted.reads.tiles || more.reads.tiles,
        bands: painted.reads.bands || more.reads.bands,
      },
      ...(next < count ? { rest: next } : {}),
    });
    return { painted: extended, weight };
  }

  /** Note that `painted` spent its share while it painted, and answer it. */
  #counts(painted: PaintedItem): PaintedItem {
    this.#counted.add(painted);
    return painted;
  }

  /** Run `paint`, and count what it spent of the budgets, which it read, and its definitions. */
  #measure<T>(paint: () => T): Measured<T> {
    const budgets = this.#budgets;
    const [from, asks, startLeft] = [budgets.spent, budgets.asks, budgets.left];
    const fresh = this.#defs.capture(paint);
    const [to, now] = [budgets.spent, budgets.asks];
    return {
      value: fresh.value,
      defs: [...fresh.defs],
      spent: { tiles: to.tiles - from.tiles, bands: to.bands - from.bands },
      startLeft,
      reads: { tiles: now.tiles > asks.tiles, bands: now.bands > asks.bands },
    };
  }
}
