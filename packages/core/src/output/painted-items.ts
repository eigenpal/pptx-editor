// The record of one painted layout item (`PaintedItem`), and the budget accounting of one slide
// paint (`PaintMeter`). `SlidePainter` keeps these records between paints; this module decides
// what a record spent and read of the paint budgets, and whether it still fits them.

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

/** The result of `paintChildren`: the next child to paint, and the weight that it painted. */
export interface ChildPaint {
  readonly next: number;
  readonly spent: number;
}

/** A record and the paint weight that making it used of the frame's budget. */
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

/** The fields of a record that one measured paint gives. */
interface Measured<T> {
  readonly value: T;
  readonly defs: readonly SVGElement[];
  readonly spent: BudgetAmounts;
  readonly startLeft: BudgetAmounts;
  readonly reads: BudgetReads;
}

/**
 * The budgets of one slide paint, and the records of what each item paint spends and reads of
 * them. The budgets start whole for each paint and are spent in paint order: the background, then
 * each item, kept or new. Each item must spend its share when the walk reaches it (`keep`,
 * `extend`, or a new paint), so that `fits` compares with what is left at the item's own place.
 */
export class PaintMeter {
  readonly #defs: Defs;
  readonly #budgets: PaintBudgets;

  constructor(defs: Defs) {
    this.#defs = defs;
    this.#budgets = startPaintBudgets(defs);
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

  /** Keep `painted` as it is: spend what it spent. */
  keep(painted: PaintedItem): void {
    this.#budgets.spend(painted.spent);
  }

  /** Paint a whole item with `paint`, and record it. */
  whole(paint: () => SVGElement | null): PaintedItem {
    const fresh = this.#measure(paint);
    return {
      node: fresh.value,
      defs: fresh.defs,
      spent: fresh.spent,
      startLeft: fresh.startLeft,
      reads: fresh.reads,
    };
  }

  /**
   * Begin a group in `node`: `paint` paints its first children, as many as fit the frame. The
   * record is partial (`rest`) when some of its `count` children are left.
   */
  startGroup(node: SVGElement, count: number, paint: () => ChildPaint): MeteredPaint {
    const fresh = this.#measure(paint);
    const { next, spent } = fresh.value;
    const painted: PaintedItem = {
      node,
      defs: fresh.defs,
      spent: fresh.spent,
      startLeft: fresh.startLeft,
      reads: fresh.reads,
      ...(next < count ? { rest: next } : {}),
    };
    return { painted, weight: spent + 1 };
  }

  /**
   * Keep `painted`, and when it is a partial group of `count` children, continue it: `paint`
   * paints its children from `rest` into its node. Answers the new record, which joins the
   * earlier and the new nodes, spends, and reads. A budget that the record reads for the first
   * time starts at what was left when this item's turn began: no budget changes between that
   * point and the new read, because the record's earlier children did not read it.
   */
  extend(
    painted: PaintedItem,
    count: number,
    paint: (node: SVGElement, from: number) => ChildPaint
  ): MeteredPaint {
    const atTurn = this.#budgets.left;
    this.keep(painted);
    const { node, rest } = painted;
    if (rest === undefined || !node) return { painted, weight: 0 };
    const more = this.#measure(() => paint(node, rest));
    const { next, spent } = more.value;
    const extended: PaintedItem = {
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
    };
    return { painted: extended, weight: spent };
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
