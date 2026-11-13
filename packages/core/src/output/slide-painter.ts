// The painter of one view (the editing surface, a thumbnail): it keeps the nodes it painted
// and paints again only the items that changed, within a budget for each frame. Split from
// `paint.ts`, whose `paintItem` paints each item for both.

import type { LayoutItem, SlideLayout } from '../layout/index.ts';
import {
  groupNode,
  paintBackground,
  paintChildren,
  paintItem,
  paintWeight,
  type PaintOptions,
  type ViewRect,
} from './paint.ts';
import { groupPaintsEffects } from './paint-effects.ts';
import { startPaintBudgets, type BudgetAmounts } from './paint-budgets.ts';
import type { PaintContext } from './paint-fill.ts';
import { Defs, svg, SVG_NS } from './svg.ts';

/** Whether two views are the same rectangle (both absent counts). */
function sameView(a: ViewRect | undefined, b: ViewRect | undefined): boolean {
  if (!a || !b) return a === b;
  return a.x === b.x && a.y === b.y && a.width === b.width && a.height === b.height;
}

/** The painted node of one item and the definitions it added. */
interface PaintedItem {
  readonly node: SVGElement | null;
  readonly defs: SVGElement[];
  /**
   * What it spent of the paint budgets (`PaintBudgets`: tile images and shape gradient bands): a
   * later paint that keeps it spends them.
   */
  spent: BudgetAmounts;
  /**
   * What was left of each budget when it began to paint, and which budgets a fill in it read. Its
   * paint depends on nothing else of them, so a later paint keeps it only with as much left of
   * each budget it read.
   */
  readonly startLeft: BudgetAmounts;
  reads: { tiles: boolean; bands: boolean };
  /** A group painted in part (`paint`'s budget): the index of its first child left out. */
  rest?: number;
}

const dropPainted = (painted: PaintedItem): void => {
  painted.node?.remove();
  for (const d of painted.defs) d.remove();
};

/**
 * Paints the slides of one view again and again, and reuses the nodes of the items it painted
 * before. A layout item is an immutable value, and the layout keeps an item's object while its
 * shape is unchanged (`layoutSlide`), so the same item paints the same nodes: an edit paints the
 * items it changed, not the whole slide. `paintSlide` paints from nothing, for exports,
 * thumbnails, and the fidelity check; both paint each item with `paintItem`.
 *
 * The state, and what holds after each `paint`:
 * - `#items` holds the node on screen of each item, whole or painted in part (`rest`). An item
 *   is current when it is not in `#stale` and each budget it read has as much left now as when
 *   it began to paint (`PaintBudgets`); only a current item is kept. A paint decides this again
 *   for each item, from these two facts alone.
 * - `#stale` holds the items that `invalidate` or `view` made old. An item leaves it when it
 *   paints again whole, or leaves the layout.
 * - `#building` holds the new paint in part of an item that is not current, while its earlier
 *   node stays on screen in `#items`. It goes on only while it is current itself; it goes when
 *   its item is current again, paints whole, or leaves the layout.
 * - `#twice` holds the second paints of items that a layout holds twice: never reused.
 * - When `pending` is false, every item of the layout is whole and current, so the slide equals
 *   a fresh `paintSlide` of the layout.
 *
 * Nothing else may change the nodes it returns, except for a while: a drag preview hides nodes
 * and must put them back (`PaintedPreview.clear`) before the next `paint`. @internal
 */
export class SlidePainter {
  readonly #doc: Document;
  #opts: PaintOptions;
  readonly #defs: Defs;
  #root: SVGSVGElement | null = null;
  #content: SVGElement | null = null;
  #background: PaintedItem | null = null;
  #items = new Map<LayoutItem, PaintedItem>();
  #twice: PaintedItem[] = [];
  #building = new Map<LayoutItem, PaintedItem>();
  #stale = new Set<LayoutItem>();
  #slide = '';

  constructor(doc: Document, opts: PaintOptions) {
    this.#doc = doc;
    this.#opts = opts;
    this.#defs = new Defs(doc, opts.idPrefix);
  }

  /** The prefix of the ids this painter writes. */
  get idPrefix(): string {
    return this.#opts.idPrefix;
  }

  #key: unknown = undefined;

  /**
   * Note `key`: what the picture depends on besides the layout, such as converted pictures or
   * loaded faces. When it changed, every item paints again, and until it does, its earlier node
   * stays on screen in its place. Answers whether the key changed.
   */
  invalidate(key: unknown): boolean {
    if (key === this.#key) return false;
    this.#key = key;
    for (const item of this.#items.keys()) this.#stale.add(item);
    this.#dropBuilding();
    return true;
  }

  /** Drop the paints in part made before a change of key or view. */
  #dropBuilding(): void {
    for (const painted of this.#building.values()) dropPainted(painted);
    this.#building.clear();
  }

  /**
   * Paint for `view` (`PaintOptions.view`) from now on. When it changed, every item paints
   * again, and until it does, its earlier node stays on screen in its place. Answers whether
   * the view changed.
   */
  view(view: ViewRect | undefined): boolean {
    if (sameView(view, this.#opts.view)) return false;
    const { view: _, ...rest } = this.#opts;
    this.#opts = view ? { ...rest, view } : rest;
    for (const item of this.#items.keys()) if (readsView(item)) this.#stale.add(item);
    this.#dropBuilding();
    return true;
  }

  /** Whether this painter paints with `opts`: the same ports and id prefix (any view). */
  paints(opts: PaintOptions): boolean {
    const a = this.#opts;
    return (
      a.images === opts.images &&
      a.imageSize === opts.imageSize &&
      a.idPrefix === opts.idPrefix &&
      a.frameLabel === opts.frameLabel &&
      a.hitTargets === opts.hitTargets &&
      a.altText === opts.altText
    );
  }

  /**
   * Whether the last `paint` left items out because they passed its budget: a later `paint` of
   * the same layout adds them.
   */
  get pending(): boolean {
    return this.#pending;
  }

  #pending = false;

  /** Start over on a new `<svg>` for `layout`'s slide. */
  #reset(layout: SlideLayout): void {
    const doc = this.#doc;
    const { width, height } = layout.slide;
    this.#root = doc.createElementNS(SVG_NS, 'svg') as SVGSVGElement;
    this.#root.setAttribute('viewBox', `0 0 ${width} ${height}`);
    this.#root.setAttribute('preserveAspectRatio', 'none');
    this.#root.setAttribute('data-slide-part', layout.slide.partName);
    this.#defs.element.replaceChildren();
    this.#root.appendChild(this.#defs.element);
    this.#content = svg(doc, 'g', { 'data-kind': 'content' });
    this.#root.appendChild(this.#content);
    this.#background = null;
    this.#items.clear();
    this.#building.clear();
    this.#stale.clear();
    this.#twice = [];
  }

  /**
   * Paint `layout`, reusing the nodes of items painted before. Returns the same `<svg>`. With a
   * `budget`, it paints at most that many new items, in paint order, and keeps every item it
   * painted before (`pending` tells when some are left): a slide of 24,000 shapes then goes on
   * screen over several frames, and no frame blocks for seconds. The items in a group count
   * (`paintWeight`): a group of 20,000 shapes paints its children over several frames too.
   */
  paint(layout: SlideLayout, budget = Infinity): SVGSVGElement {
    const doc = this.#doc;
    const { width, height } = layout.slide;
    const slide = `${layout.slide.partName} ${width}x${height}`;
    if (!this.#root || !this.#content || this.#slide !== slide) {
      this.#reset(layout);
      this.#slide = slide;
    }
    const root = this.#root!;
    const content = this.#content!;
    const opts = this.#opts;
    const ctx: PaintContext = {
      defs: this.#defs,
      images: opts.images,
      ...(opts.imageSize ? { imageSize: opts.imageSize } : {}),
      sourcePart: layout.slide.backgroundPart,
    };
    // Each paint has whole budgets (tile images and shape gradient bands), spent in paint order:
    // the background, then each item on screen, kept or new. So the edits before a paint do not
    // use them up, and what is on screen stays within one budget.
    const budgets = startPaintBudgets(this.#defs);
    /** Run `paint` and count what it spent of the budgets, and which it read. */
    const counted = <T>(paint: () => T) => {
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
    };
    /** The painted item of a counted paint. */
    const paintedOf = (fresh: ReturnType<typeof counted<SVGElement | null>>): PaintedItem => ({
      node: fresh.value,
      defs: fresh.defs,
      spent: fresh.spent,
      startLeft: fresh.startLeft,
      reads: fresh.reads,
    });
    /**
     * Whether a paint, whole or in part, still draws what a fresh paint draws here: each budget
     * it read has as much left as when it began.
     */
    const fits = ({ reads, startLeft }: PaintedItem): boolean =>
      (!reads.tiles || startLeft.tiles === budgets.tiles.left) &&
      (!reads.bands || startLeft.bands === budgets.bands.left);
    /** Whether the node of `item` on screen is kept. */
    const current = (item: LayoutItem, painted: PaintedItem): boolean =>
      !this.#stale.has(item) && fits(painted);
    // The background is one rectangle: it paints each time.
    if (this.#background) dropPainted(this.#background);
    for (const painted of this.#twice.splice(0)) dropPainted(painted);
    const background = counted(() =>
      paintBackground(doc, layout.slide.background, width, height, ctx)
    );
    this.#background = paintedOf(background);
    root.insertBefore(background.value, content);
    /** Paint `item` from nothing, a group in part when it weighs more than `room`. */
    const paintFresh = (item: LayoutItem, room: number): [PaintedItem, number] => {
      // A group with effects paints whole: its shadow and filters cover all of its children.
      // The cascade bounds its weight (`MAX_GROUP_EFFECT_ELEMENTS`) and nesting, so it fits a frame.
      if (item.kind === 'group' && paintWeight(item) > room && !groupPaintsEffects(item)) {
        const g = groupNode(doc, item, opts);
        const fresh = counted(() => paintChildren(doc, item, g, ctx, opts, 0, room));
        const { next: rest, spent } = fresh.value;
        const painted = paintedOf({ ...fresh, value: g });
        if (rest < item.children.length) painted.rest = rest;
        return [painted, spent + 1];
      }
      return [paintedOf(counted(() => paintItem(doc, item, ctx, opts, 1))), paintWeight(item)];
    };
    /**
     * Go on with a group painted in part, in `room`, after its earlier children spend their
     * budgets. Answers the weight painted.
     */
    const paintMore = (item: LayoutItem, painted: PaintedItem, room: number): number => {
      budgets.spend(painted.spent);
      if (item.kind !== 'group' || painted.rest === undefined || !painted.node) return 0;
      const node = painted.node;
      const more = counted(() => paintChildren(doc, item, node, ctx, opts, painted.rest!, room));
      painted.defs.push(...more.defs);
      painted.spent = {
        tiles: painted.spent.tiles + more.spent.tiles,
        bands: painted.spent.bands + more.spent.bands,
      };
      painted.reads = {
        tiles: painted.reads.tiles || more.reads.tiles,
        bands: painted.reads.bands || more.reads.bands,
      };
      if (more.value.next < item.children.length) painted.rest = more.value.next;
      else delete painted.rest;
      return more.value.spent;
    };
    /** Drop the paint in part of `item`, if any. */
    const dropBuild = (item: LayoutItem): void => {
      const building = this.#building.get(item);
      if (building) dropPainted(building);
      this.#building.delete(item);
    };
    const next = new Map<LayoutItem, PaintedItem>();
    const nodes: SVGElement[] = [];
    /** Show `painted` as the node of `item`. */
    const show = (item: LayoutItem, painted: PaintedItem): void => {
      next.set(item, painted);
      if (painted.node) nodes.push(painted.node);
    };
    let left = budget;
    this.#pending = false;
    // The earlier nodes of the shapes that an edit replaced (`itemKey`): while the budget leaves
    // a new shape out, its earlier node stays on screen in its place.
    let earlier: Map<string, [LayoutItem, PaintedItem]> | null = null;
    let inLayout: Set<LayoutItem> | null = null;
    for (const item of layout.items) {
      const old = next.has(item) ? undefined : this.#items.get(item);
      if (old && current(item, old)) {
        dropBuild(item);
        // A group painted in part goes on with its children.
        if (old.rest !== undefined && left > 0) left -= paintMore(item, old, left);
        else budgets.spend(old.spent);
        if (old.rest !== undefined) this.#pending = true;
        show(item, old);
        continue;
      }
      if (left <= 0) {
        this.#pending = true;
        earlier ??= new Map([...this.#items].map(([i, p]) => [itemKey(i), [i, p]] as const));
        inLayout ??= new Set(layout.items);
        const replaced = earlier.get(itemKey(item));
        const stand = old ? ([item, old] as const) : replaced;
        // An earlier item stands in for itself, or for the item that replaced it in the layout.
        if (stand && !next.has(stand[0]) && (stand[0] === item || !inLayout.has(stand[0]))) {
          budgets.spend(stand[1].spent);
          show(stand[0], stand[1]);
        }
        continue;
      }
      if (next.has(item)) {
        // An item twice in one layout paints twice: one node cannot stand in two places.
        const [painted] = paintFresh(item, Infinity);
        this.#twice.push(painted);
        left -= paintWeight(item);
        if (painted.node) nodes.push(painted.node);
        continue;
      }
      let painted = this.#building.get(item);
      if (painted && fits(painted)) left -= paintMore(item, painted, left);
      else {
        dropBuild(item);
        let spent: number;
        [painted, spent] = paintFresh(item, left);
        left -= spent;
      }
      if (painted.rest !== undefined) {
        this.#pending = true;
        if (old) {
          // A group that paints again in part: its earlier node stays until the new one is whole.
          this.#building.set(item, painted);
          show(item, old);
          continue;
        }
      }
      this.#building.delete(item);
      this.#stale.delete(item);
      if (old) dropPainted(old);
      show(item, painted);
    }
    for (const [item, painted] of this.#items) if (!next.has(item)) dropPainted(painted);
    for (const item of [...this.#building.keys()]) if (!next.has(item)) dropBuild(item);
    for (const item of this.#stale) if (!next.has(item)) this.#stale.delete(item);
    this.#items = next;
    // Put the nodes in paint order, moving only the ones out of place.
    nodes.forEach((node, i) => {
      const at = content.childNodes[i] ?? null;
      if (at !== node) content.insertBefore(node, at);
    });
    while (content.childNodes.length > nodes.length) content.lastChild?.remove();
    return root;
  }
}

/** Whether an item's paint depends on `PaintOptions.view`: it holds text or a table. */
function readsView(item: LayoutItem): boolean {
  switch (item.kind) {
    case 'shape':
      return !!item.text;
    case 'picture':
      return false;
    case 'group':
      return item.children.some(readsView);
    case 'frame':
      return !!item.table || !!item.drawing?.some(readsView) || !!item.chart?.some(readsView);
  }
}

/**
 * An item's place in the document: its part, its layer, and its node id. Node ids are numbered
 * per part, so a master shape and a slide shape can share one.
 */
const itemKey = (item: LayoutItem): string =>
  `${item.element.partName}\u0000${item.element.layer}\u0000${item.element.nodeId}`;
