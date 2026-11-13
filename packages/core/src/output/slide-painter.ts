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
import { startTileBudget, type PaintContext } from './paint-fill.ts';
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
  /** The tile images it draws (`MAX_SLIDE_TILE_IMAGES`): a later paint that keeps it spends them. */
  tiles: number;
  /**
   * The tile images left when it began to paint, and whether a tiled fill in it read them. Its
   * paint depends on nothing else of the budget, so a later paint keeps it only with as many left.
   */
  tileLeft: number;
  tiled: boolean;
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
  /** Items that a layout held twice: their second paint is never reused. */
  #twice: PaintedItem[] = [];
  /**
   * The new paints, in part, of groups that paint again: each group's earlier node stays on
   * screen until its new one is whole.
   */
  #building = new Map<LayoutItem, PaintedItem>();
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
  /** Items painted before the last change of `invalidate`'s key: each paints again. */
  #stale: Set<LayoutItem> | null = null;

  /**
   * Note `key`: what the picture depends on besides the layout, such as converted pictures or
   * loaded faces. When it changed, every item paints again, and until it does, its earlier node
   * stays on screen in its place. Answers whether the key changed.
   */
  invalidate(key: unknown): boolean {
    if (key === this.#key) return false;
    this.#key = key;
    this.#stale = new Set(this.#items.keys());
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
    const stale = this.#stale ?? new Set<LayoutItem>();
    for (const item of this.#items.keys()) if (readsView(item)) stale.add(item);
    this.#stale = stale;
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
      this.#stale = null;
      this.#twice = [];
      this.#slide = slide;
    }
    const root = this.#root;
    const content = this.#content;
    const ctx: PaintContext = {
      defs: this.#defs,
      images: this.#opts.images,
      ...(this.#opts.imageSize ? { imageSize: this.#opts.imageSize } : {}),
      sourcePart: layout.slide.backgroundPart,
    };
    // Each paint has the whole tile budget, spent in paint order: the background, then each item
    // on screen, kept or new. So the edits before a paint do not use up its budget, the tiles on
    // screen stay within one budget, and a kept item with tiles paints again when the images left
    // for it changed (`retiles`): the slide draws the tiles of a fresh paint.
    const tileBudget = startTileBudget(this.#defs);
    /** Run `paint` and count the tile images it spent, and whether it read the budget. */
    const counted = <T>(paint: () => T) => {
      const [from, asks, tileLeft] = [tileBudget.spent, tileBudget.asks, tileBudget.left];
      const fresh = this.#defs.capture(paint);
      const tiles = tileBudget.spent - from;
      return {
        value: fresh.value,
        defs: [...fresh.defs],
        tiles,
        tileLeft,
        tiled: tileBudget.asks > asks,
      };
    };
    /** The painted item of a counted paint. */
    const paintedOf = (fresh: ReturnType<typeof counted<SVGElement | null>>): PaintedItem => ({
      node: fresh.value,
      defs: fresh.defs,
      tiles: fresh.tiles,
      tileLeft: fresh.tileLeft,
      tiled: fresh.tiled,
    });
    /**
     * Whether a kept item paints otherwise now: a tiled fill in it read a budget with other
     * images left. An item before it in paint order changed its tiles, so it paints again, as a
     * fresh paint draws it. A group painted in part goes on with its children instead.
     */
    const retiles = (painted: PaintedItem | undefined): boolean =>
      !!painted?.tiled && painted.rest === undefined && painted.tileLeft !== tileBudget.left;
    // The background is one rectangle: it paints each time.
    const drop = dropPainted;
    if (this.#background) drop(this.#background);
    for (const painted of this.#twice.splice(0)) drop(painted);
    const background = counted(() =>
      paintBackground(doc, layout.slide.background, width, height, ctx)
    );
    this.#background = paintedOf(background);
    root.insertBefore(background.value, content);
    const next = new Map<LayoutItem, PaintedItem>();
    const nodes: SVGElement[] = [];
    const opts = this.#opts;
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
     * tiles. Answers the weight painted.
     */
    const paintMore = (item: LayoutItem, painted: PaintedItem, room: number): number => {
      tileBudget.spend(painted.tiles);
      if (item.kind !== 'group' || painted.rest === undefined || !painted.node) return 0;
      const node = painted.node;
      const from = painted.rest;
      const more = counted(() => paintChildren(doc, item, node, ctx, opts, from, room));
      painted.defs.push(...more.defs);
      painted.tiles += more.tiles;
      painted.tiled ||= more.tiled;
      if (more.value.next < item.children.length) painted.rest = more.value.next;
      else delete painted.rest;
      return more.value.spent;
    };
    let left = budget;
    this.#pending = false;
    // The earlier nodes of the shapes that a budget leaves out (`itemKey`): each stays on
    // screen in its place until its shape paints again (a font load or a theme edit makes
    // every item new).
    let earlier: Map<string, [LayoutItem, PaintedItem]> | null = null;
    let current: Set<LayoutItem> | null = null;
    for (const item of layout.items) {
      let painted: PaintedItem;
      const stale =
        this.#stale?.has(item) === true || (!next.has(item) && retiles(this.#items.get(item)));
      const kept = next.has(item) || stale ? undefined : this.#items.get(item);
      if (kept) {
        // A group painted in part goes on with its children.
        if (kept.rest !== undefined && left > 0) left -= paintMore(item, kept, left);
        else tileBudget.spend(kept.tiles);
        if (kept.rest !== undefined) this.#pending = true;
        next.set(item, kept);
        if (kept.node) nodes.push(kept.node);
        continue;
      }
      if (left <= 0) {
        this.#pending = true;
        earlier ??= new Map([...this.#items].map(([i, p]) => [itemKey(i), [i, p]] as const));
        current ??= new Set(layout.items);
        const stand = earlier.get(itemKey(item));
        // An item of the layout stands in only for itself, when it waits to paint again.
        const own = stand?.[0] === item && stale;
        if (stand && !next.has(stand[0]) && (own || !current.has(stand[0]))) {
          tileBudget.spend(stand[1].tiles);
          next.set(stand[0], stand[1]);
          if (stand[1].node) nodes.push(stand[1].node);
        }
        continue;
      }
      if (next.has(item)) {
        // An item twice in one layout paints twice: one node cannot stand in two places.
        [painted] = paintFresh(item, Infinity);
        this.#twice.push(painted);
        left -= paintWeight(item);
      } else {
        const building = this.#building.get(item);
        if (building) {
          painted = building;
          left -= paintMore(item, painted, left);
        } else {
          let spent: number;
          [painted, spent] = paintFresh(item, left);
          left -= spent;
        }
        const old = stale ? this.#items.get(item) : undefined;
        if (painted.rest !== undefined) {
          this.#pending = true;
          if (old) {
            // A group that paints again in part: its earlier node stays until the new one is whole.
            this.#building.set(item, painted);
            next.set(item, old);
            if (old.node) nodes.push(old.node);
            continue;
          }
        }
        this.#building.delete(item);
        next.set(item, painted);
        if (stale) {
          // Painted again: its earlier node and definitions go.
          this.#stale?.delete(item);
          if (old) drop(old);
        }
      }
      if (painted.node) nodes.push(painted.node);
    }
    for (const [item, painted] of this.#items) if (!next.has(item)) drop(painted);
    for (const [item, painted] of this.#building)
      if (!next.has(item) || !this.#stale?.has(item)) {
        drop(painted);
        this.#building.delete(item);
      }
    this.#items = next;
    // An item an edit replaced before it painted again waits no more.
    if (this.#stale) for (const item of this.#stale) if (!next.has(item)) this.#stale.delete(item);
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
