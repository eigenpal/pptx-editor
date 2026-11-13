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
import type { PaintContext } from './paint-fill.ts';
import { dropPainted, PaintMeter, type MeteredPaint, type PaintedItem } from './painted-items.ts';
import { Defs, svg, SVG_NS } from './svg.ts';

/** Whether two views are the same rectangle (both absent counts). */
function sameView(a: ViewRect | undefined, b: ViewRect | undefined): boolean {
  if (!a || !b) return a === b;
  return a.x === b.x && a.y === b.y && a.width === b.width && a.height === b.height;
}

/** The state of one `paint` call. It lives only while that call runs. */
class Frame {
  /** The paint budgets of this paint (tile images and shape gradient bands). */
  readonly meter: PaintMeter;
  /** The weight (`paintWeight`) that this frame may still paint. */
  left: number;
  /** Whether some item is left out or partial at the end of this frame. */
  pending = false;
  /** The record on screen of each item after this frame. */
  readonly shown = new Map<LayoutItem, PaintedItem>();
  /** The nodes of the content group after this frame, in paint order. */
  readonly nodes: SVGElement[] = [];
  /** The records before this frame by `itemKey`, made when an item is first left out. */
  earlier: Map<string, readonly [LayoutItem, PaintedItem]> | null = null;
  /** The items of the layout, made with `earlier`. */
  inLayout: Set<LayoutItem> | null = null;

  constructor(
    readonly layout: SlideLayout,
    readonly ctx: PaintContext,
    defs: Defs,
    budget: number
  ) {
    this.meter = new PaintMeter(defs);
    this.left = budget;
  }

  /** Show `painted` as the node of `item`. */
  show(item: LayoutItem, painted: PaintedItem): void {
    this.shown.set(item, painted);
    if (painted.node) this.nodes.push(painted.node);
  }
}

/**
 * Paints the slides of one view again and again, and reuses the nodes of the items it painted
 * before. A layout item is an immutable value, and the layout keeps an item's object while its
 * shape is unchanged (`layoutSlide`), so the same item paints the same nodes: an edit paints only
 * the items that it changed. `paintSlide` paints from nothing, for exports, thumbnails, and the
 * fidelity check. Both paint each item with `paintItem`.
 *
 * Terms:
 * - A record (`PaintedItem`) holds the nodes of one item and what they spent and read of the
 *   paint budgets. A record is partial when it holds only the first children of a group.
 * - A budget read: a fill read a budget (tile images, shape gradient bands) when what it drew
 *   depended on what was left of it.
 * - A record is stale when `invalidate` or `view` marked it. Its nodes may be wrong.
 * - A record is current when it is not stale and it fits the budgets (`PaintMeter.fits`): each
 *   budget that it read has as much left at its place in paint order as when it began to read it.
 *   `#isCurrent` is the one test for this, for whole, partial, and building records.
 * - An item is kept when its record is current: the paint spends the record's share of the
 *   budgets and does not paint the item again. A kept partial record continues within the frame.
 *
 * Invariants after each `paint`:
 * - `#items` holds the record on screen of each item of the layout that has one.
 * - `#building` holds a new partial paint of an item whose record in `#items` is not current.
 *   The earlier record stays on screen until the new one is whole. A building record is never
 *   stale, because `invalidate` and `view` drop all of them.
 * - `#twice` holds the extra paints of an item that the layout holds more than once. They paint
 *   again on each paint and take nothing of the frame budget.
 * - `pending` is false only when every item of the layout has a current, whole record. Then the
 *   slide equals a fresh `paintSlide` of the layout.
 *
 * One `paint` call:
 * 1. Starts a new `<svg>` when the slide or its size changed (`#begin`).
 * 2. Starts whole budgets and paints the background again (`#paintBackground`).
 * 3. Walks the layout items in paint order (`#paintItem`): it keeps current records, paints new
 *    items while the frame budget lasts, and shows an earlier record in place of an item that it
 *    leaves out.
 * 4. Drops the records that left the layout, and puts the nodes in paint order (`#commit`).
 *
 * While `pending` is true, the caller paints the same layout again in a later frame. Nothing else
 * may change the nodes it returns, except for a while: a drag preview hides nodes and must put
 * them back (`PaintedPreview.clear`) before the next `paint`. @internal
 */
export class SlidePainter {
  readonly #doc: Document;
  #opts: PaintOptions;
  readonly #defs: Defs;
  #root: SVGSVGElement | null = null;
  #content: SVGElement | null = null;
  #slide = '';
  #key: unknown = undefined;
  #pending = false;
  #background: PaintedItem | null = null;
  #items = new Map<LayoutItem, PaintedItem>();
  #building = new Map<LayoutItem, PaintedItem>();
  #twice: PaintedItem[] = [];
  /** The stale records. A record that leaves the slide is never read again. */
  #stale = new WeakSet<PaintedItem>();

  constructor(doc: Document, opts: PaintOptions) {
    this.#doc = doc;
    this.#opts = opts;
    this.#defs = new Defs(doc, opts.idPrefix);
  }

  /** The prefix of the ids this painter writes. */
  get idPrefix(): string {
    return this.#opts.idPrefix;
  }

  /**
   * Note `key`: what the picture depends on besides the layout, such as converted pictures or
   * loaded faces. When it changed, every item paints again, and until it does, its earlier node
   * stays on screen in its place. Answers whether the key changed.
   */
  invalidate(key: unknown): boolean {
    if (key === this.#key) return false;
    this.#key = key;
    for (const painted of this.#items.values()) this.#stale.add(painted);
    this.#dropAllBuilding();
    return true;
  }

  /**
   * Paint for `view` (`PaintOptions.view`) from now on. When it changed, every item that holds
   * text or a table paints again, and until it does, its earlier node stays on screen in its
   * place. Answers whether the view changed.
   */
  view(view: ViewRect | undefined): boolean {
    if (sameView(view, this.#opts.view)) return false;
    const { view: _, ...rest } = this.#opts;
    this.#opts = view ? { ...rest, view } : rest;
    for (const [item, painted] of this.#items) if (readsView(item)) this.#stale.add(painted);
    this.#dropAllBuilding();
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
   * Whether the last `paint` left items out or partial because they passed its budget: a later
   * `paint` of the same layout adds them.
   */
  get pending(): boolean {
    return this.#pending;
  }

  /**
   * Paint `layout`, reusing the nodes of items painted before. Returns the same `<svg>`. With a
   * `budget`, it paints at most that much new weight, in paint order, and keeps every item it
   * painted before (`pending` tells when some are left): a slide of 24,000 shapes then goes on
   * screen over several frames, and no frame blocks for seconds. The items in a group count
   * (`paintWeight`): a group of 20,000 shapes paints its children over several frames too.
   */
  paint(layout: SlideLayout, budget = Infinity): SVGSVGElement {
    const { root, content } = this.#begin(layout);
    const opts = this.#opts;
    const ctx: PaintContext = {
      defs: this.#defs,
      images: opts.images,
      ...(opts.imageSize ? { imageSize: opts.imageSize } : {}),
      sourcePart: layout.slide.backgroundPart,
    };
    const frame = new Frame(layout, ctx, this.#defs, budget);
    this.#paintBackground(frame, root, content);
    for (const item of layout.items) this.#paintItem(frame, item);
    this.#commit(frame, content);
    return root;
  }

  /** Step 1: use the current `<svg>`, or start a new one when the slide or its size changed. */
  #begin(layout: SlideLayout): { root: SVGSVGElement; content: SVGElement } {
    const { width, height, partName } = layout.slide;
    const slide = `${partName} ${width}x${height}`;
    if (this.#root && this.#content && this.#slide === slide) {
      return { root: this.#root, content: this.#content };
    }
    const doc = this.#doc;
    const root = doc.createElementNS(SVG_NS, 'svg') as SVGSVGElement;
    root.setAttribute('viewBox', `0 0 ${width} ${height}`);
    root.setAttribute('preserveAspectRatio', 'none');
    root.setAttribute('data-slide-part', partName);
    this.#defs.element.replaceChildren();
    root.appendChild(this.#defs.element);
    const content = svg(doc, 'g', { 'data-kind': 'content' });
    root.appendChild(content);
    this.#root = root;
    this.#content = content;
    this.#slide = slide;
    this.#background = null;
    this.#items.clear();
    this.#building.clear();
    this.#stale = new WeakSet();
    this.#twice = [];
    return { root, content };
  }

  /**
   * Step 2: paint the background again (one rectangle), first in the budgets, and drop the extra
   * paints of the last frame.
   */
  #paintBackground(frame: Frame, root: SVGSVGElement, content: SVGElement): void {
    if (this.#background) dropPainted(this.#background);
    for (const painted of this.#twice.splice(0)) dropPainted(painted);
    const { width, height, background } = frame.layout.slide;
    const painted = frame.meter.whole(() =>
      paintBackground(this.#doc, background, width, height, frame.ctx)
    );
    this.#background = painted;
    root.insertBefore(painted.node!, content);
  }

  /** Step 3: paint one layout item, or keep it, or leave it out of this frame. */
  #paintItem(frame: Frame, item: LayoutItem): void {
    // A second copy of an item has no record of its own: `#items` holds the first copy's.
    const twice = frame.shown.has(item);
    const old = twice ? undefined : this.#items.get(item);
    if (old && this.#isCurrent(frame, old)) this.#keep(frame, item, old);
    else if (frame.left <= 0) this.#leaveOut(frame, item, old);
    else if (twice) this.#paintTwice(frame, item);
    else this.#paintAgain(frame, item, old);
  }

  /**
   * Whether the nodes of `painted` are what a new paint draws here: it is not stale, and each
   * budget that it read has as much left as when it began to read it.
   */
  #isCurrent(frame: Frame, painted: PaintedItem): boolean {
    return !this.#stale.has(painted) && frame.meter.fits(painted);
  }

  /** Keep the current record of `item`. A partial group continues while the frame has room. */
  #keep(frame: Frame, item: LayoutItem, old: PaintedItem): void {
    this.#dropBuilding(item);
    let painted = old;
    if (old.rest !== undefined && frame.left > 0) painted = this.#continue(frame, item, old);
    else frame.meter.keep(old);
    if (painted.rest !== undefined) frame.pending = true;
    frame.show(item, painted);
  }

  /**
   * Leave `item` out of this frame, which has no room. Its earlier record stays on screen, or the
   * earlier record of the item that it replaced (`#standIn`).
   */
  #leaveOut(frame: Frame, item: LayoutItem, old: PaintedItem | undefined): void {
    frame.pending = true;
    const stand = this.#standIn(frame, item, old);
    if (!stand) return;
    frame.meter.keep(stand[1]);
    frame.show(stand[0], stand[1]);
  }

  /**
   * The earlier item and record that stand in for `item` while it is left out: its own record,
   * or else the record of an item at the same place (`itemKey`) that is no longer in the layout,
   * such as the item that an edit replaced.
   */
  #standIn(
    frame: Frame,
    item: LayoutItem,
    old: PaintedItem | undefined
  ): readonly [LayoutItem, PaintedItem] | undefined {
    frame.earlier ??= new Map([...this.#items].map(([i, p]) => [itemKey(i), [i, p]] as const));
    frame.inLayout ??= new Set(frame.layout.items);
    if (old) return [item, old];
    const replaced = frame.earlier.get(itemKey(item));
    if (!replaced) return undefined;
    const [earlierItem] = replaced;
    if (frame.shown.has(earlierItem)) return undefined;
    return earlierItem === item || !frame.inLayout.has(earlierItem) ? replaced : undefined;
  }

  /**
   * Paint a second copy of an item that the layout holds more than once: one node cannot stand
   * in two places. The copy paints whole on each paint and takes nothing of the frame budget, so
   * the items after it still paint and the slide settles.
   */
  #paintTwice(frame: Frame, item: LayoutItem): void {
    const { painted } = this.#paintNew(frame, item, Infinity);
    this.#twice.push(painted);
    if (painted.node) frame.nodes.push(painted.node);
  }

  /**
   * Paint `item`, which has no current record. Continue its building record when that is
   * current, or else paint it new. A partial result of an item with an earlier record goes to
   * `#building`, and the earlier record stays on screen until the new one is whole.
   */
  #paintAgain(frame: Frame, item: LayoutItem, old: PaintedItem | undefined): void {
    const building = this.#building.get(item);
    let painted: PaintedItem;
    if (building && this.#isCurrent(frame, building)) {
      painted = this.#continue(frame, item, building);
    } else {
      this.#dropBuilding(item);
      const fresh = this.#paintNew(frame, item, frame.left);
      frame.left -= fresh.weight;
      painted = fresh.painted;
    }
    if (painted.rest !== undefined) {
      frame.pending = true;
      if (old) {
        this.#building.set(item, painted);
        frame.show(item, old);
        return;
      }
    }
    this.#building.delete(item);
    if (old) dropPainted(old);
    frame.show(item, painted);
  }

  /**
   * Paint `item` from nothing. A group that weighs more than `room` paints only its first
   * children, unless it has effects: its shadow and filters cover all of its children, and the
   * cascade bounds its weight (`MAX_GROUP_EFFECT_ELEMENTS`) and nesting, so it fits a frame.
   */
  #paintNew(frame: Frame, item: LayoutItem, room: number): MeteredPaint {
    const doc = this.#doc;
    const opts = this.#opts;
    const { meter, ctx } = frame;
    if (item.kind === 'group' && paintWeight(item) > room && !groupPaintsEffects(item)) {
      const g = groupNode(doc, item, opts);
      return meter.startGroup(g, item.children.length, () =>
        paintChildren(doc, item, g, ctx, opts, 0, room)
      );
    }
    const painted = meter.whole(() => paintItem(doc, item, ctx, opts, 1));
    return { painted, weight: paintWeight(item) };
  }

  /** Keep `painted` and paint more of it within the frame's room. Answers the new record. */
  #continue(frame: Frame, item: LayoutItem, painted: PaintedItem): PaintedItem {
    // Only a group has a partial record (`rest`).
    if (item.kind !== 'group') {
      frame.meter.keep(painted);
      return painted;
    }
    const more = frame.meter.extend(painted, item.children.length, (node, from) =>
      paintChildren(this.#doc, item, node, frame.ctx, this.#opts, from, frame.left)
    );
    frame.left -= more.weight;
    return more.painted;
  }

  /**
   * Step 4: make the shown records the new `#items`, drop the records and building records that
   * left the slide, and put the nodes in paint order, moving only the ones out of place.
   */
  #commit(frame: Frame, content: SVGElement): void {
    const { shown, nodes } = frame;
    for (const [item, painted] of this.#items) if (!shown.has(item)) dropPainted(painted);
    for (const item of [...this.#building.keys()]) if (!shown.has(item)) this.#dropBuilding(item);
    this.#items = shown;
    this.#pending = frame.pending;
    nodes.forEach((node, i) => {
      const at = content.childNodes[i] ?? null;
      if (at !== node) content.insertBefore(node, at);
    });
    while (content.childNodes.length > nodes.length) content.lastChild?.remove();
  }

  /** Drop the building record of `item`, if any. */
  #dropBuilding(item: LayoutItem): void {
    const building = this.#building.get(item);
    if (building) dropPainted(building);
    this.#building.delete(item);
  }

  /** Drop every building record: a change of key or view makes them old. */
  #dropAllBuilding(): void {
    for (const painted of this.#building.values()) dropPainted(painted);
    this.#building.clear();
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
