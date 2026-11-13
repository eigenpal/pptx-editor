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
  type ChildPaint,
  type PaintOptions,
  type ViewRect,
} from './paint.ts';
import { groupPaintsEffects } from './paint-effects.ts';
import type { PaintContext } from './paint-fill.ts';
import { dropPainted, PaintMeter, type PaintedItem } from './painted-items.ts';
import { Defs, svg, SVG_NS } from './svg.ts';

/** Whether two views are the same rectangle (both absent counts). */
function sameView(a: ViewRect | undefined, b: ViewRect | undefined): boolean {
  if (!a || !b) return a === b;
  return a.x === b.x && a.y === b.y && a.width === b.width && a.height === b.height;
}

declare const checked: unique symbol;
/**
 * A record that `SlidePainter.#current` found current in this frame. Only `#current` makes one,
 * so a step that takes it cannot skip the check.
 */
type CurrentRecord = PaintedItem & { readonly [checked]: true };

/** An item and its record. */
type Entry = readonly [LayoutItem, PaintedItem];

/**
 * The state of one `paint` call. It lives only while that call runs. It owns the frame's room and
 * the nodes on screen: a paint step adds a node only through `show`, and counts its weight only
 * through the `record` methods.
 */
class Frame {
  /** The paint budgets of this paint (tile images and shape gradient bands). */
  readonly #meter: PaintMeter;
  /** Whether some item is left out or partial at the end of this frame. */
  pending = false;
  readonly #earlierItems: ReadonlyMap<LayoutItem, PaintedItem>;
  readonly #shown = new Map<LayoutItem, PaintedItem>();
  readonly #nodes: SVGElement[] = [];
  /** The records before this frame by `itemKey`, made when the first stand-in is looked up. */
  #byKey: Map<string, Entry> | null = null;
  /** The items of the layout, made with `#byKey`. */
  #inLayout: Set<LayoutItem> | null = null;
  #room: number;

  /**
   * A frame for `layout` within `budget` of paint weight. `earlierItems` are the records on screen
   * before it, by item; the frame reads them only, to find stand-ins.
   */
  constructor(
    readonly layout: SlideLayout,
    readonly ctx: PaintContext,
    defs: Defs,
    budget: number,
    earlierItems: ReadonlyMap<LayoutItem, PaintedItem>
  ) {
    this.#meter = PaintMeter.resetBudgets(defs);
    this.#room = budget;
    this.#earlierItems = earlierItems;
  }

  /**
   * The paint weight (`paintWeight`) that this frame may still paint. It is not a paint budget:
   * it bounds the work of one frame, and the slide's budgets bound what the slide draws.
   */
  get room(): number {
    return this.#room;
  }

  /** The record on screen of each item after this frame. */
  get shown(): ReadonlyMap<LayoutItem, PaintedItem> {
    return this.#shown;
  }

  /** The nodes of the content group after this frame, in paint order. */
  get nodes(): readonly SVGElement[] {
    return this.#nodes;
  }

  /**
   * Put `painted` on screen as the node of `item`, or as an extra copy when `item` is null. This
   * is the only way that a node joins the frame. It spends the budget share of `spentBy` (the
   * record shown, by default) at this place in paint order, unless this paint already spent it.
   */
  show(item: LayoutItem | null, painted: PaintedItem, spentBy: PaintedItem = painted): void {
    this.#meter.spendShare(spentBy);
    if (item) this.#shown.set(item, painted);
    if (painted.node) this.#nodes.push(painted.node);
  }

  /**
   * Whether the nodes of `painted` fit the budgets at this place in paint order
   * (`PaintMeter.fits`).
   */
  fits(painted: PaintedItem): boolean {
    return this.#meter.fits(painted);
  }

  /** Paint the background with `paint`, first in the budgets. It takes nothing of the room. */
  recordBackground(paint: () => SVGElement): PaintedItem {
    return this.#meter.recordWhole(paint);
  }

  /** Paint a whole item with `paint`, and count its `weight` against the room. */
  recordWhole(paint: () => SVGElement | null, weight: number): PaintedItem {
    const painted = this.#meter.recordWhole(paint);
    this.#room -= weight;
    return painted;
  }

  /** Paint an extra copy whole with `paint`. A copy takes nothing of the room. */
  recordCopy(paint: () => SVGElement | null): PaintedItem {
    return this.#meter.recordWhole(paint);
  }

  /**
   * Begin a group of `count` children in `node`: `paint` paints its first children within the
   * room that it gets. Counts their weight, and the group node's, against the room.
   */
  recordGroupStart(
    node: SVGElement,
    count: number,
    paint: (room: number) => ChildPaint
  ): PaintedItem {
    const made = this.#meter.recordGroupStart(node, count, () => paint(this.#room));
    this.#room -= made.weight;
    return made.painted;
  }

  /**
   * Spend the share of `painted`, and continue it when it is a partial group of `count`
   * children: `paint` paints from child `from` into `node` within the room that it gets. Counts
   * the new children's weight against the room. Answers the new record.
   */
  extend(
    painted: PaintedItem,
    count: number,
    paint: (node: SVGElement, from: number, room: number) => ChildPaint
  ): PaintedItem {
    const more = this.#meter.extend(painted, count, (node, from) => paint(node, from, this.#room));
    this.#room -= more.weight;
    return more.painted;
  }

  /**
   * The earlier item and record that stand in for `item`, which has no record of its own: the
   * record of an item at the same place (`itemKey`) that is no longer in the layout, such as the
   * item that an edit replaced. None when that earlier item is on screen already.
   */
  standIn(item: LayoutItem): Entry | undefined {
    // Made once for each frame, when the first item without a record is left out.
    this.#byKey ??= new Map([...this.#earlierItems].map(([i, p]) => [itemKey(i), [i, p]] as const));
    this.#inLayout ??= new Set(this.layout.items);
    const replaced = this.#byKey.get(itemKey(item));
    if (!replaced) return undefined;
    const [earlierItem] = replaced;
    if (this.#shown.has(earlierItem)) return undefined;
    return earlierItem === item || !this.#inLayout.has(earlierItem) ? replaced : undefined;
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
 *   `#current` is the one test for this, for whole, partial, and building records.
 * - An item is kept when its record is current: the paint spends the record's share of the
 *   budgets and does not paint the item again. A kept partial record continues within the frame.
 * - A stand-in is the record of an item that a frame leaves out for lack of room. It is the
 *   item's own earlier record, or the record of the earlier item at the same place (`itemKey`)
 *   that it replaced in the layout.
 *
 * Invariants after each `paint`:
 * - `#items` holds each record on screen, by its item. That item is in the layout, except for a
 *   stand-in of a replaced item: that record stays under the earlier item until the new item
 *   paints.
 * - For each item on screen, the record on screen, or the partial record that will replace it,
 *   spent its budget share once, at its place in paint order (`Frame.show`).
 * - `#building` holds a new partial paint of an item whose record in `#items` is not current.
 *   The earlier record stays on screen until the new one is whole. A building record is never
 *   stale, because `invalidate` and `view` drop all of them.
 * - `#twice` holds the extra copies of the items that the layout holds more than once. They have
 *   no records in `#items`.
 * - `pending` is false only when every item of the layout has a current, whole record in
 *   `#items`, and each extra copy painted whole in this frame. Then the slide equals a fresh
 *   `paintSlide` of the layout with the same key and view.
 *
 * One `paint` call:
 * 1. Starts a new `<svg>` when the slide or its size changed (`#begin`).
 * 2. Starts whole budgets and paints the background again (`#paintBackground`).
 * 3. Walks the layout items in paint order (`#paintItem`): it keeps current records, paints new
 *    items while the frame has room, and shows a stand-in for an item that it leaves out.
 * 4. Drops the records that left the slide, and puts the nodes in paint order (`#commit`).
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
  #items: ReadonlyMap<LayoutItem, PaintedItem> = new Map();
  #building = new Map<LayoutItem, PaintedItem>();
  /**
   * The extra copies of the last paint, one for each repeat of an item in the layout. One node
   * cannot stand in two places, so each repeat paints its own copy, and the next paint drops it.
   */
  #twice: PaintedItem[] = [];
  /**
   * The stale records. A `WeakSet` needs no cleanup: a record that leaves the slide is never read
   * again, and a new record of the same item is not stale.
   */
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
   * Paint `layout`, reusing the nodes of items painted before. Returns the same `<svg>`.
   *
   * With a `budget`, no new paint starts after the frame has painted that much weight
   * (`paintWeight`), in paint order. Items that it leaves out keep their earlier nodes, and
   * `pending` tells when some are left. A slide of 24,000 shapes then goes on screen over several
   * frames, and no frame blocks for seconds. The items in a group count, so a group of 20,000
   * shapes paints its children over several frames too. A frame can go past its budget: an item
   * that is not a group paints whole, a group with effects paints whole, and a group paints at
   * least one child each time. Extra copies of a repeated item take nothing of the budget.
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
    const frame = new Frame(layout, ctx, this.#defs, budget, this.#items);
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
    this.#items = new Map();
    this.#building.clear();
    this.#stale = new WeakSet();
    this.#twice = [];
    return { root, content };
  }

  /**
   * Step 2: paint the background again (one rectangle), first in the budgets, and drop the extra
   * copies of the last frame.
   */
  #paintBackground(frame: Frame, root: SVGSVGElement, content: SVGElement): void {
    if (this.#background) dropPainted(this.#background);
    for (const painted of this.#twice.splice(0)) dropPainted(painted);
    const { width, height, background } = frame.layout.slide;
    const painted = frame.recordBackground(() =>
      paintBackground(this.#doc, background, width, height, frame.ctx)
    );
    this.#background = painted;
    root.insertBefore(painted.node!, content);
  }

  /** Step 3: paint one layout item, or keep it, or leave it out of this frame. */
  #paintItem(frame: Frame, item: LayoutItem): void {
    // A repeat of an item has no record of its own: `#items` holds the first copy's.
    const repeat = frame.shown.has(item);
    const old = repeat ? undefined : this.#items.get(item);
    const kept = this.#current(frame, old);
    if (kept) this.#keep(frame, item, kept);
    else if (frame.room <= 0) this.#leaveOut(frame, item, old);
    else if (repeat) this.#paintCopy(frame, item);
    else this.#paintOrContinue(frame, item, old);
  }

  /**
   * `painted` when its nodes are what a new paint draws here: it is not stale, and each budget
   * that it read has as much left as when it began to read it. Call it at the item's turn, before
   * the item spends anything.
   */
  #current(frame: Frame, painted: PaintedItem | undefined): CurrentRecord | undefined {
    if (!painted || this.#stale.has(painted) || !frame.fits(painted)) return undefined;
    return painted as CurrentRecord;
  }

  /** Keep the current record of `item`. A partial group continues while the frame has room. */
  #keep(frame: Frame, item: LayoutItem, old: CurrentRecord): void {
    this.#dropBuilding(item);
    const painted =
      old.rest !== undefined && frame.room > 0 ? this.#continue(frame, item, old) : old;
    if (painted.rest !== undefined) frame.pending = true;
    frame.show(item, painted);
  }

  /** Leave `item` out of this frame, which has no room, and show its stand-in, if any. */
  #leaveOut(frame: Frame, item: LayoutItem, old: PaintedItem | undefined): void {
    frame.pending = true;
    const stand = old ? ([item, old] as const) : frame.standIn(item);
    if (stand) frame.show(stand[0], stand[1]);
  }

  /**
   * Paint an extra copy of an item that the layout holds more than once: one node cannot stand
   * in two places. The copy paints whole on each paint and takes nothing of the frame's room, so
   * the items after it still paint and the slide settles.
   */
  #paintCopy(frame: Frame, item: LayoutItem): void {
    const painted = this.#paintNew(frame, item, true);
    this.#twice.push(painted);
    frame.show(null, painted);
  }

  /**
   * Paint `item`, which has no current record: it is new to this painter, or its record is stale
   * or no longer fits. Continue its building record when that is current, or else paint it new.
   * A partial result of an item with an earlier record goes to `#building`. The earlier record
   * stays on screen until the new one is whole, and the new one holds the budget share.
   */
  #paintOrContinue(frame: Frame, item: LayoutItem, old: PaintedItem | undefined): void {
    const building = this.#current(frame, this.#building.get(item));
    let painted: PaintedItem;
    if (building) {
      painted = this.#continue(frame, item, building);
    } else {
      this.#dropBuilding(item);
      painted = this.#paintNew(frame, item, false);
    }
    if (painted.rest !== undefined) {
      frame.pending = true;
      if (old) {
        this.#building.set(item, painted);
        frame.show(item, old, painted);
        return;
      }
    }
    this.#building.delete(item);
    if (old) dropPainted(old);
    frame.show(item, painted);
  }

  /**
   * Paint `item` from nothing. An extra `copy` paints whole and takes nothing of the frame's room.
   * Otherwise, a group that weighs more than the room paints only its first children, unless it
   * has effects: its shadow and filters cover all of its children, and the cascade bounds its
   * weight (`MAX_GROUP_EFFECT_ELEMENTS`) and nesting, so it fits a frame.
   */
  #paintNew(frame: Frame, item: LayoutItem, copy: boolean): PaintedItem {
    const doc = this.#doc;
    const opts = this.#opts;
    const { ctx } = frame;
    const paint = () => paintItem(doc, item, ctx, opts, 1);
    if (copy) return frame.recordCopy(paint);
    if (item.kind === 'group' && paintWeight(item) > frame.room && !groupPaintsEffects(item)) {
      const g = groupNode(doc, item, opts);
      return frame.recordGroupStart(g, item.children.length, (room) =>
        paintChildren(doc, item, g, ctx, opts, 0, room)
      );
    }
    return frame.recordWhole(paint, paintWeight(item));
  }

  /**
   * Spend the share of the current record `painted`, and paint more of it within the frame's
   * room when it is partial. Answers the new record. The callers pass only partial records, and
   * only a group has one.
   */
  #continue(frame: Frame, item: LayoutItem, painted: CurrentRecord): PaintedItem {
    if (item.kind !== 'group') throw new Error('SlidePainter: only a group has a partial record');
    return frame.extend(painted, item.children.length, (node, from, room) =>
      paintChildren(this.#doc, item, node, frame.ctx, this.#opts, from, room)
    );
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
