// A pointer drag of selected shapes: move, resize, rotate, and a line's end points.
//
// Each pointer move computes the new boxes (with smart-guide snapping) but writes nothing: the
// surface shows a preview (drag-preview.ts). The release commits the last boxes as one store
// transaction, so the drag is one undo step and one set of ops for collaboration. The slide,
// the toolbar, and the thumbnails do not lay out or repaint while the pointer moves. Crop drags
// (surface-crop.ts) and table grid line drags (surface-table.ts) are not shape drags: they
// write on each move.
//
// `ShapeDrag` holds the drag from press to release. Escape, the window losing focus, an undo,
// or a change to a dragged shape (a peer's edit) cancels it: the preview goes and nothing is
// written. An edit to another shape keeps the drag; its smart guides follow the new layout.

import type { TransformItem } from '../contracts/index.ts';
import type { ResolvedElement } from '../cascade/index.ts';
import { apply, type LayoutItem, type SlideLayout } from '../layout/index.ts';
import type { PaintOptions } from '../output/index.ts';
import { PaintedPreview, preview, previewLayout, type Relayout } from './drag-preview.ts';
import {
  boxOf,
  deltaToParent,
  findItem,
  lineEnds,
  lineTransform,
  rotationAfterDrag,
  slideCenter,
  toFile,
  toParent,
  unionBounds,
  type Box,
} from '../automation/commands/placement.ts';
import {
  byDrawingId,
  connectorRoute,
  gluedSite,
  elbowCorners,
  isConnector,
  moveElbowSegment,
  segmentAxis,
  sitesNear,
  type RouteEnd,
} from '../automation/commands/connectors.ts';
import {
  snapContext,
  snapMove,
  snapResize,
  type DragGuide,
  type SnapContext,
} from './surface-guides.ts';
import type { SnapGuide } from './snapping.ts';
import { resizeOne, scaleMany } from './surface-resize.ts';
import { clampTableResize } from './surface-table.ts';

/** A segment of an elbow or curve connector's route, by its index from the start. */
export type SegmentHandle = `segment${0 | 1 | 2 | 3 | 4}`;
export type Handle =
  'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w' | 'rot' | 'start' | 'end' | SegmentHandle;
export type BoxHandle = Exclude<Handle, 'start' | 'end' | SegmentHandle>;

/** Whether a handle moves a segment of a connector's route. */
export function isSegment(handle: Handle | 'move'): handle is SegmentHandle {
  return /^segment\d$/.test(handle);
}
export const HANDLES: readonly BoxHandle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w', 'rot'];
export const HANDLE_POS: Record<BoxHandle, [number, number]> = {
  nw: [0, 0],
  n: [0.5, 0],
  ne: [1, 0],
  e: [1, 0.5],
  se: [1, 1],
  s: [0.5, 1],
  sw: [0, 1],
  w: [0, 0.5],
  rot: [0.5, 0],
};

/** A shape being dragged: its laid-out box in parent space and its slide-space center. */
interface DragItem {
  readonly item: LayoutItem;
  readonly box: Box;
  readonly center: { x: number; y: number };
}

/** Bounds as `x`, `y`, right, bottom, in slide EMU. */
type Bounds = { x: number; y: number; r: number; b: number };

/** A pointer drag of a selection or one of its handles. */
export interface Drag {
  readonly handle: Handle | 'move';
  readonly items: readonly DragItem[];
  readonly start: { x: number; y: number };
  readonly bounds: Bounds;
  readonly gesture: string;
  /** The smart guide targets: the other shapes, from the layout the drag last saw. */
  snap: SnapContext;
  /** The boxes the last pointer move previews, which the release writes. */
  pending: readonly TransformItem[] | null;
  /** Each previewed shape laid out at its new box (drag-preview.ts). */
  shown: ReadonlyMap<string, LayoutItem>;
  /** The slide's items when the drag started: connector ends glue to them. */
  readonly slideItems: readonly LayoutItem[];
}

/** A drag of the items `ids` by `handle`, or null when none of them is on the slide. */
export function startDrag(
  layout: SlideLayout | undefined,
  handle: Handle | 'move',
  ids: readonly string[],
  start: { x: number; y: number },
  gesture: string
): Drag | null {
  const items = ids
    .map((id) => (layout ? findItem(layout.items, id) : undefined))
    .filter((i): i is LayoutItem => !!i)
    .map((item) => ({ item, box: boxOf(item), center: slideCenter(item) }));
  if (items.length === 0) return null;
  return {
    handle,
    items,
    start,
    bounds: unionBounds(items.map((i) => i.item)),
    gesture,
    snap: layout ? snapContext(layout, ids) : { targets: [], slide: { width: 0, height: 0 } },
    pending: null,
    shown: new Map(),
    slideItems: layout?.items ?? [],
  };
}

/** What a drag step reads of the pointer and the editor. */
export interface StepInput {
  readonly shift: boolean;
  /** The snap distance in EMU, or null when snapping is off. */
  readonly snapTo: number | null;
  readonly aspectLocked: (nodeId: string) => boolean;
  /** How near a connection site a connector end glues to it, in EMU; null or absent: never. */
  readonly glueTo?: number | null;
}

/** The boxes after the pointer moved to `p`, and the smart guides they snapped to. */
export function dragStep(
  d: Drag,
  p: { x: number; y: number },
  input: StepInput
): { items: TransformItem[]; guides: readonly DragGuide[] } {
  let dx = p.x - d.start.x;
  let dy = p.y - d.start.y;
  let guides: readonly SnapGuide[] = [];
  if (d.handle === 'move') {
    const s = snapMove(d.snap, d.bounds, dx, dy, input.shift, input.snapTo);
    ({ dx, dy, guides } = s);
    const items = d.items.map(({ item, box }) => {
      const dp = deltaToParent(item, dx, dy);
      return toFile(item, { ...box, x: box.x + dp.x, y: box.y + dp.y });
    });
    return { items, guides };
  }
  if ((d.handle === 'start' || d.handle === 'end') && isConnector(d.items[0].item))
    return moveConnectorEnd(d, d.handle, p, input);
  if (d.handle === 'start' || d.handle === 'end')
    return { items: [moveLineEnd(d, p, input.shift)], guides };
  if (isSegment(d.handle)) {
    // A route the presets cannot draw keeps the last route the drag showed.
    const t = moveElbowSegment(
      d.items[0].item,
      Number(d.handle.slice(7)),
      d.start,
      p,
      input.snapTo
    );
    return { items: t ? [t] : [...(d.pending ?? [])], guides };
  }
  const [hx, hy] = HANDLE_POS[d.handle];
  // A resize snaps its dragged edges.
  if (d.handle !== 'rot') {
    const items = d.items.map((i) => i.item);
    const s = snapResize(d.snap, items, d.bounds, [hx, hy], [dx, dy], input.shift, input.snapTo);
    ({ dx, dy, guides } = s);
  }
  if (d.items.length > 1)
    return { items: scaleMany(d.bounds, d.items, [hx, hy], dx, dy, input.shift), guides };
  // One shape: work in its parent's space, where its box lives (`surface-resize.ts`).
  const { item, box: x } = d.items[0];
  if (d.handle === 'rot') {
    const rot = rotationAfterDrag(item, x, d.start, p, input.shift);
    return { items: [toFile(item, { ...x, rot })], guides };
  }
  const keep = { shift: input.shift, locked: input.aspectLocked(item.element.nodeId) };
  // A table resizes from its drawn grid, and each dragged edge stops at the slide's edge.
  const aspect = keep.shift || keep.locked;
  ({ dx, dy } = clampTableResize(item, [hx, hy], dx, dy, d.snap.slide, aspect));
  const s0 = toParent(item, d.start.x, d.start.y);
  const s1 = toParent(item, d.start.x + dx, d.start.y + dy);
  return { items: [resizeOne(item, x, [hx, hy], s1.x - s0.x, s1.y - s0.y, keep)], guides };
}

/** A line with one end at the pointer; Shift snaps its angle to 45° steps. */
function moveLineEnd(d: Drag, p: { x: number; y: number }, snap: boolean): TransformItem {
  const { item } = d.items[0];
  const ends = lineEnds(item)!;
  const end = d.handle === 'start' ? ends.end : ends.start;
  const fixed = toParent(item, end.x, end.y);
  let moved = toParent(item, p.x, p.y);
  if (snap) {
    const len = Math.hypot(moved.x - fixed.x, moved.y - fixed.y);
    const angle =
      Math.round(Math.atan2(moved.y - fixed.y, moved.x - fixed.x) / (Math.PI / 4)) * (Math.PI / 4);
    moved = { x: fixed.x + Math.cos(angle) * len, y: fixed.y + Math.sin(angle) * len };
  }
  return d.handle === 'start'
    ? lineTransform(item, moved, fixed)
    : lineTransform(item, fixed, moved);
}

/**
 * A connector with one end at the pointer. Near a shape, the shape's connection sites show,
 * and the end glues to the nearest one within `glueTo`. An elbow or a curve with a glued end
 * routes from site to site (`connectorRoute`); else the connector stretches as a line does.
 */
function moveConnectorEnd(
  d: Drag,
  which: 'start' | 'end',
  p: { x: number; y: number },
  input: StepInput
): { items: TransformItem[]; guides: readonly DragGuide[] } {
  const { item } = d.items[0];
  const other = which === 'start' ? 'end' : 'start';
  const near =
    input.glueTo != null
      ? sitesNear(d.slideItems, p, input.glueTo, new Set([item.element.nodeId]))
      : { shown: [], snap: null };
  const site = near.snap;
  const fixed = gluedSite(item, other, byDrawingId(d.slideItems));
  const ends = lineEnds(item)!;
  const free: RouteEnd = site
    ? { p: site, dir: { x: site.dx, y: site.dy } }
    : { p: input.shift ? angleSnapped(ends[other], p) : p };
  const kept: RouteEnd = fixed
    ? { p: fixed, dir: { x: fixed.dx, y: fixed.dy } }
    : { p: ends[other] };
  const t = which === 'start' ? connectorRoute(item, free, kept) : connectorRoute(item, kept, free);
  const glue = site ? { nodeId: site.nodeId, index: site.index } : null;
  const guides: DragGuide[] = near.shown.map((s) => ({
    kind: 'site',
    x: s.x,
    y: s.y,
    glued: s === site,
  }));
  return { items: [{ ...t, connector: { ...t.connector, [which]: glue } }], guides };
}

/** `p` on the nearest 45° line through `from` (Shift). */
function angleSnapped(
  from: { x: number; y: number },
  p: { x: number; y: number }
): { x: number; y: number } {
  const len = Math.hypot(p.x - from.x, p.y - from.y);
  const angle = Math.round(Math.atan2(p.y - from.y, p.x - from.x) / (Math.PI / 4)) * (Math.PI / 4);
  return { x: from.x + Math.cos(angle) * len, y: from.y + Math.sin(angle) * len };
}

/** What a shape drag asks of the surface. */
export interface DragHost {
  label(key: string): string;
  layout(): SlideLayout | undefined;
  /** Write the boxes on release, in file coordinates: one undo step. */
  transform(items: readonly TransformItem[], gesture: string): void;
  relayout?: Relayout;
  paintLayout?: (layout: SlideLayout) => SlideLayout;
  /** The painted slide, where the preview goes. */
  slide(): Element | null;
  paintOptions(layout: SlideLayout): PaintOptions;
  /** Render the surface again. */
  repaint(): void;
  /** Draw the selection chrome over `layout` (the preview's, while a drag previews). */
  overlay(layout: SlideLayout): void;
}

type Point = { x: number; y: number };

/** A cascaded element as text, to compare two that are not the same object. */
function signOf(element: ResolvedElement): string {
  return JSON.stringify(element, (_, v: unknown) =>
    v instanceof Map || v instanceof Set ? [...v] : ArrayBuffer.isView(v) ? v.byteLength : v
  );
}

/** The shape drag of one surface, from the press to the release (see the file comment). */
export class ShapeDrag {
  #drag: Drag | null = null;
  /** The layout the drag last checked, and each dragged shape as the press laid it out. */
  #seen: { layout: SlideLayout | undefined; shapes: Map<string, LayoutItem> } | null = null;
  readonly #preview = new PaintedPreview();
  /** A cancel the live region says in place of the selection (`over`), until it changes. */
  #announce: { text: string; over?: string; state: 'new' | 'wait' | 'shown' } | null = null;
  #guides: readonly DragGuide[] = [];

  constructor(
    readonly host: DragHost,
    readonly doc: Document
  ) {
    doc.addEventListener('keydown', this.#onKeyDown, true);
    doc.defaultView?.addEventListener('blur', this.#onBlur);
  }

  dispose(): void {
    this.doc.removeEventListener('keydown', this.#onKeyDown, true);
    this.doc.defaultView?.removeEventListener('blur', this.#onBlur);
    // A drag in progress ends here; a preview writes nothing.
    this.#end();
  }

  get active(): boolean {
    return this.#drag !== null;
  }

  /** The drag turns its shape by the rotation handle. */
  get rotating(): boolean {
    return this.#drag?.handle === 'rot';
  }

  /**
   * Keep the drag's pointer on the whole surface, off the handle too: the rotate pointer
   * (`pptx-surface--rotating`), or a segment's resize pointer (`pptx-surface--segment-x`, `-y`).
   */
  pointer(classes: DOMTokenList): void {
    classes.toggle('pptx-surface--rotating', this.rotating);
    const axis = this.segmentAxis;
    classes.toggle('pptx-surface--segment-x', axis === 'x');
    classes.toggle('pptx-surface--segment-y', axis === 'y');
  }

  /** The axis a connector segment drag moves along, or null for another drag. */
  get segmentAxis(): 'x' | 'y' | null {
    const d = this.#drag;
    if (!d || !isSegment(d.handle)) return null;
    const { item } = d.items[0];
    const corners = elbowCorners(item);
    const i = Number(d.handle.slice(7));
    const [a, b] = [corners?.[i], corners?.[i + 1]];
    // On the slide, as the handle's own pointer: a turned group turns the segment too.
    const slide = (q: Point): Point => apply(item.parentMatrix, q.x, q.y);
    return a && b ? segmentAxis(slide(a), slide(b)) : null;
  }

  /** The smart guides of the last pointer move. */
  get guides(): readonly DragGuide[] {
    return this.#guides;
  }

  /** Start dragging the items `ids` (or a handle of their box) from slide point `start`. */
  start(handle: Handle | 'move', ids: readonly string[], start: Point, gesture: string): void {
    const layout = this.host.layout();
    this.#drag = startDrag(layout, handle, ids, start, gesture);
    this.#announce = null;
    const shapes = new Map(this.#drag?.items.map((i) => [i.item.element.nodeId, i.item]));
    this.#seen = { layout, shapes };
  }

  /** A pointer move to `p`: preview the new boxes, once the pointer moved `slop` EMU. */
  move(layout: SlideLayout, p: Point, slop: number, input: StepInput): void {
    const d = this.#drag;
    if (!d || Math.abs(p.x - d.start.x) + Math.abs(p.y - d.start.y) < slop) return;
    const step = dragStep(d, p, input);
    this.#guides = step.guides;
    preview(d, step.items, this.host.relayout);
    this.#show(layout);
    this.host.overlay(previewLayout(layout, d));
  }

  /** Before a render of `layout`: a drag whose shapes changed ends (see the file comment). */
  check(layout: SlideLayout | undefined): void {
    const d = this.#drag;
    if (!d || this.#seen?.layout === layout) return;
    if (this.#changed(layout)) return this.#drop();
    if (layout) d.snap = snapContext(layout, [...(this.#seen?.shapes.keys() ?? [])]);
  }

  /** After the slide was painted again: the preview goes on the new painting. */
  /** Put the painted slide back as it was painted, before the painter reuses its nodes. */
  beforeRepaint(): void {
    this.#preview.clear();
  }

  repainted(layout: SlideLayout): void {
    this.#preview.clear();
    if (this.#drag) this.#show(layout);
  }

  /** The layout the overlay draws: the preview's while a drag previews. */
  overlayLayout(layout: SlideLayout): SlideLayout {
    return this.#drag ? previewLayout(layout, this.#drag) : layout;
  }

  /** End the drag with no edit, and render. */
  cancel(): void {
    if (!this.#drag) return;
    this.#drop();
    this.host.repaint();
  }

  /** The pointer went up (or the browser took it): write the last boxes as one edit. */
  release(): void {
    const d = this.#drag;
    if (!d) return;
    // A shape that changed since the last render cancels the drag, as a render would.
    if (d.pending && this.#changed(this.host.layout())) return this.cancel();
    // The preview goes first: a refused edit repaints nothing, and the shapes must not stay
    // where the preview put them.
    this.#end();
    if (d.pending) this.host.transform(d.pending, d.gesture);
    const guides = this.#guides.length > 0;
    this.#guides = [];
    if (guides || d.pending) this.host.repaint();
  }

  /**
   * Write the live region: `said` (the selection), or a cancel in its place until the
   * selection changes. A cancel clears the region first and is set on the next frame, so a
   * screen reader says each cancel, also the same words twice.
   */
  speak(live: HTMLElement, said: string): void {
    const a = this.#announce;
    if (a) a.over ??= said;
    if (a && a.over !== said) this.#announce = null;
    const now = this.#announce;
    if (now?.state === 'wait') return;
    if (now?.state === 'new') {
      now.state = 'wait';
      live.textContent = '';
      const show = (): void => {
        if (this.#announce !== now) return;
        now.state = 'shown';
        live.textContent = now.text;
      };
      const win = live.ownerDocument.defaultView;
      if (win?.requestAnimationFrame) win.requestAnimationFrame(show);
      else setTimeout(show, 0);
      return;
    }
    const say = now?.text ?? said;
    if (live.textContent !== say) live.textContent = say;
  }

  /**
   * Whether a dragged shape is gone from `layout`, or differs from the shape the press took:
   * its own element, or the place its groups give it (an ancestor group moved or scaled).
   */
  #changed(layout: SlideLayout | undefined): boolean {
    const seen = this.#seen;
    if (!seen || seen.layout === layout) return false;
    seen.layout = layout;
    for (const [id, was] of seen.shapes) {
      const now = layout ? findItem(layout.items, id) : undefined;
      if (!now) return true;
      if (now.element !== was.element && signOf(now.element) !== signOf(was.element)) return true;
      const placed = (i: LayoutItem) => JSON.stringify([i.parentMatrix, i.toSource]);
      if (placed(now) !== placed(was)) return true;
    }
    return false;
  }

  #show(layout: SlideLayout): void {
    const d = this.#drag;
    if (!d) return;
    const options = this.host.paintOptions(layout);
    this.#preview.show(this.host.slide(), layout, d, options, this.host.paintLayout);
  }

  /** End the drag with no edit; a drag that moved says so in the live region. */
  #drop(): void {
    if (this.#drag?.pending)
      this.#announce = { text: this.host.label('canvas.dragCanceled'), state: 'new' };
    this.#end();
    this.#guides = [];
  }

  #end(): void {
    this.#drag = null;
    this.#seen = null;
    this.#preview.clear();
  }

  /** Escape during a shape drag drops its preview and writes nothing. */
  #onKeyDown = (e: KeyboardEvent): void => {
    if (e.key !== 'Escape' || !this.#drag) return;
    e.preventDefault();
    e.stopPropagation();
    this.cancel();
  };

  /** The window lost focus during a shape drag: it ends as Escape ends it. */
  #onBlur = (): void => this.cancel();
}
