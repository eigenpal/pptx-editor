// Re-fit a table's frame after an edit: the frame's `p:xfrm/a:ext` is
// the table's laid-out size, the sum of the grid widths by the sum of the row heights after
// rows grow to fit their text (design D5 of `table-rendering`). Row heights (`a:tr/@h`) stay
// minimums; a saved deck can hold the grown height there instead (an APPROXIMATION,
// listed on `docs/site/content/pptx-fidelity.mdx`).

import {
  resolveSlide,
  type ResolvedElement,
  type ResolvedGraphicFrame,
  type ResolvedSlide,
} from '../../cascade/index.ts';
import { layoutTable, type TextMeasurer } from '../../layout/index.ts';
import {
  attr,
  child,
  children,
  elements,
  intAttr,
  path,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../../store/index.ts';
import type { OpBatch } from './op-batch.ts';
import { setAttr, tblOf } from './table-model.ts';
import { setTransform } from './shape-commands.ts';

function findFrame(
  elements: readonly ResolvedElement[],
  nodeId: string
): ResolvedGraphicFrame | undefined {
  for (const el of elements) {
    // Node ids repeat across parts: a layout or master frame can have the same one.
    if (el.layer !== 'slide') continue;
    if (el.nodeId === nodeId && el.kind === 'graphicFrame') return el;
    if (el.kind === 'group') {
      const hit = findFrame(el.children, nodeId);
      if (hit) return hit;
    }
  }
  return undefined;
}

/** The slide's table frames of a resolved slide by node id, built once. */
const frameTables = new WeakMap<ResolvedSlide, Map<string, ResolvedGraphicFrame>>();

function frameTable(slide: ResolvedSlide): Map<string, ResolvedGraphicFrame> {
  let table = frameTables.get(slide);
  if (table) return table;
  table = new Map();
  const walk = (elements: readonly ResolvedElement[]): void => {
    for (const el of elements) {
      if (el.layer !== 'slide') continue;
      if (el.kind === 'graphicFrame' && !table!.has(el.nodeId)) table!.set(el.nodeId, el);
      if (el.kind === 'group') walk(el.children);
    }
  };
  walk(slide.elements);
  frameTables.set(slide, table);
  return table;
}

/** `resolveSlide` of a package that never changes (a drag's start), once per slide. */
const baseSlides = new WeakMap<PptxPackage, Map<number, ResolvedSlide | undefined>>();

function resolveBase(base: PptxPackage, slide: number): ResolvedSlide | undefined {
  let known = baseSlides.get(base);
  if (!known) baseSlides.set(base, (known = new Map()));
  if (!known.has(slide)) known.set(slide, resolveSlide(base, slide));
  return known.get(slide);
}

/**
 * Resize the table in frame `nodeId` to the box `cx` by `cy`, when a table is resized by its
 * handles (an APPROXIMATION of how the change is shared among rows and columns).
 * Every size comes from `base`, the package when the drag started, so one box gives one table
 * whatever path the pointer took. An axis changes only when the box differs from the frame's
 * box in `base` (`p:xfrm/a:ext`), so a width drag keeps every row, and it changes by that ratio,
 * so a frame size that differs from the drawn table still scales as the user sees it. Columns
 * scale by the ratio. Rows grow in proportion to their drawn heights; to shrink, rows that text
 * holds above their minimum keep their height and the others share what is left. Returns false
 * when the frame holds no table. No size gets smaller than 1 EMU.
 */
export function resizeTable(
  batch: OpBatch,
  slide: number,
  part: string,
  nodeId: string,
  box: { readonly cx: number; readonly cy: number },
  base: PptxPackage,
  measurer: TextMeasurer
): boolean {
  const tbl = tblOf(batch.element(part, nodeId));
  const from = xmlPartOf(base, part) ? findElement(xmlPartOf(base, part)!.root, nodeId) : undefined;
  const baseTbl = tblOf(from);
  if (!tbl || !baseTbl) return !!tbl;
  const ext = path(from, 'p:xfrm', 'a:ext');
  const bcx = intAttr(ext, 'cx') ?? 0;
  const bcy = intAttr(ext, 'cy') ?? 0;
  const set = (el: Element, name: string, v: number): void => {
    const value = String(Math.max(1, Math.round(v)));
    const now = batch.element(part, el.id);
    if (attr(now, name) !== value) setAttr(batch, part, el.id, name, value);
  };
  const fx = bcx > 0 && Math.abs(box.cx - bcx) > 1 ? box.cx / bcx : 1;
  for (const c of children(child(baseTbl, 'a:tblGrid'), 'a:gridCol')) {
    const w = intAttr(c, 'w');
    if (w !== undefined) set(c, 'w', w * fx);
  }
  const rows = children(baseTbl, 'a:tr');
  const mins = rows.map((r) => intAttr(r, 'h') ?? 0);
  const restore = (): void => rows.forEach((r, i) => set(r, 'h', mins[i]));
  if (!(bcy > 0 && Math.abs(box.cy - bcy) > 1) || !Number.isFinite(box.cy) || box.cy <= 0)
    return (restore(), true);
  const frame = findFrame(resolveBase(base, slide)?.elements ?? [], nodeId);
  if (!frame?.table) return (restore(), true);
  const laid = layoutTable(frame.table, measurer);
  if (laid.rowY.length !== rows.length + 1) return (restore(), true);
  const drawn = rows.map((_, i) => laid.rowY[i + 1] - laid.rowY[i]);
  const total = laid.rowY[rows.length] - laid.rowY[0];
  if (total <= 0) return (restore(), true);
  const heights = scaledRowHeights(drawn, mins, (total * box.cy) / bcy);
  rows.forEach((r, i) => set(r, 'h', heights[i]!));
  return true;
}

/**
 * The row heights of a table resized to `target` in height, from the rows' `drawn` heights and
 * their minimums (`a:tr/@h`): rows grow in proportion to their drawn heights; to shrink, rows
 * that text holds above their minimum keep their height and the others share what is left.
 * `resizeTable` writes them, and the drag preview lays the table out with them.
 */
export function scaledRowHeights(
  drawn: readonly number[],
  mins: readonly number[],
  target: number
): number[] {
  const total = drawn.reduce((n, d) => n + d, 0);
  if (total <= 0) return [...mins];
  if (target >= total) return drawn.map((d) => (d * target) / total);
  const held = drawn.map((d, i) => d > mins[i]! + 1);
  const heldSum = drawn.reduce((n, d, i) => n + (held[i] ? d : 0), 0);
  const freeSum = mins.reduce((n, m, i) => n + (held[i] ? 0 : m), 0);
  const room = Math.max(target - heldSum, 1);
  return mins.map((m, i) => (held[i] || freeSum <= 0 ? m : (m * room) / freeSum));
}

/** The element `id` under `root`, or undefined. */
function findElement(root: Element, id: string): Element | undefined {
  if (root.id === id) return root;
  for (const c of elements(root)) {
    const hit = findElement(c, id);
    if (hit) return hit;
  }
  return undefined;
}

/**
 * The shapes and table frames that each measurer fit, by part name. A deck that measures with
 * the fixed measurer until the text engine starts fits exactly these again (`late-fit.ts`).
 */
const fitWith = new WeakMap<TextMeasurer, Map<string, Set<string>>>();

/** Record that `measurer` fit `nodeId` in `part`. @internal */
export function noteFit(measurer: TextMeasurer, part: string, nodeId: string): void {
  let parts = fitWith.get(measurer);
  if (!parts) fitWith.set(measurer, (parts = new Map()));
  let ids = parts.get(part);
  if (!ids) parts.set(part, (ids = new Set()));
  ids.add(nodeId);
}

/** Give `to` the fits of `from`: a rebuilt fixed measurer keeps them. @internal */
export function carryFits(from: TextMeasurer, to: TextMeasurer): void {
  for (const [part, ids] of fitWith.get(from) ?? []) for (const id of ids) noteFit(to, part, id);
}

/** The shapes and table frames that `measurer` fit, by part name. @internal */
export const fittedWith = (measurer: TextMeasurer): ReadonlyMap<string, ReadonlySet<string>> =>
  fitWith.get(measurer) ?? new Map();

/**
 * Set the size of the table frame `nodeId` on slide `index` to its laid-out table. Returns
 * false when `nodeId` is not a table frame on the slide.
 */
export function refitTable(
  batch: OpBatch,
  index: number,
  part: string,
  nodeId: string,
  measurer: TextMeasurer,
  resolved?: ResolvedSlide
): boolean {
  // Many tables re-fit from one resolve (`refitAll`): a re-fit changes only its own frame.
  const frame = resolved
    ? frameTable(resolved).get(nodeId)
    : findFrame(resolveSlide(batch.pkg, index)?.elements ?? [], nodeId);
  if (!frame?.table || frame.layer !== 'slide') return false;
  noteFit(measurer, part, nodeId);
  const t = layoutTable(frame.table, measurer);
  const { xfrm } = frame;
  const cx = Math.round(t.width);
  const cy = Math.round(t.height);
  if (Math.abs(cx - xfrm.cx) > 1 || Math.abs(cy - xfrm.cy) > 1)
    setTransform(batch, part, nodeId, xfrm.x, xfrm.y, cx, cy, xfrm.rot);
  return true;
}
