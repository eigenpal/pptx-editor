// Re-fit a table's frame after an edit: the frame's `p:xfrm/a:ext` is
// the table's laid-out size, the sum of the grid widths by the sum of the row heights after
// rows grow to fit their text (design D5 of `table-rendering`). Row heights (`a:tr/@h`) stay
// minimums; a saved deck can hold the grown height there instead (an APPROXIMATION,
// listed on `docs/site/content/pptx-fidelity.mdx`).

import {
  resolveSlide,
  type ResolvedElement,
  type ResolvedGraphicFrame,
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
    if (el.nodeId === nodeId && el.kind === 'graphicFrame') return el;
    if (el.kind === 'group') {
      const hit = findFrame(el.children, nodeId);
      if (hit) return hit;
    }
  }
  return undefined;
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
  const frame = findFrame(resolveSlide(base, slide)?.elements ?? [], nodeId);
  if (!frame?.table) return (restore(), true);
  const laid = layoutTable(frame.table, measurer);
  if (laid.rowY.length !== rows.length + 1) return (restore(), true);
  const drawn = rows.map((_, i) => laid.rowY[i + 1] - laid.rowY[i]);
  const total = laid.rowY[rows.length] - laid.rowY[0];
  const target = (total * box.cy) / bcy;
  if (total <= 0) return (restore(), true);
  if (target >= total) {
    rows.forEach((r, i) => set(r, 'h', (drawn[i] * target) / total));
    return true;
  }
  const held = rows.map((_, i) => drawn[i] > mins[i] + 1);
  const heldSum = drawn.reduce((n, d, i) => n + (held[i] ? d : 0), 0);
  const freeSum = mins.reduce((n, m, i) => n + (held[i] ? 0 : m), 0);
  const room = Math.max(target - heldSum, 1);
  rows.forEach((r, i) =>
    set(r, 'h', held[i] || freeSum <= 0 ? mins[i] : (mins[i] * room) / freeSum)
  );
  return true;
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
 * Set the size of the table frame `nodeId` on slide `index` to its laid-out table. Returns
 * false when `nodeId` is not a table frame on the slide.
 */
export function refitTable(
  batch: OpBatch,
  index: number,
  part: string,
  nodeId: string,
  measurer: TextMeasurer
): boolean {
  const slide = resolveSlide(batch.pkg, index);
  const frame = slide ? findFrame(slide.elements, nodeId) : undefined;
  if (!frame?.table || frame.layer !== 'slide') return false;
  const t = layoutTable(frame.table, measurer);
  const { xfrm } = frame;
  const cx = Math.round(t.width);
  const cy = Math.round(t.height);
  if (Math.abs(cx - xfrm.cx) > 1 || Math.abs(cy - xfrm.cy) > 1)
    setTransform(batch, part, nodeId, xfrm.x, xfrm.y, cx, cy, xfrm.rot);
  return true;
}
