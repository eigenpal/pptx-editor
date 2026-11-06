// The adjust handles of a selected preset shape (ECMA-376 Part 1, 20.1.9.1 `a:ahLst`): a
// diamond at each handle place, and the drag that moves it. Each pointer move finds the adjust
// values that put the handle nearest to the pointer (`adjustForHandle`, from the preset's own
// formulas) and writes them; the steps of one drag are one undo step, as a crop drag's are.

import {
  adjustForHandle,
  apply,
  invert,
  presetAdjustments,
  presetHandles,
  type LayoutItem,
  type Matrix,
} from '../layout/index.ts';
import { SVG_NS } from '../output/index.ts';

/** The preset geometry of an item whose handles the canvas shows, or null. */
function presetOf(
  item: LayoutItem | undefined
): { preset: string; adjust: ReadonlyMap<string, number>; w: number; h: number } | null {
  if (!item || (item.kind !== 'shape' && item.kind !== 'picture')) return null;
  const el = item.element;
  if (el.kind !== 'shape' && el.kind !== 'picture') return null;
  if (el.geometry.kind !== 'preset') return null;
  return { preset: el.geometry.preset, adjust: el.geometry.adjust, w: el.xfrm.cx, h: el.xfrm.cy };
}

/** The diamonds of the adjust handles of `item`, in slide units (`px` is one CSS pixel). */
export function adjustHandleElements(
  doc: Document,
  item: LayoutItem | undefined,
  px: number,
  label: string
): Element[] {
  const g = presetOf(item);
  if (!item || !g) return [];
  return presetHandles(g.preset, g.w, g.h, g.adjust).map((h, i) => {
    const p = apply(item.matrix, h.x, h.y);
    const r = px * 5;
    const node = doc.createElementNS(SVG_NS, 'path');
    node.setAttribute('class', 'pptx-surface__handle pptx-surface__handle--adjust');
    node.setAttribute('data-adjust-handle', String(i));
    node.setAttribute(
      'd',
      `M${p.x} ${p.y - r}L${p.x + r} ${p.y}L${p.x} ${p.y + r}L${p.x - r} ${p.y}Z`
    );
    node.setAttribute('stroke-width', String(px));
    const title = doc.createElementNS(SVG_NS, 'title');
    title.textContent = label;
    node.append(title);
    return node;
  });
}

/** A drag of one adjust handle. */
export interface AdjustDrag {
  readonly nodeId: string;
  readonly index: number;
  /** Slide space to the shape's own box (EMU), as the press found it. */
  readonly inverse: Matrix;
  readonly preset: string;
  readonly adjust: ReadonlyMap<string, number>;
  readonly w: number;
  readonly h: number;
  readonly gesture: string;
  /** The values the last move wrote, so a move that changes nothing writes nothing. */
  last: string;
}

/** A drag of handle `index` of `item`, or null when it has no such handle. */
export function startAdjustDrag(
  item: LayoutItem | undefined,
  index: number,
  gesture: string
): AdjustDrag | null {
  const g = presetOf(item);
  const handle = g ? presetHandles(g.preset, g.w, g.h, g.adjust)[index] : undefined;
  if (!item || !g || !handle) return null;
  // The values the shape shows now, so a move that keeps them writes nothing.
  const shown = new Map(
    (presetAdjustments(g.preset, g.w, g.h, g.adjust) ?? [])
      .filter((a) => handle.guides.includes(a.name))
      .map((a) => [a.name, Math.round(a.value)] as const)
  );
  const last = JSON.stringify([...new Map([...g.adjust, ...shown])]);
  return {
    nodeId: item.element.nodeId,
    index,
    inverse: invert(item.matrix),
    ...g,
    gesture,
    last,
  };
}

/**
 * The adjust values for the pointer at slide point `p`, or null when they are the values the
 * last move wrote.
 */
export function adjustAt(
  d: AdjustDrag,
  p: { x: number; y: number }
): Record<string, number> | null {
  const local = apply(d.inverse, p.x, p.y);
  const values = adjustForHandle(d.preset, d.w, d.h, d.adjust, d.index, local);
  if (!values) return null;
  const key = JSON.stringify([...new Map([...d.adjust, ...values])]);
  if (key === d.last) return null;
  d.last = key;
  return Object.fromEntries(values);
}
