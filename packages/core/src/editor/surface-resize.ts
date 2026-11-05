// The resize math of a handle drag on the slide surface: one shape in its own (rotated) frame,
// or several shapes by a handle of their common bounds. Moved out of `surface.ts` (the
// 1000-line cap).

import type { LayoutItem } from '../layout/index.ts';
import type { TransformItem } from '../contracts/index.ts';
import { toFile, toParent, type Box } from '../automation/commands/placement.ts';

const EMU_PER_PX = 9_525;

/** A handle's place on its box: 0, 0.5, or 1 of the width and of the height. */
type HandleAt = readonly [number, number];

/**
 * Resize one shape by a handle moved `ldx`, `ldy` in its parent's space, keeping the opposite
 * handle fixed. Shift keeps the ratio on a corner drag. A locked aspect ratio
 * (`noChangeAspect`) keeps it on every handle: a side handle drives the other axis, which grows
 * about the center, as DOCX Editor's image handles do (`image-overlay-geometry.ts`). @internal
 */
export function resizeOne(
  item: LayoutItem,
  x: Box,
  [hx, hy]: HandleAt,
  ldx: number,
  ldy: number,
  keep: { readonly shift: boolean; readonly locked: boolean }
): TransformItem {
  const cxm = x.x + x.cx / 2;
  const cym = x.y + x.cy / 2;
  const r = (x.rot * Math.PI) / 180;
  const cos = Math.cos(r);
  const sin = Math.sin(r);
  const local = { x: ldx * cos + ldy * sin, y: -ldx * sin + ldy * cos };
  let w = x.cx + (hx === 1 ? local.x : hx === 0 ? -local.x : 0);
  let h = x.cy + (hy === 1 ? local.y : hy === 0 ? -local.y : 0);
  const corner = hx !== 0.5 && hy !== 0.5;
  if (((keep.shift && corner) || keep.locked) && x.cx > 0 && x.cy > 0) {
    const k = corner ? Math.max(w / x.cx, h / x.cy) : hx === 0.5 ? h / x.cy : w / x.cx;
    w = x.cx * k;
    h = x.cy * k;
  }
  w = Math.max(EMU_PER_PX * 2, w);
  h = Math.max(x.cy === 0 ? 0 : EMU_PER_PX * 2, h);
  // The fixed point: the opposite handle, in parent space.
  const fx = (1 - hx) * x.cx - x.cx / 2;
  const fy = (1 - hy) * x.cy - x.cy / 2;
  const fixed = { x: cxm + fx * cos - fy * sin, y: cym + fx * sin + fy * cos };
  const nfx = (1 - hx) * w - w / 2;
  const nfy = (1 - hy) * h - h / 2;
  const ncx = fixed.x - (nfx * cos - nfy * sin);
  const ncy = fixed.y - (nfx * sin + nfy * cos);
  return toFile(item, { x: ncx - w / 2, y: ncy - h / 2, cx: w, cy: h, rot: x.rot });
}

/**
 * Resize several shapes by a handle of their common bounds (`u`, slide EMU): each shape's
 * center and size scale from the fixed side (design D4 of `group-and-multi-edit`). Rotated
 * shapes keep their rotation. @internal
 */
export function scaleMany(
  u: { readonly x: number; readonly y: number; readonly r: number; readonly b: number },
  items: readonly {
    readonly item: LayoutItem;
    readonly box: Box;
    readonly center: { x: number; y: number };
  }[],
  [hx, hy]: HandleAt,
  dx: number,
  dy: number,
  keepAspect: boolean
): TransformItem[] {
  const W = u.r - u.x;
  const H = u.b - u.y;
  let sx = W > 0 ? Math.max(0.01, (W + (hx === 1 ? dx : hx === 0 ? -dx : 0)) / W) : 1;
  let sy = H > 0 ? Math.max(0.01, (H + (hy === 1 ? dy : hy === 0 ? -dy : 0)) / H) : 1;
  if (keepAspect && hx !== 0.5 && hy !== 0.5) sx = sy = Math.max(sx, sy);
  const fixed = { x: u.x + (1 - hx) * W, y: u.y + (1 - hy) * H };
  return items.map(({ item, box, center }) => {
    const c = toParent(
      item,
      fixed.x + (center.x - fixed.x) * sx,
      fixed.y + (center.y - fixed.y) * sy
    );
    const cx = box.cx * sx;
    const cy = box.cy * sy;
    return toFile(item, { x: c.x - cx / 2, y: c.y - cy / 2, cx, cy, rot: box.rot });
  });
}
