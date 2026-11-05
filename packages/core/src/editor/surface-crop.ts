// Crop mode on the slide surface (`crop-mode.ts` holds the state): the whole image of the
// picture, dimmed outside its frame, and the eight crop handles on the frame. A handle
// drag crops with the image fixed on the slide; a drag on the image moves the image inside the
// frame. Each step is a `cropImage` command, and one drag is one undo step.

import { apply, invert, type LayoutItem, type Matrix } from '../layout/index.ts';
import type { ImageCrop } from '../contracts/index.ts';
import { STORED_PIXELS_STYLE, SVG_NS } from '../output/index.ts';
import { MAX_PAD, usableCrop } from './picture-commands.ts';

/** A crop handle, or `move` for a drag on the image. */
export type CropHandle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w' | 'move';

const HANDLES: readonly Exclude<CropHandle, 'move'>[] = [
  'nw',
  'n',
  'ne',
  'e',
  'se',
  's',
  'sw',
  'w',
];
const HANDLE_POS: Record<Exclude<CropHandle, 'move'>, readonly [number, number]> = {
  nw: [0, 0],
  n: [0.5, 0],
  ne: [1, 0],
  e: [1, 0.5],
  se: [1, 1],
  s: [0.5, 1],
  sw: [0, 1],
  w: [0, 0.5],
};

/** The smallest part of the image a crop handle leaves visible. */
const MIN_VISIBLE = 0.01;
/** The smallest visible part a crop can have (`cropImage`, Format Picture), for the image box. */
const MIN_CROP_VISIBLE = 0.001;

/** A crop drag in progress. */
export interface CropDrag {
  readonly handle: CropHandle;
  /** Slide to frame-local coordinates, at the start of the drag. */
  readonly inverse: Matrix;
  readonly start: { readonly x: number; readonly y: number };
  readonly crop: { readonly l: number; readonly t: number; readonly r: number; readonly b: number };
  /** The image size in frame-local units. */
  readonly imgW: number;
  readonly imgH: number;
  readonly gesture: string;
}

type Picture = Extract<LayoutItem, { kind: 'picture' }>;

/** The crop of a laid-out picture (`a:srcRect`), or null when it shows no image. */
function cropOfItem(item: LayoutItem): CropDrag['crop'] | null {
  if (item.kind !== 'picture' || item.element.fill.type !== 'picture') return null;
  // An axis that leaves nothing visible draws uncropped, as paint and the commands read it.
  return usableCrop(item.element.fill.srcRect);
}

/** The full image in frame-local coordinates. */
function imageBox(item: Picture, crop: CropDrag['crop']) {
  const { cx, cy } = item.element.xfrm;
  const w = cx / Math.max(MIN_CROP_VISIBLE, 1 - crop.l - crop.r);
  const h = cy / Math.max(MIN_CROP_VISIBLE, 1 - crop.t - crop.b);
  return { x: -crop.l * w, y: -crop.t * h, w, h };
}

const set = (el: Element, attrs: Record<string, string | number>): Element => {
  for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, String(v));
  return el;
};

/**
 * The crop mode overlay of a picture: the image outside the frame, dimmed, and the crop
 * handles. `url` is the image the slide paints; `px` is one screen pixel in slide units;
 * `idPrefix` keeps the clip id unique to the editor.
 */
export function cropOverlay(
  doc: Document,
  item: LayoutItem,
  url: string | undefined,
  px: number,
  idPrefix: string
): Element[] {
  const crop = cropOfItem(item);
  if (!crop || item.kind !== 'picture') return [];
  const { cx, cy } = item.element.xfrm;
  const img = imageBox(item, crop);
  const g = set(doc.createElementNS(SVG_NS, 'g'), {
    class: 'pptx-surface__crop',
    transform: `matrix(${item.matrix.join(' ')})`,
  });
  if (url) {
    // The dimmed image shows only outside the frame: an even-odd clip of the image less the frame.
    const clip = set(doc.createElementNS(SVG_NS, 'path'), {
      'clip-rule': 'evenodd',
      d: `M${img.x} ${img.y}h${img.w}v${img.h}h${-img.w}Z M0 0h${cx}v${cy}h${-cx}Z`,
    });
    const id = `${idPrefix}crop-${item.element.nodeId}`;
    const clipPath = set(doc.createElementNS(SVG_NS, 'clipPath'), { id });
    clipPath.append(clip);
    const image = set(doc.createElementNS(SVG_NS, 'image'), {
      class: 'pptx-surface__crop-image',
      href: url,
      x: img.x,
      y: img.y,
      width: img.w,
      height: img.h,
      preserveAspectRatio: 'none',
      // The stored pixels, as paint draws them (`STORED_PIXELS_STYLE`).
      style: STORED_PIXELS_STYLE,
      'clip-path': `url(#${id})`,
    });
    g.append(clipPath, image);
  }
  g.append(
    set(doc.createElementNS(SVG_NS, 'rect'), {
      class: 'pptx-surface__crop-extent',
      x: img.x,
      y: img.y,
      width: img.w,
      height: img.h,
      'stroke-width': px,
    }),
    set(doc.createElementNS(SVG_NS, 'rect'), {
      class: 'pptx-surface__outline',
      width: cx,
      height: Math.max(cy, 1),
      'stroke-width': px,
    })
  );
  // The crop handles: an L at each corner and a bar on each edge, inside the frame.
  const len = Math.min(px * 12, cx / 3, cy / 3);
  for (const handle of HANDLES) {
    const [hx, hy] = HANDLE_POS[handle];
    const x = hx * cx;
    const y = hy * cy;
    const ix = hx === 0 ? 1 : -1;
    const iy = hy === 0 ? 1 : -1;
    const d =
      hx !== 0.5 && hy !== 0.5
        ? `M${x} ${y + iy * len}L${x} ${y}L${x + ix * len} ${y}`
        : hx === 0.5
          ? `M${x - len} ${y}L${x + len} ${y}`
          : `M${x} ${y - len}L${x} ${y + len}`;
    // Two tones, a light halo under a dark bar, so a handle shows on any picture (WCAG 1.4.11).
    // The wider halo is the hit target too.
    g.append(
      set(doc.createElementNS(SVG_NS, 'path'), {
        class: `pptx-surface__crop-halo pptx-surface__handle--${handle}`,
        'data-crop-handle': handle,
        d,
        'stroke-width': px * 7,
      }),
      set(doc.createElementNS(SVG_NS, 'path'), {
        class: `pptx-surface__crop-handle pptx-surface__handle--${handle}`,
        'data-crop-handle': handle,
        d,
        'stroke-width': px * 4,
      })
    );
  }
  return [g];
}

/** Whether a slide point is on the picture's whole image, where a drag moves the image. */
export function onCropImage(item: LayoutItem, p: { x: number; y: number }): boolean {
  const crop = cropOfItem(item);
  if (!crop || item.kind !== 'picture') return false;
  const img = imageBox(item, crop);
  const local = apply(invert(item.matrix), p.x, p.y);
  return (
    local.x >= img.x && local.x <= img.x + img.w && local.y >= img.y && local.y <= img.y + img.h
  );
}

/** Start a crop drag on a picture, or null when it shows no image. */
export function startCropDrag(
  item: LayoutItem,
  handle: CropHandle,
  start: { x: number; y: number },
  gesture: string
): CropDrag | null {
  const crop = cropOfItem(item);
  if (!crop || item.kind !== 'picture') return null;
  const img = imageBox(item, crop);
  return { handle, inverse: invert(item.matrix), start, crop, imgW: img.w, imgH: img.h, gesture };
}

/** A `rect` crop from fractions, in the percent the command takes. */
function rect(l: number, t: number, r: number, b: number, keepFrame: boolean): ImageCrop {
  const crop = {
    kind: 'rect',
    left: l * 100,
    top: t * 100,
    right: r * 100,
    bottom: b * 100,
  } as const;
  return keepFrame ? { ...crop, keepFrame } : crop;
}

/**
 * The part of a move `d` that keeps both edges of one axis inside the crop range: the leading
 * edge (`near - d`) and the trailing edge (`far + d`) each stay below 1, the image edge, and at
 * or above `-MAX_PAD`. The visible part does not change in a move.
 */
function moveLimit(d: number, near: number, far: number): number {
  const lo = Math.max(near - 1 + MIN_CROP_VISIBLE, -MAX_PAD - far);
  const hi = Math.min(1 - far - MIN_CROP_VISIBLE, near + MAX_PAD);
  return lo > hi ? 0 : Math.max(lo, Math.min(hi, d));
}

/** The crop for the pointer at slide point `p`. */
export function cropAt(drag: CropDrag, p: { x: number; y: number }): ImageCrop {
  const a = apply(drag.inverse, drag.start.x, drag.start.y);
  const b = apply(drag.inverse, p.x, p.y);
  const du = (b.x - a.x) / drag.imgW;
  const dv = (b.y - a.y) / drag.imgH;
  const { l, t, r, b: bottom } = drag.crop;
  if (drag.handle === 'move') {
    // The frame stays; the image moves under it, so the crop shifts against the pointer. The
    // shift stops where an edge would leave the range `cropImage` takes, so the image stops at
    // the limit instead of each pointer move being refused.
    const u = moveLimit(du, l, r);
    const v = moveLimit(dv, t, bottom);
    return rect(l - u, t - v, r + u, bottom + v, true);
  }
  const [hx, hy] = HANDLE_POS[drag.handle];
  // Each moved edge stops short of the opposite one, and inside the padding limit.
  const edge = (from: number, to: number, opposite: number): number =>
    Math.max(-MAX_PAD, Math.min(from + to, 1 - opposite - MIN_VISIBLE));
  const nl = hx === 0 ? edge(l, du, r) : l;
  const nr = hx === 1 ? edge(r, -du, l) : r;
  const nt = hy === 0 ? edge(t, dv, bottom) : t;
  const nb = hy === 1 ? edge(bottom, -dv, t) : bottom;
  return rect(nl, nt, nr, nb, false);
}
