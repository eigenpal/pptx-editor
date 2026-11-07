// The picture of a slide, or of one shape of it, for the automation host's `getImageAsBase64`.
// It draws from the same paint as the slide picture export (`export-image.ts`): no second
// renderer. A shape draws alone, without the slide background, and its picture is the box of its
// geometry, so an outline or an effect outside that box is cut off.
//
// In a page, the browser draws the SVG on a canvas (`renderImagePng`). With no DOM, the paint
// goes into a DOM-free SVG document, and a rasterizer that the host supplies turns its text into
// PNG bytes (`renderImageWith`). The engine checks that the answer is a PNG of the asked size.

import { apply, type LayoutItem, type SlideLayout } from '../layout/index.ts';
import { imageSize } from '../store/index.ts';
import { createSvgDocument, elementsWithAttribute, serializeSvgElement } from '../output/index.ts';
import { CommandRefusal } from '../automation/commands/op-batch.ts';
import type { DocumentSession } from './document-session.ts';
import { drawPng, mayTravel, paintLayout } from './export-image.ts';
import type { RegisteredFace } from '../layout/index.ts';

/** What `renderImagePng` draws: a slide, or one shape of it, fit into a box of pixels. */
export interface ImageRender {
  /** The node id of a shape of the slide; absent for the whole slide. */
  readonly shapeNodeId?: string;
  /** The box the picture fits into, keeping its aspect ratio. Absent: its own size. */
  readonly width?: number;
  readonly height?: number;
}

/** EMU per CSS pixel (96 per inch): the picture's own size. */
const EMU_PER_PX = 9525;
/** The longest side the renderer draws, as for the slide picture export. */
const MAX_SIDE = 8192;

/** The items on the way to one shape: its groups, each with only that branch. */
function branchTo(items: readonly LayoutItem[], nodeId: string, depth = 0): LayoutItem[] {
  if (depth > 32) return [];
  for (const item of items) {
    if (item.element.nodeId === nodeId) return [item];
    if (item.kind !== 'group') continue;
    const inner = branchTo(item.children, nodeId, depth + 1);
    if (inner.length) return [{ ...item, children: inner }];
  }
  return [];
}

/** The item of the shape itself, at the end of a branch. */
function leaf(items: readonly LayoutItem[], nodeId: string): LayoutItem | undefined {
  const item = items[0];
  if (!item || item.element.nodeId === nodeId) return item;
  return item.kind === 'group' ? leaf(item.children, nodeId) : undefined;
}

/** The slide-space bounds of an item's box, through its matrix (rotation and flips). */
function boundsOf(item: LayoutItem): { x: number; y: number; w: number; h: number } {
  const { cx, cy } = item.element.xfrm;
  const corners = [
    apply(item.matrix, 0, 0),
    apply(item.matrix, cx, 0),
    apply(item.matrix, 0, cy),
    apply(item.matrix, cx, cy),
  ];
  const xs = corners.map((p) => p.x);
  const ys = corners.map((p) => p.y);
  let x = Math.min(xs[0]!, xs[1]!, xs[2]!, xs[3]!);
  let y = Math.min(ys[0]!, ys[1]!, ys[2]!, ys[3]!);
  let w = Math.max(xs[0]!, xs[1]!, xs[2]!, xs[3]!) - x;
  let h = Math.max(ys[0]!, ys[1]!, ys[2]!, ys[3]!) - y;
  // A line has no height (or no width): give it one pixel, so the picture is not empty.
  if (w < EMU_PER_PX) [x, w] = [x - (EMU_PER_PX - w) / 2, EMU_PER_PX];
  if (h < EMU_PER_PX) [y, h] = [y - (EMU_PER_PX - h) / 2, EMU_PER_PX];
  return { x, y, w, h };
}

/**
 * A picture of `w` by `h` pixels fit into the asked box, keeping its aspect ratio: both sides
 * scale to fit, one side sets the scale, and no side gives the own size. A size larger than the
 * renderer draws is clamped to it, as Office.js clamps.
 */
export function fitSize(w: number, h: number, request: ImageRender): [number, number] {
  let scale = 1;
  if (request.width !== undefined && request.height !== undefined)
    scale = Math.min(request.width / w, request.height / h);
  else if (request.width !== undefined) scale = request.width / w;
  else if (request.height !== undefined) scale = request.height / h;
  scale = Math.min(scale, MAX_SIDE / w, MAX_SIDE / h);
  return [
    Math.min(MAX_SIDE, Math.max(1, Math.round(w * scale))),
    Math.min(MAX_SIDE, Math.max(1, Math.round(h * scale))),
  ];
}

/** The picture as SVG in `doc`: the painted layout, cut to the shape's box for a shape. */
function paintPicture(
  session: DocumentSession,
  index: number,
  request: ImageRender,
  doc: Document,
  inline: (face: RegisteredFace) => boolean
): { svg: SVGSVGElement; width: number; height: number } {
  const lay = session.layout(index);
  if (!lay) throw new CommandRefusal('errors.exportFailed', `slide ${index + 1}`);
  let picture: SlideLayout = lay;
  let box = { x: 0, y: 0, w: lay.slide.width, h: lay.slide.height };
  if (request.shapeNodeId !== undefined) {
    const items = branchTo(lay.items, request.shapeNodeId);
    const item = leaf(items, request.shapeNodeId);
    if (!item) throw new CommandRefusal('errors.unsupported', 'shape');
    picture = { ...lay, items };
    box = boundsOf(item);
  }
  const [width, height] = fitSize(box.w / EMU_PER_PX, box.h / EMU_PER_PX, request);
  const svg = paintLayout(session, picture, width, height, inline, undefined, doc);
  if (request.shapeNodeId !== undefined) {
    const background = elementsWithAttribute(svg, 'data-kind').find(
      (el) => el.getAttribute('data-kind') === 'background'
    );
    background?.parentNode?.removeChild(background);
    svg.setAttribute('viewBox', `${box.x} ${box.y} ${box.w} ${box.h}`);
  }
  return { svg, width, height };
}

/**
 * Slide `index`, or one shape of it, as PNG bytes. Rejects with a `CommandRefusal`:
 * `errors.exportFailed` (no DOM, no canvas, or a decode that failed) or `errors.unsupported` (a
 * shape the slide does not draw).
 */
export async function renderImagePng(
  session: DocumentSession,
  index: number,
  request: ImageRender
): Promise<Uint8Array> {
  if (typeof document === 'undefined') throw new CommandRefusal('errors.exportFailed', 'no DOM');
  await Promise.all([document.fonts?.ready, session.picturesReady(index)]);
  const { svg, width, height } = paintPicture(session, index, request, document, () => true);
  const png = await drawPng(svg, width, height);
  return new Uint8Array(await png.arrayBuffer());
}

/**
 * One face that the SVG of a picture names: `family` is the name in its `font-family`
 * attributes and `@font-face` rules, and `name` is the family the presentation asked for.
 *
 * @public
 */
export interface ImageRasterFont {
  /** The name of the face in the SVG: its `font-family` attributes and `@font-face` rules. */
  readonly family: string;
  /** The family the presentation asked for. */
  readonly name: string;
  /** Whether the face is bold. */
  readonly bold: boolean;
  /** Whether the face is italic. */
  readonly italic: boolean;
  /**
   * The font file. Use it for this picture only: never store it, cache it, or write it to disk.
   * A rasterizer with a process-wide font registry registers the faces for one call and removes
   * them before it returns, in a `finally` block, so one document never draws with the fonts of
   * another. `family` is unique for each open document, so faces of two documents never collide.
   */
  readonly bytes: Uint8Array;
}

/**
 * What a rasterizer draws: a standalone SVG document of `width` by `height` pixels from the
 * editor's own paint. Pictures are inline as data URLs. Each face is inline as an `@font-face`
 * rule too, and listed in `fonts` for a rasterizer that does not read CSS fonts. Glyphs carry
 * their x positions from layout, so a rasterizer never breaks a line again.
 *
 * @public
 */
export interface ImageRasterRequest {
  /** The SVG document as text. */
  readonly svg: string;
  /** The width of the PNG to answer, in pixels. */
  readonly width: number;
  /** The height of the PNG to answer, in pixels. */
  readonly height: number;
  /** The faces that the SVG names. */
  readonly fonts: readonly ImageRasterFont[];
}

/**
 * Turns the SVG of a picture into PNG bytes of exactly `width` by `height` pixels.
 *
 * @public
 */
export type ImageRasterizer = (request: ImageRasterRequest) => Promise<Uint8Array> | Uint8Array;

/** The largest PNG a rasterizer may answer. */
const MAX_PNG_BYTES = 256 * 1024 * 1024;

/**
 * @internal Slide `index`, or one shape of it, as PNG bytes from `rasterize`, with no DOM: the
 * same paint as {@link renderImagePng}, into a DOM-free SVG document. Rejects with a
 * `CommandRefusal` (`errors.exportFailed`) when the rasterizer throws or answers anything but a
 * PNG of the asked size.
 */
export async function renderImageWith(
  session: DocumentSession,
  index: number,
  request: ImageRender,
  rasterize: ImageRasterizer
): Promise<Uint8Array> {
  // The SVG and the fonts leave the engine for the host's rasterizer: only faces that may travel
  // go with them, as for the SVG export (`export-image.ts`).
  await session.picturesReady(index);
  const { svg, width, height } = paintPicture(
    session,
    index,
    request,
    createSvgDocument(),
    mayTravel
  );
  const used = new Set(
    elementsWithAttribute(svg, 'font-family').map((el) =>
      (el.getAttribute('font-family') ?? '').split(',')[0]!.trim()
    )
  );
  const fonts: ImageRasterFont[] = session.faces
    .filter((face) => used.has(face.id) && mayTravel(face))
    .map((face) => ({
      family: face.id,
      name: face.family,
      bold: face.bold,
      italic: face.italic,
      bytes: face.bytes.slice(),
    }));
  let png: unknown;
  try {
    png = await rasterize({ svg: serializeSvgElement(svg), width, height, fonts });
  } catch {
    throw new CommandRefusal('errors.exportFailed', 'rasterize');
  }
  if (!(png instanceof Uint8Array) || png.length > MAX_PNG_BYTES)
    throw new CommandRefusal('errors.exportFailed', 'png');
  const isPng = png[0] === 0x89 && png[1] === 0x50 && png[2] === 0x4e && png[3] === 0x47;
  const size = isPng ? imageSize(png) : null;
  if (!size || size.width !== width || size.height !== height)
    throw new CommandRefusal('errors.exportFailed', 'png');
  return png;
}
