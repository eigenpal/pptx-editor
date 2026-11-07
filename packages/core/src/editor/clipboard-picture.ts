// A picture of the selected shapes for the clipboard: other applications paste copied shapes as
// this picture. It comes from the same paint as the slide (`paintLayout`), cut to the shapes'
// bounds, without the slide background. The PNG states its resolution (`pHYs`), so an
// application that sizes pictures by it pastes the picture at the shapes' size.
//
// Drawing a PNG takes a moment and a copy event cannot wait, so the picture is drawn ahead of
// the copy (`warm`, when Cmd or Ctrl goes down or a menu opens) and kept for that selection.

import type { EditorSelection } from '../contracts/index.ts';
import type { LayoutItem, SlideLayout } from '../layout/index.ts';
import { toBase64 } from '../automation/commands/base64.ts';
import { scopeItems, unionBounds } from '../automation/commands/placement.ts';
import type { DocumentSession } from './document-session.ts';
import { drawPng, paintLayout } from './export-image.ts';

/** Pixels per point of the picture: 144 pixels per inch. */
const PX_PER_PT = 2;
const EMU_PER_PT = 12_700;
/** Room around the shapes for a line's width and a shadow, EMU (4 pt). */
const PAD = 4 * EMU_PER_PT;
const MAX_SIDE = 4096;

/** A drawn picture of shapes. */
export interface ShapesPicture {
  /** The PNG as base64. */
  readonly base64: string;
  readonly width: number;
  readonly height: number;
}

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(bytes: Uint8Array): number {
  let c = 0xffffffff;
  for (const b of bytes) c = CRC_TABLE[(c ^ b) & 0xff]! ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

const PNG_SIGNATURE = [137, 80, 78, 71, 13, 10, 26, 10];

/** Whether `bytes` start as a PNG with its IHDR chunk. */
function isPng(bytes: Uint8Array): boolean {
  return bytes.length > 33 && PNG_SIGNATURE.every((b, i) => bytes[i] === b);
}

/**
 * A PNG with a `pHYs` chunk of `dpi` pixels per inch after its IHDR (a canvas writes none).
 * Bytes that are not a PNG come back unchanged.
 */
export function withPngResolution(bytes: Uint8Array, dpi: number): Uint8Array {
  if (!isPng(bytes)) return bytes;
  const ppm = Math.round(dpi / 0.0254);
  const chunk = new Uint8Array(21);
  const view = new DataView(chunk.buffer);
  view.setUint32(0, 9);
  chunk.set([0x70, 0x48, 0x59, 0x73], 4); // 'pHYs'
  view.setUint32(8, ppm);
  view.setUint32(12, ppm);
  chunk[16] = 1; // the unit is the metre
  view.setUint32(17, crc32(chunk.subarray(4, 17)));
  const at = 8 + 25; // the signature, then IHDR (length, type, 13 bytes of data, CRC)
  const out = new Uint8Array(bytes.length + chunk.length);
  out.set(bytes.subarray(0, at), 0);
  out.set(chunk, at);
  out.set(bytes.subarray(at), at + chunk.length);
  return out;
}

/** A PNG's resolution in pixels per inch from its `pHYs` chunk, or undefined. */
export function pngResolution(bytes: Uint8Array): number | undefined {
  if (!isPng(bytes)) return undefined;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  // The chunks before the image data, at most 64 of them.
  for (let at = 8, n = 0; at + 12 <= bytes.length && n < 64; n += 1) {
    const length = view.getUint32(at);
    const type = String.fromCharCode(...bytes.subarray(at + 4, at + 8));
    if (type === 'IDAT' || type === 'IEND') return undefined;
    if (type === 'pHYs' && length === 9 && at + 17 <= bytes.length) {
      if (bytes[at + 16] !== 1) return undefined;
      const ppm = view.getUint32(at + 8);
      const dpi = ppm * 0.0254;
      return dpi >= 24 && dpi <= 2400 ? dpi : undefined;
    }
    at += 12 + length;
  }
  return undefined;
}

/** The selected top-level shapes, in z-order, or none for another selection. */
function selectedItems(lay: SlideLayout, selection: EditorSelection): LayoutItem[] {
  if (selection.kind !== 'shapes' || selection.group) return [];
  const ids = new Set(selection.nodeIds);
  return scopeItems(lay.items, undefined).filter((i) => ids.has(i.element.nodeId));
}

/** Draw the selected shapes as a PNG. Null when there is nothing to draw. */
export async function drawShapesPicture(
  session: DocumentSession,
  lay: SlideLayout,
  selection: EditorSelection
): Promise<ShapesPicture | null> {
  const items = selectedItems(lay, selection);
  if (!items.length || typeof document === 'undefined') return null;
  const b = unionBounds(items);
  const x = b.x - PAD;
  const y = b.y - PAD;
  const w = b.r - b.x + 2 * PAD;
  const h = b.b - b.y + 2 * PAD;
  if (!(w > 0 && h > 0)) return null;
  const scale = Math.min(PX_PER_PT / EMU_PER_PT, MAX_SIDE / w, MAX_SIDE / h);
  const width = Math.max(1, Math.round(w * scale));
  const height = Math.max(1, Math.round(h * scale));
  const svg = paintLayout(session, { ...lay, items }, width, height, () => true);
  svg.querySelector('[data-kind="background"]')?.remove();
  svg.setAttribute('viewBox', `${x} ${y} ${w} ${h}`);
  const blob = await drawPng(svg, width, height);
  const png = new Uint8Array(await blob.arrayBuffer());
  // The resolution that makes the picture as large as the shapes.
  const dpi = (width / (w / EMU_PER_PT)) * 72;
  return { base64: toBase64(withPngResolution(png, dpi)), width, height };
}

/** HTML holding a picture: what other applications paste copied shapes as. */
export function pictureHtml(picture: ShapesPicture, alt: string): string {
  const text = alt.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
  return (
    `<meta charset="utf-8"><!--StartFragment-->` +
    `<img src="data:image/png;base64,${picture.base64}" width="${picture.width / PX_PER_PT}" ` +
    `height="${picture.height / PX_PER_PT}" alt="${text}"><!--EndFragment-->`
  );
}

/** A piece of a picture's base64 that tells it apart, to find it in pasted HTML. */
export function pictureMark(picture: ShapesPicture): string {
  // The end of the image data, before the IEND chunk: its checksum differs from picture to picture.
  return picture.base64.slice(-96, -32);
}

/** Pictures drawn ahead of a copy, one selection at a time. */
export function createPictureCache() {
  let key = '';
  let pending: Promise<ShapesPicture | null> | null = null;
  let done: ShapesPicture | null = null;
  return {
    /** Start drawing the picture for `key` unless it is drawn or drawing. */
    warm(next: string, draw: () => Promise<ShapesPicture | null>): Promise<ShapesPicture | null> {
      if (next === key && pending) return pending;
      key = next;
      done = null;
      const mine = draw()
        .catch(() => null)
        .then((picture) => {
          if (key === next) done = picture;
          return picture;
        });
      pending = mine;
      return mine;
    },
    /** The picture drawn for `key`, or null when it is not ready. */
    ready(at: string): ShapesPicture | null {
      return at === key ? done : null;
    },
  };
}
