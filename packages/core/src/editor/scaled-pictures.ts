// Pictures that draw smaller on screen than they are stored. A browser decodes each picture of a
// slide at its stored size, and some engines decode an SVG `<image>` on the main thread while they
// paint: a slide of 200 pictures of 16 million pixels each holds the page for tens of seconds.
// When the raster pictures that one part (a slide, layout, master, or drawing) relates to hold
// more decoded pixels than `PART_PIXEL_BUDGET`, each one draws scaled down to its share of the
// budget, decoded off the main thread (`createImageBitmap`). The placeholder paints until a scaled
// picture is ready, then `changed` asks for a repaint. A part within the budget draws its
// pictures as they are stored. Export, copy, and save always use the stored bytes.

import { DEFAULT_IMAGE_RESOURCE_LIMITS } from '@eigenpal/ooxml-core';
import { binaryPartOf, imageSize, REL, relationshipsOf, type PptxPackage } from '../store/index.ts';
import { drawableImageMime } from './drawable-image.ts';
import type { DrawablePicture } from './metafile-pictures.ts';

/** The decoded pixels that the raster pictures of one part draw with: the decode limit at 4 bytes a pixel. */
export const PART_PIXEL_BUDGET = DEFAULT_IMAGE_RESOURCE_LIMITS.maxDecodedBytes / 4;
/** How many pictures scale at the same time. */
const PARALLEL = 2;
/** How often scaled pictures repaint at most while others still scale, in ms. */
const REPAINT_INTERVAL_MS = 250;
/**
 * How far, as a ratio, a part's factor can move before its pictures scale again: inserting or
 * deleting one picture of many does not scale every other one again. The pictures of a part
 * hold at most 1.25² times the budget.
 */
const STEADY = 1.25;

const scaleByPackage = new WeakMap<PptxPackage, Map<string, number>>();

/**
 * The factor, at most 1, that each raster picture of `part` scales its width and height by on
 * screen: 1 when the distinct pictures that `part` relates to hold at most `PART_PIXEL_BUDGET`
 * decoded pixels, else the factor that brings them to it.
 */
export function partPictureScale(pkg: PptxPackage, part: string): number {
  let byPart = scaleByPackage.get(pkg);
  if (!byPart) scaleByPackage.set(pkg, (byPart = new Map()));
  const hit = byPart.get(part);
  if (hit !== undefined) return hit;
  const seen = new Set<string>();
  let pixels = 0;
  for (const rel of relationshipsOf(pkg, part)) {
    if (rel.type !== REL.image || rel.target === undefined || seen.has(rel.target)) continue;
    seen.add(rel.target);
    const bytes = binaryPartOf(pkg, rel.target);
    // An SVG picture has no stored pixels; a picture the browser does not draw decodes nothing.
    const size = bytes && drawableImageMime(bytes, undefined) ? imageSize(bytes) : null;
    if (size) pixels += size.width * size.height;
  }
  const scale = pixels > PART_PIXEL_BUDGET ? Math.sqrt(PART_PIXEL_BUDGET / pixels) : 1;
  byPart.set(part, scale);
  return scale;
}

/**
 * `bytes` with a JPEG's EXIF orientation set to 1, in a copy, so they decode to their stored
 * pixels: the editor draws pictures unturned, but `createImageBitmap` always turns them. Other
 * bytes come back as they are.
 */
export function withStoredOrientation(bytes: Uint8Array): Uint8Array {
  if (bytes[0] !== 0xff || bytes[1] !== 0xd8) return bytes;
  let at = 2;
  while (at + 4 <= bytes.length && bytes[at] === 0xff) {
    const marker = bytes[at + 1]!;
    // Fill bytes before a marker.
    if (marker === 0xff) {
      at += 1;
      continue;
    }
    // The image data starts: no metadata after it.
    if (marker === 0xda || marker === 0xd9) break;
    const length = (bytes[at + 2]! << 8) | bytes[at + 3]!;
    if (length < 2) break;
    const end = Math.min(bytes.length, at + 2 + length);
    const exif = [0x45, 0x78, 0x69, 0x66, 0, 0].every((b, i) => bytes[at + 4 + i] === b);
    if (marker === 0xe1 && exif) {
      const turned = orientationAt(bytes, at + 10, end);
      if (turned === undefined) return bytes;
      const copy = bytes.slice();
      const little = bytes[at + 10] === 0x49;
      copy[turned] = little ? 1 : 0;
      copy[turned + 1] = little ? 0 : 1;
      return copy;
    }
    at += 2 + length;
  }
  return bytes;
}

/** Where the EXIF orientation value of the TIFF block at `tiff` is, when it is not 1. */
function orientationAt(bytes: Uint8Array, tiff: number, end: number): number | undefined {
  const little = bytes[tiff] === 0x49;
  const u16 = (i: number): number =>
    i + 2 > end ? -1 : little ? bytes[i]! | (bytes[i + 1]! << 8) : (bytes[i]! << 8) | bytes[i + 1]!;
  const u32 = (i: number): number =>
    i + 4 > end ? -1 : little ? u16(i) + u16(i + 2) * 65_536 : u16(i) * 65_536 + u16(i + 2);
  const offset = u32(tiff + 4);
  if (offset < 0) return undefined;
  const ifd = tiff + offset;
  const count = u16(ifd);
  for (let i = 0; i < count; i += 1) {
    const entry = ifd + 2 + i * 12;
    if (entry + 12 > end) return undefined;
    if (u16(entry) !== 0x0112) continue;
    return u16(entry + 8) === 1 ? undefined : entry + 8;
  }
  return undefined;
}

/** Whether this runtime can scale a picture off the main thread. */
const canScale = (): boolean =>
  typeof createImageBitmap === 'function' && typeof OffscreenCanvas === 'function';

/**
 * `bytes` decoded at `width` by `height` pixels off the main thread and encoded again: JPEG for a
 * JPEG, else PNG. Null when the browser cannot decode them.
 */
export async function scalePicture(
  bytes: Uint8Array,
  mime: string,
  width: number,
  height: number
): Promise<Uint8Array | null> {
  const blob = new Blob([withStoredOrientation(bytes).slice().buffer], { type: mime });
  const bitmap = await createImageBitmap(blob, {
    resizeWidth: width,
    resizeHeight: height,
    resizeQuality: 'high',
  });
  try {
    const canvas = new OffscreenCanvas(width, height);
    const ctx = canvas.getContext('2d');
    if (!ctx) return null;
    ctx.drawImage(bitmap, 0, 0);
    const jpeg = mime === 'image/jpeg';
    const out = await canvas.convertToBlob(
      jpeg ? { type: 'image/jpeg', quality: 0.92 } : { type: 'image/png' }
    );
    return new Uint8Array(await out.arrayBuffer());
  } finally {
    bitmap.close();
  }
}

/** A picture as it draws on screen, with the key of its size: empty for the stored size. */
export interface ShownPicture {
  readonly picture: DrawablePicture;
  readonly key: string;
}

interface Job {
  readonly picture: DrawablePicture;
  readonly width: number;
  readonly height: number;
  readonly key: string;
  /** The parts that show the picture at this size. */
  readonly sources: Set<string>;
}

export class ScaledPictures {
  /** The scaled bytes of each picture, by size key; null when it did not scale. */
  #done = new WeakMap<Uint8Array, Map<string, Uint8Array | null>>();
  /** The jobs of each picture that waits or scales, by size key. */
  #pending = new WeakMap<Uint8Array, Map<string, Job>>();
  #queue: Job[] = [];
  #running = 0;
  /** For each part that shows pictures, how many of its scaled pictures are ready. */
  #serials = new Map<string, number>();
  /** The size key of the last scaled bytes of each picture: they draw while another size scales. */
  #latest = new WeakMap<Uint8Array, string>();
  /**
   * For each picture, the size each part that shows it asks for (`want`) and the size it drew
   * last (`drawn`). A size that no part wants or draws is dropped (`#finish`).
   */
  #uses = new WeakMap<Uint8Array, Map<string, { want: string; drawn?: string }>>();
  /** The factor that each part over the budget draws its pictures with (`steady`). */
  #factors = new Map<string, number>();
  #repaint: ReturnType<typeof setTimeout> | null = null;
  #disposed = false;
  readonly #changed: () => void;
  readonly #scale: typeof scalePicture;
  readonly #can: () => boolean;

  /** `changed` hears that scaled pictures are ready. `scale` and `can` are for tests. */
  constructor(changed: () => void, scale: typeof scalePicture = scalePicture, can = canScale) {
    this.#changed = changed;
    this.#scale = scale;
    this.#can = can;
  }

  /** A key that changes each time a picture that part `source` shows finishes scaling. */
  serialOf(source: string): string {
    const n = this.#serials.get(source);
    return n ? `${source}:${n}` : '';
  }

  /** Whether a picture waits or scales. @internal */
  get busy(): boolean {
    return this.#running > 0 || this.#queue.length > 0;
  }

  /**
   * How `picture` draws when part `source` shows it and scales its pictures by `factor`
   * (`partPictureScale`): as it is for a factor of 1, an SVG picture, or a runtime that cannot
   * scale; scaled once that is ready. While it scales to a new size, it draws at the last size
   * that is ready. Undefined while it scales for the first time or when it did not scale, so the
   * placeholder paints.
   */
  shown(picture: DrawablePicture, factor: number, source: string): ShownPicture | undefined {
    const stored = { picture, key: '' };
    if (factor >= 1 || picture.mime === 'image/svg+xml' || !this.#can()) return stored;
    const size = imageSize(picture.bytes);
    if (!size) return stored;
    const width = Math.max(1, Math.floor(size.width * factor));
    const height = Math.max(1, Math.floor(size.height * factor));
    const key = `${width}x${height}`;
    let uses = this.#uses.get(picture.bytes);
    if (!uses) this.#uses.set(picture.bytes, (uses = new Map()));
    const use = uses.get(source) ?? { want: key };
    use.want = key;
    uses.set(source, use);
    const sizes = this.#done.get(picture.bytes);
    const done = sizes?.get(key);
    if (done) {
      use.drawn = key;
      return { picture: { bytes: done, mime: picture.mime }, key };
    }
    if (done === null || this.#disposed) return undefined;
    // Meanwhile: the size this part drew last, else the last size of the picture that is ready.
    const fallback = [use.drawn, this.#latest.get(picture.bytes)].find((k) => k && sizes?.get(k));
    const before = fallback
      ? { picture: { bytes: sizes!.get(fallback)!, mime: picture.mime }, key: fallback }
      : undefined;
    if (before) use.drawn = before.key;
    let pending = this.#pending.get(picture.bytes);
    if (!pending) this.#pending.set(picture.bytes, (pending = new Map()));
    const job = pending.get(key);
    if (job) job.sources.add(source);
    else {
      const next = { picture, width, height, key, sources: new Set([source]) };
      pending.set(key, next);
      this.#queue.push(next);
      this.#pump();
    }
    return before;
  }

  /**
   * The factor that part `source` draws its pictures with, for its factor `raw` now
   * (`partPictureScale`): the factor from before while `raw` is within `STEADY` of it, so an
   * edit of the part's pictures does not scale all of them again.
   */
  steady(source: string, raw: number): number {
    if (raw >= 1) {
      this.#factors.delete(source);
      return 1;
    }
    const last = this.#factors.get(source);
    if (last !== undefined && raw * STEADY >= last && raw <= last * STEADY) return last;
    this.#factors.set(source, raw);
    return raw;
  }

  /** Stop: no picture scales any more, and `changed` stays quiet. */
  dispose(): void {
    this.#disposed = true;
    this.#queue = [];
    if (this.#repaint) clearTimeout(this.#repaint);
    this.#repaint = null;
  }

  #pump(): void {
    while (this.#running < PARALLEL && this.#queue.length > 0 && !this.#disposed) {
      const job = this.#queue.shift()!;
      this.#running += 1;
      void this.#scale(job.picture.bytes, job.picture.mime, job.width, job.height)
        .catch(() => null)
        .then((bytes) => this.#finish(job, bytes));
    }
  }

  #finish(job: Job, bytes: Uint8Array | null): void {
    this.#running -= 1;
    if (this.#disposed) return;
    let done = this.#done.get(job.picture.bytes);
    if (!done) this.#done.set(job.picture.bytes, (done = new Map()));
    done.set(job.key, bytes);
    if (bytes) {
      // Keep the sizes that a part wants or drew last: no part draws the others any more.
      const kept = new Set([job.key]);
      for (const use of this.#uses.get(job.picture.bytes)?.values() ?? []) {
        kept.add(use.want);
        if (use.drawn) kept.add(use.drawn);
      }
      for (const key of done.keys()) if (!kept.has(key)) done.delete(key);
      this.#latest.set(job.picture.bytes, job.key);
    }
    this.#pending.get(job.picture.bytes)?.delete(job.key);
    for (const s of job.sources) this.#serials.set(s, (this.#serials.get(s) ?? 0) + 1);
    this.#pump();
    this.#schedule();
  }

  /** Repaint once the pictures are scaled, and every `REPAINT_INTERVAL_MS` while more scale. */
  #schedule(): void {
    const idle = !this.busy;
    if (this.#repaint && !idle) return;
    if (this.#repaint) clearTimeout(this.#repaint);
    this.#repaint = setTimeout(
      () => {
        this.#repaint = null;
        if (this.#disposed) return;
        try {
          this.#changed();
        } catch (error) {
          (globalThis.reportError ?? console.error)(error);
        }
      },
      idle ? 0 : REPAINT_INTERVAL_MS
    );
  }
}
