// The pictures of a session as the browser draws them. Most picture parts draw as they are. An
// EMF or WMF part draws as its PNG (`output/metafile-png.ts`), which converts in the background:
// the placeholder paints until the PNG is ready, then `changed` asks for a repaint.

import { metafileFormat } from '@eigenpal/ooxml-core';
import { metafilePng } from '../output/metafile-png.ts';
import { binaryPartOf, REL, relationshipsOf, type PptxPackage } from '../store/index.ts';
import { drawableImageMime } from './drawable-image.ts';

/** Picture bytes and the MIME type to draw them with. */
export interface DrawablePicture {
  readonly bytes: Uint8Array;
  readonly mime: string;
}

/** How many metafiles convert at the same time. */
const PARALLEL = 2;
/** How long a wait for converted pictures lasts at most, in ms (`ready`). */
export const PICTURES_WAIT_MS = 30_000;
/** How long the repaint after a conversion waits for others to finish with it, in ms. */
const REPAINT_DELAY_MS = 50;
/** Relationships that do not lead to what a slide shows. */
const OTHER_CONTENT = new Set<string>([
  REL.slide,
  REL.notesSlide,
  REL.notesMaster,
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster',
]);

/**
 * The metafile parts that `slides` can show: the parts that the slides, their layouts, and their
 * masters relate to, also through a drawing part (a diagram). Another slide, the notes, and the
 * other layouts of a master are not followed. Each part is read once.
 */
export function metafilesOf(pkg: PptxPackage, slides: readonly string[]): Uint8Array[] {
  const seen = new Set<string>(slides);
  const queue = [...slides];
  const out: Uint8Array[] = [];
  for (let i = 0; i < queue.length; i += 1) {
    const source = queue[i]!;
    const master = relationshipsOf(pkg, source).some((r) => r.type === REL.theme);
    for (const rel of relationshipsOf(pkg, source)) {
      const target = rel.target;
      if (target === undefined || seen.has(target) || OTHER_CONTENT.has(rel.type)) continue;
      // A master names every layout it has; only the layout of a slide shows.
      if (master && rel.type === REL.slideLayout) continue;
      seen.add(target);
      const bytes = binaryPartOf(pkg, target);
      if (!bytes) queue.push(target);
      else if (metafileFormat(bytes)) out.push(bytes);
    }
  }
  return out;
}

export class MetafilePictures {
  /** The PNG of each converted metafile part, or null when it did not convert. */
  #done = new WeakMap<Uint8Array, Uint8Array | null>();
  /** The PNGs from before the fonts changed: each draws until its new PNG is ready. */
  #stale = new WeakMap<Uint8Array, Uint8Array | null>();
  /** Each metafile that waits or converts, with what resolves its `convert` promise. */
  #pending = new Map<
    Uint8Array,
    { readonly promise: Promise<void>; readonly resolve: () => void }
  >();
  #queue: Uint8Array[] = [];
  #running = 0;
  /** Bumped when the fonts change: a conversion with the old fonts runs again. */
  #epoch = 0;
  #fontFiles: readonly Uint8Array[] | null = null;
  #repaint: ReturnType<typeof setTimeout> | null = null;
  #abort = new AbortController();
  /** Blob URLs by the bytes they draw. */
  #urls = new Map<Uint8Array, string>();
  #disposed = false;
  readonly #fonts: () => readonly Uint8Array[];
  readonly #changed: () => void;

  /**
   * `fonts` gives the font files that metafile text draws with. `changed` hears that PNGs are
   * ready, once for the metafiles that finish close together.
   */
  constructor(fonts: () => readonly Uint8Array[], changed: () => void) {
    this.#fonts = fonts;
    this.#changed = changed;
  }

  /**
   * The bytes to draw a picture part with, or undefined while it cannot draw: a damaged or
   * oversized picture (`drawable-image.ts`), or a metafile that converts or did not convert.
   * `named` is the type the part name gives. A metafile starts to convert here.
   */
  drawable(bytes: Uint8Array, named: string | undefined): DrawablePicture | undefined {
    const mime = drawableImageMime(bytes, named);
    if (mime) return { bytes, mime };
    const png = this.#done.get(bytes);
    if (png) return { bytes: png, mime: 'image/png' };
    if (png === null || !metafileFormat(bytes)) return undefined;
    void this.convert(bytes);
    const old = this.#stale.get(bytes);
    return old ? { bytes: old, mime: 'image/png' } : undefined;
  }

  /** A blob URL for `picture`, made once and revoked by `dispose()`. Undefined with no `Blob`. */
  url(picture: DrawablePicture): string | undefined {
    const hit = this.#urls.get(picture.bytes);
    if (hit) return hit;
    if (typeof URL === 'undefined' || typeof Blob === 'undefined') return undefined;
    const blob = new Blob([picture.bytes.slice().buffer], { type: picture.mime });
    const url = URL.createObjectURL(blob);
    this.#urls.set(picture.bytes, url);
    return url;
  }

  /** Convert the metafile `bytes` once. Resolves when its PNG is ready or it did not convert. */
  convert(bytes: Uint8Array): Promise<void> {
    if (this.#done.has(bytes) || this.#disposed) return Promise.resolve();
    const pending = this.#pending.get(bytes);
    if (pending) return pending.promise;
    let resolve = (): void => {};
    const promise = new Promise<void>((r) => (resolve = r));
    this.#pending.set(bytes, { promise, resolve });
    this.#queue.push(bytes);
    this.#pump();
    return promise;
  }

  /**
   * Convert `metafiles`. Resolves when each converted or failed, or after `waitMs`: a picture
   * that is not ready by then keeps converting, and draws as the placeholder until it is.
   */
  async ready(metafiles: readonly Uint8Array[], waitMs = PICTURES_WAIT_MS): Promise<void> {
    let timer: ReturnType<typeof setTimeout> | undefined;
    const late = new Promise<void>((resolve) => (timer = setTimeout(resolve, waitMs)));
    try {
      await Promise.race([Promise.all(metafiles.map((b) => this.convert(b))), late]);
    } finally {
      clearTimeout(timer);
    }
  }

  /** The fonts changed: each metafile converts again with them. The old PNGs draw until then. */
  fontsChanged(): void {
    this.#stale = this.#done;
    this.#done = new WeakMap();
    this.#fontFiles = null;
    this.#epoch += 1;
  }

  /** Whether a metafile waits or converts. @internal */
  get busy(): boolean {
    return this.#pending.size > 0;
  }

  /**
   * Stop: running conversions abort, every `convert` promise resolves, `changed` stays quiet,
   * and the blob URLs are revoked.
   */
  dispose(): void {
    this.#disposed = true;
    for (const url of this.#urls.values()) URL.revokeObjectURL(url);
    this.#urls.clear();
    this.#abort.abort();
    if (this.#repaint) clearTimeout(this.#repaint);
    this.#queue = [];
    for (const { resolve } of this.#pending.values()) resolve();
    this.#pending.clear();
  }

  #pump(): void {
    while (this.#running < PARALLEL && this.#queue.length > 0 && !this.#disposed) {
      const bytes = this.#queue.shift()!;
      const epoch = this.#epoch;
      // One array for each font set, so the converter reuses the fonts it has read.
      this.#fontFiles ??= this.#fonts();
      this.#running += 1;
      void metafilePng(bytes, this.#fontFiles, this.#abort.signal)
        .catch(() => null)
        .then((png) => this.#finish(bytes, png, epoch));
    }
  }

  #finish(bytes: Uint8Array, png: Uint8Array | null, epoch: number): void {
    this.#running -= 1;
    if (this.#disposed) return;
    if (epoch !== this.#epoch) {
      // The fonts changed while it converted: convert again with the new fonts.
      this.#queue.push(bytes);
      this.#pump();
      return;
    }
    this.#done.set(bytes, png);
    const pending = this.#pending.get(bytes);
    this.#pending.delete(bytes);
    pending?.resolve();
    this.#pump();
    if (png && !this.#repaint) this.#repaint = setTimeout(() => this.#notify(), REPAINT_DELAY_MS);
  }

  /** Tell the session to repaint. A listener that throws does not stop the conversions. */
  #notify(): void {
    this.#repaint = null;
    if (this.#disposed) return;
    try {
      this.#changed();
    } catch (error) {
      // Report it as an uncaught error, but keep the conversions going.
      (globalThis.reportError ?? console.error)(error);
    }
  }
}
