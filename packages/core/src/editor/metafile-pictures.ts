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
/** How often converted pictures repaint at most while others still convert, in ms. */
const REPAINT_INTERVAL_MS = 250;
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
  /** The PNG of each metafile part with the current fonts, or null when it did not convert. */
  #done = new WeakMap<Uint8Array, Uint8Array | null>();
  /** The last PNG of each metafile part that converted: it draws until a new one is ready. */
  #last = new WeakMap<Uint8Array, Uint8Array>();
  /** A number for each PNG, so a key can name the PNGs a slide draws (`stateOf`). */
  #serials = new WeakMap<Uint8Array, number>();
  #serial = 0;
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
  /** Stops the conversions with the current fonts: a font change or `dispose()` aborts it. */
  #abort = new AbortController();
  /** The blob URL of each picture part name, with the bytes it draws. */
  #urls = new Map<string, { readonly drawn: Uint8Array; readonly url: string }>();
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
    if (!metafileFormat(bytes)) return undefined;
    if (!this.#done.has(bytes)) void this.convert(bytes);
    const png = this.#done.get(bytes) ?? this.#last.get(bytes);
    return png ? { bytes: png, mime: 'image/png' } : undefined;
  }

  /**
   * A blob URL for `picture`, the drawable bytes of the part `part`. New bytes for the part (a
   * new PNG, or a replaced picture) revoke the URL of the old ones; `dispose()` revokes the
   * rest. Undefined with no `Blob`.
   */
  url(part: string, picture: DrawablePicture): string | undefined {
    const hit = this.#urls.get(part);
    if (hit?.drawn === picture.bytes) return hit.url;
    if (this.#disposed || typeof URL === 'undefined' || typeof Blob === 'undefined') return;
    if (hit) URL.revokeObjectURL(hit.url);
    const blob = new Blob([picture.bytes.slice().buffer], { type: picture.mime });
    const url = URL.createObjectURL(blob);
    this.#urls.set(part, { drawn: picture.bytes, url });
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
   * Convert `metafiles`. Resolves `true` when each one converted, so all draw. Resolves `false`
   * when one did not convert, and after `waitMs`: a picture that is not ready by then keeps
   * converting, and draws as the placeholder (or its PNG from before a font change) until it is.
   */
  async ready(metafiles: readonly Uint8Array[], waitMs = PICTURES_WAIT_MS): Promise<boolean> {
    let timer: ReturnType<typeof setTimeout> | undefined;
    const late = new Promise<boolean>((resolve) => (timer = setTimeout(resolve, waitMs, false)));
    const all = (async (): Promise<boolean> => {
      // A font change while it waits converts again: wait for the conversions with the new fonts.
      for (;;) {
        const epoch = this.#epoch;
        await Promise.all(metafiles.map((b) => this.convert(b)));
        if (epoch === this.#epoch) return metafiles.every((b) => !!this.#done.get(b));
      }
    })();
    try {
      // A cache disposed meanwhile converted nothing more: its pictures are not ready.
      return (await Promise.race([all, late])) && !this.#disposed;
    } finally {
      clearTimeout(timer);
    }
  }

  /** A key that changes when the PNG of a metafile that slide part `slide` shows changes. */
  stateOf(pkg: PptxPackage, slide: string | undefined): string {
    return (slide ? metafilesOf(pkg, [slide]) : [])
      .map((b) => {
        const png = this.#done.get(b) ?? this.#last.get(b);
        return png ? this.#serials.get(png) : '-';
      })
      .join(',');
  }

  /** The fonts changed: each metafile converts again with them. The last PNGs draw until then. */
  fontsChanged(): void {
    // Conversions with the old fonts stop, and run again with the new ones.
    this.#abort.abort();
    this.#abort = new AbortController();
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
    for (const { url } of this.#urls.values()) URL.revokeObjectURL(url);
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
    if (png) {
      this.#last.set(bytes, png);
      this.#serials.set(png, (this.#serial += 1));
    }
    const pending = this.#pending.get(bytes);
    this.#pending.delete(bytes);
    pending?.resolve();
    this.#pump();
    if (png) this.#schedule();
  }

  /**
   * Repaint once the conversions end, and every `REPAINT_INTERVAL_MS` while they go on, so a
   * deck with many metafiles does not repaint once for each one.
   */
  #schedule(): void {
    const idle = this.#pending.size === 0;
    if (this.#repaint && !idle) return;
    if (this.#repaint) clearTimeout(this.#repaint);
    this.#repaint = setTimeout(() => this.#notify(), idle ? 0 : REPAINT_INTERVAL_MS);
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
