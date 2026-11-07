// The pictures of a session as the browser draws them. Most picture parts draw as they are. An
// EMF or WMF part draws as its PNG (`output/metafile-png.ts`), which converts in the background:
// the placeholder paints until the PNG is ready, then `changed` asks for a repaint.

import { metafileFormat } from '@eigenpal/ooxml-core';
import { metafilePng } from '../output/metafile-png.ts';
import { drawableImageMime } from './drawable-image.ts';

/** Picture bytes and the MIME type to draw them with. */
export interface DrawablePicture {
  readonly bytes: Uint8Array;
  readonly mime: string;
}

/** How many metafiles convert at the same time. */
const PARALLEL = 2;

export class MetafilePictures {
  /** The PNG of each converted metafile part, or null when it did not convert. */
  #done = new WeakMap<Uint8Array, Uint8Array | null>();
  /** Each metafile that waits or converts, with what resolves its `convert` promise. */
  #pending = new Map<
    Uint8Array,
    { readonly promise: Promise<void>; readonly resolve: () => void }
  >();
  #queue: Uint8Array[] = [];
  #running = 0;
  #ready = false;
  #disposed = false;

  /**
   * `fonts` gives the font files that metafile text draws with. `changed` hears that PNGs are
   * ready, once for the metafiles that finish together.
   */
  constructor(
    private readonly fonts: () => readonly Uint8Array[],
    private readonly changed: () => void
  ) {}

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
    if (png === undefined && metafileFormat(bytes)) void this.convert(bytes);
    return undefined;
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

  /** Convert every metafile among `parts`. Resolves when each converted or failed. */
  async convertAll(parts: Iterable<Uint8Array | undefined>): Promise<void> {
    const work: Promise<void>[] = [];
    for (const bytes of parts) if (bytes && metafileFormat(bytes)) work.push(this.convert(bytes));
    await Promise.all(work);
  }

  /** Whether a metafile waits or converts. @internal */
  get busy(): boolean {
    return this.#pending.size > 0;
  }

  /** Stop: no conversion starts, every `convert` promise resolves, and `changed` stays quiet. */
  dispose(): void {
    this.#disposed = true;
    this.#queue = [];
    for (const { resolve } of this.#pending.values()) resolve();
    this.#pending.clear();
  }

  #pump(): void {
    while (this.#running < PARALLEL && this.#queue.length > 0) {
      const bytes = this.#queue.shift()!;
      this.#running += 1;
      void metafilePng(bytes, this.fonts()).then((png) => this.#finish(bytes, png));
    }
  }

  #finish(bytes: Uint8Array, png: Uint8Array | null): void {
    this.#running -= 1;
    if (this.#disposed) return;
    this.#done.set(bytes, png);
    const pending = this.#pending.get(bytes);
    this.#pending.delete(bytes);
    if (png) this.#ready = true;
    // One repaint for the metafiles that finish together.
    if (this.#ready && this.#pending.size === 0) {
      this.#ready = false;
      this.changed();
    }
    pending?.resolve();
    this.#pump();
  }
}
