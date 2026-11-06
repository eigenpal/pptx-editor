// Slide thumbnails for the rail.
//
// A container is painted once, then again only when its slide renders differently
// (`renderKey`). A slide that changed in place (typing, a drag, undo, a theme edit) repaints
// when the editor is idle: after a quiet time, and at least once a second while edits go on.
// The painted picture is always the slide as it is when the repaint runs.
//
// A first paint (a new container, or one that changes slide or document) is immediate, unless
// the caller asks to defer it (`defer`): the rail and the slide sorter mount one container per
// slide, and a long deck must not paint every slide before the editor can be used. Then the
// current slide's container paints at once. Of the others, `ThumbnailQueue`
// (thumbnail-queue.ts) paints the first of a task at once and the ones in view before the
// frame; the rest paint when they scroll into view or when idle. A container never shows a
// wrong slide: one that waits is empty.

import type { SlideLayout } from '../layout/index.ts';
import { paintSlide } from '../output/index.ts';
import type { PptxPackage } from '../store/index.ts';
import type { DocumentSession } from './document-session.ts';
import { renderKey } from './slide-dependencies.ts';
import { idleCallback, ThumbnailQueue, type Ask, type IdleDeadline } from './thumbnail-queue.ts';

/** The label of a frame the engine does not draw ("Chart (not rendered in this version)"). */
export function frameLabelOf(
  translate: ((key: string, params?: Record<string, string | number>) => string) | undefined,
  kind: string
): string {
  return translate
    ? translate('canvas.placeholderObject', {
        kind: translate(`canvas.kinds.${kind === 'unknown' ? 'unknown' : kind}`),
      })
    : kind;
}

/** What the painter asks of the editor. */
export interface ThumbnailHost {
  session(): DocumentSession | null;
  /** Unique to the editor instance: ids in two editors' thumbnails never collide. */
  readonly idPrefix: string;
  frameLabel(kind: string): string;
  /** How a new slide on a layout looks, for the layout galleries. */
  layoutPreview(session: DocumentSession, layoutId: string): SlideLayout | undefined;
  /** The current slide, whose thumbnail paints first. */
  currentSlide?(): number;
}

interface Painted {
  readonly session: DocumentSession;
  readonly part: string;
  readonly key: string;
}

/** Quiet time after the last change before a changed slide repaints. */
const QUIET_MS = 150;
/** The longest a changed slide waits while edits go on. */
const MAX_WAIT_MS = 1000;

export class ThumbnailPainter {
  readonly #host: ThumbnailHost;
  #painted = new WeakMap<HTMLElement, Painted>();
  #layouts = new WeakMap<HTMLElement, { pkg: PptxPackage; id: string; fonts: number }>();
  /** Containers waiting for an idle repaint, with the slide index asked for last. */
  #pending = new Map<HTMLElement, number>();
  /** How many times each container was asked for: a disposer acts for the last ask only. */
  #asks = new WeakMap<HTMLElement, number>();
  #timer: ReturnType<typeof setTimeout> | null = null;
  #idle: { cancel(): void } | null = null;
  /** When the oldest waiting repaint was asked for (0: none waits). */
  #since = 0;
  #count = 0;
  /** First paints of containers out of view, which wait for view or idle time. */
  readonly #queue = new ThumbnailQueue({
    paint: (container, ask) => this.#paintFirst(container, ask),
    blank: (container) => {
      this.#painted.delete(container);
      container.replaceChildren();
    },
  });

  constructor(host: ThumbnailHost) {
    this.#host = host;
  }

  /**
   * Paint slide `index` into `container`, now or, with `defer`, when it is in view or idle (see
   * the file comment). Returns a disposer that forgets the container (`release`), unless a later
   * call asked for it again.
   */
  render(index: number, container: HTMLElement, defer = false): () => void {
    const ask = (this.#asks.get(container) ?? 0) + 1;
    this.#asks.set(container, ask);
    this.#render(index, container, defer);
    return () => {
      if (this.#asks.get(container) === ask) this.release(container);
    };
  }

  #render(index: number, container: HTMLElement, defer: boolean): void {
    const session = this.#host.session();
    if (!session) return;
    // A slide still opening waits as a placeholder in a deferred list. A caller that paints at
    // once (no `defer`) gets the slide now: its layout parses it.
    if (defer && session.slidePending(index)) return this.#placeholder(session, index, container);
    const key = renderKey(session.store.package, index, session.fontGeneration);
    const part = session.info.slides[index]?.partName;
    if (!key || !part) return;
    const was = this.#painted.get(container);
    if (was?.session === session && was.key === key) {
      this.#pending.delete(container);
      this.#queue.drop(container);
      return;
    }
    if (!was || was.session !== session || was.part !== part) {
      this.#pending.delete(container);
      const current = !defer || index === this.#host.currentSlide?.();
      if (current || this.#queue.paintNow(container, { index, part }, !!was)) {
        this.#queue.drop(container);
        this.#paint(session, index, container);
      }
      return;
    }
    this.#pending.set(container, index);
    this.#schedule();
  }

  /**
   * Paint a layout's preview into `container`, for the Layout and New Slide galleries. A
   * container repaints when the document or its layout changes.
   */
  renderLayout(layoutId: string, container: HTMLElement): void {
    const session = this.#host.session();
    const pkg = session?.store.package;
    if (!session || !pkg) return;
    const was = this.#layouts.get(container);
    if (was && was.pkg === pkg && was.id === layoutId && was.fonts === session.fontGeneration)
      return;
    const lay = this.#host.layoutPreview(session, layoutId);
    if (!lay) return container.replaceChildren();
    this.#count += 1;
    const svg = paintSlide(container.ownerDocument, lay, {
      images: session.imageUrl,
      imageSize: session.imageSize,
      idPrefix: `${this.#host.idPrefix}l${this.#count}`,
      hitTargets: false,
      frameLabel: (kind) => this.#host.frameLabel(kind),
    });
    svg.setAttribute('width', '100%');
    svg.setAttribute('height', '100%');
    container.replaceChildren(svg);
    this.#layouts.set(container, { pkg, id: layoutId, fonts: session.fontGeneration });
  }

  /** Paint every waiting thumbnail now. */
  flush(): void {
    this.#cancel();
    this.#queue.flush();
    this.#drain(undefined, true);
  }

  /**
   * Forget a container that leaves for good (an unmounted rail item): it waits for no paint
   * and no observer watches it.
   */
  release(container: HTMLElement): void {
    this.#pending.delete(container);
    this.#queue.drop(container);
  }

  /** Whether any first paint waits, or any observer or timer is left (for tests). @internal */
  get waiting(): boolean {
    return this.#queue.busy || this.#pending.size > 0;
  }

  /** Drop waiting repaints (the editor is destroyed or loads another document). */
  clear(): void {
    this.#cancel();
    this.#pending.clear();
    this.#queue.clear();
  }

  /**
   * A slide still opening: the container stays empty (the rail marks it `data-pending` from the
   * slide's summary and styles it as a placeholder), and in view, the slide parses before the
   * others. Its summary changes once it is ready, and the rail asks again. Nothing is written
   * to a container that is already empty: the rail mounts many at once, and a write between two
   * of the queue's style reads would make each read lay the page out again.
   */
  #placeholder(session: DocumentSession, index: number, container: HTMLElement): void {
    this.#pending.delete(container);
    if (this.#painted.delete(container) || container.firstChild) container.replaceChildren();
    const part = session.info.slides[index]?.partName;
    if (!part) return;
    // In view, it parses first; out of view, when it scrolls into view (`#paintFirst`).
    if (!this.#queue.paintNow(container, { index, part }, false)) return;
    this.#queue.drop(container);
    session.opening?.want(part);
  }

  /**
   * A first paint that waited: the slide as it is now, unless something painted it since. The
   * slide is the one asked for: when slides moved since, it is found by its part, and a slide
   * that is gone leaves the container empty. A slide still opening parses first instead.
   */
  #paintFirst(container: HTMLElement, ask: Ask): void {
    const session = this.#host.session();
    if (!session) return;
    const slides = session.info.slides;
    const index =
      slides[ask.index]?.partName === ask.part
        ? ask.index
        : slides.findIndex((s) => s.partName === ask.part);
    if (index < 0) {
      this.#painted.delete(container);
      container.replaceChildren();
      return;
    }
    if (session.slidePending(index)) return void session.opening?.want(ask.part);
    const key = renderKey(session.store.package, index, session.fontGeneration);
    const was = this.#painted.get(container);
    if (was?.session === session && was.key === key) return;
    this.#paint(session, index, container);
  }

  #paint(session: DocumentSession, index: number, container: HTMLElement): void {
    const key = renderKey(session.store.package, index, session.fontGeneration);
    const part = session.info.slides[index]?.partName;
    const lay = session.layout(index);
    if (!key || !part || !lay) return;
    this.#count += 1;
    const svg = paintSlide(container.ownerDocument, lay, {
      images: session.imageUrl,
      imageSize: session.imageSize,
      idPrefix: `${this.#host.idPrefix}t${index}n${this.#count}`,
      hitTargets: false,
      frameLabel: (kind) => this.#host.frameLabel(kind),
    });
    svg.setAttribute('width', '100%');
    svg.setAttribute('height', '100%');
    container.replaceChildren(svg);
    this.#painted.set(container, { session, part, key });
  }

  #schedule(): void {
    if (this.#idle) return;
    const now = Date.now();
    if (this.#since === 0) this.#since = now;
    if (this.#timer !== null) clearTimeout(this.#timer);
    const wait = Math.min(QUIET_MS, Math.max(0, this.#since + MAX_WAIT_MS - now));
    this.#timer = setTimeout(() => {
      this.#timer = null;
      this.#whenIdle();
    }, wait);
  }

  #whenIdle(): void {
    this.#idle = idleCallback((d) => this.#drain(d, false));
  }

  /** Repaint waiting thumbnails while idle time lasts; continue in the next idle period. */
  #drain(deadline: IdleDeadline, all: boolean): void {
    this.#idle = null;
    const session = this.#host.session();
    let painted = 0;
    for (const [container, index] of [...this.#pending]) {
      if (!all && painted > 0 && deadline && deadline.timeRemaining() < 2) break;
      this.#pending.delete(container);
      if (!session) continue;
      const was = this.#painted.get(container);
      const key = renderKey(session.store.package, index, session.fontGeneration);
      if (was?.session === session && was.key === key) continue;
      this.#paint(session, index, container);
      painted += 1;
    }
    if (this.#pending.size > 0) this.#whenIdle();
    else this.#since = 0;
  }

  #cancel(): void {
    if (this.#timer !== null) clearTimeout(this.#timer);
    this.#timer = null;
    this.#idle?.cancel();
    this.#idle = null;
    this.#since = 0;
  }
}
