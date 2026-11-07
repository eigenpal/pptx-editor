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
// current slide's container paints at once, unless the slide is large: it paints when idle.
// Of the others, `ThumbnailQueue`
// (thumbnail-queue.ts) paints the first of a task at once and the ones in view before the
// frame; the rest paint when they scroll into view or when idle. A container never shows a
// wrong slide: one that waits is empty.
//
// At most `MAX_PAINTED` thumbnails stay painted. Each painted thumbnail holds its SVG, and the
// browser holds the layout and the shaped text of it: in a deck of thousands of slides, a rail
// scrolled end to end would keep them all. Past the limit, the thumbnails painted longest ago
// that are out of view are emptied, and paint again when they come into view.

import type { SlideLayout } from '../layout/index.ts';
import { paintSlide, SlidePainter, type PaintOptions } from '../output/index.ts';
import type { PptxPackage } from '../store/index.ts';
import type { DocumentSession } from './document-session.ts';
import { renderKey } from './slide-dependencies.ts';
import { idleCallback, ThumbnailQueue, type Ask, type IdleDeadline } from './thumbnail-queue.ts';

/** The render key of slide `index`: its parts, its faces, and its converted pictures. */
const keyOf = (session: DocumentSession, index: number): string | undefined =>
  renderKey(session.store.package, index, session.paintKey(session.info.slides[index]?.partName));

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

/** The number of items from which a current slide's first thumbnail waits for idle time. */
const LARGE_SLIDE_ITEMS = 2_000;

/**
 * The new items an idle repaint paints into one thumbnail. A large slide fills its thumbnail
 * over several idle periods, so no idle repaint blocks input for long.
 */
const IDLE_PAINT_BUDGET = 1_000;

/** The painted thumbnails kept (rail and slide sorter): several screens of the slide sorter. */
export const MAX_PAINTED = 200;

/** Quiet time after the last change before a changed slide repaints. */
const QUIET_MS = 150;
/** The longest a changed slide waits while edits go on. */
const MAX_WAIT_MS = 1000;

export class ThumbnailPainter {
  readonly #host: ThumbnailHost;
  #painted = new WeakMap<HTMLElement, Painted>();
  #layouts = new WeakMap<HTMLElement, { pkg: PptxPackage; id: string; key: string }>();
  /** The layout of each gallery container, to repaint it when its pictures convert. */
  #layoutBoxes = new Map<HTMLElement, string>();
  /** The size of `#layoutBoxes` from which the next call drops the boxes no longer shown. */
  #pruneAt = 64;
  /** Containers waiting for an idle repaint, with the slide index asked for last. */
  #pending = new Map<HTMLElement, number>();
  /** How many times each container was asked for: a disposer acts for the last ask only. */
  #asks = new WeakMap<HTMLElement, number>();
  #timer: ReturnType<typeof setTimeout> | null = null;
  #idle: { cancel(): void } | null = null;
  /** When the oldest waiting repaint was asked for (0: none waits). */
  #since = 0;
  #count = 0;
  /**
   * The painter of each container. It keeps the nodes of the items it painted, so a slide that
   * changed in place repaints the items that changed, not the whole slide.
   */
  #painters = new WeakMap<HTMLElement, SlidePainter>();
  /** The painted containers on the page, painted longest ago first, with their slide. */
  #shown = new Map<HTMLElement, Ask>();
  #evicting = false;
  readonly #frameLabel = (kind: string): string => this.#host.frameLabel(kind);
  /** First paints of containers out of view, which wait for view or idle time. */
  readonly #queue = new ThumbnailQueue({
    paint: (container, ask) => this.#paintFirst(container, ask),
    blank: (container) => {
      this.#painted.delete(container);
      container.replaceChildren();
    },
    full: () => this.#shown.size >= MAX_PAINTED,
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
    const key = keyOf(session, index);
    const part = session.info.slides[index]?.partName;
    if (!key || !part) return;
    const was = this.#painted.get(container);
    if (was?.session === session && was.key === key) {
      this.#pending.delete(container);
      this.#queue.drop(container);
      // Asked again after a release (the rail asks again when a slide's summary changes): it
      // still shows its slide, and counts as painted.
      if (!this.#shown.has(container)) this.#keep(container, { index, part });
      return;
    }
    if (!was || was.session !== session || was.part !== part) {
      this.#pending.delete(container);
      const current = !defer || index === this.#host.currentSlide?.();
      // A large current slide paints its thumbnail when idle, not in the frame that paints the
      // slide itself. Until then, the container is empty.
      if (defer && current && (session.layout(index)?.items.length ?? 0) > LARGE_SLIDE_ITEMS) {
        this.#queue.drop(container);
        if (this.#painted.delete(container) || container.firstChild) container.replaceChildren();
        this.#pending.set(container, index);
        this.#schedule();
        return;
      }
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
    // Drop detached boxes once the map doubles: a scan on each call would cost the size of the
    // gallery each time, so painting a gallery of many layouts would grow with its square.
    if (this.#layoutBoxes.size >= this.#pruneAt) {
      for (const box of this.#layoutBoxes.keys())
        if (!box.isConnected) this.#layoutBoxes.delete(box);
      this.#pruneAt = Math.max(64, this.#layoutBoxes.size * 2);
    }
    this.#layoutBoxes.set(container, layoutId);
    const was = this.#layouts.get(container);
    const key = session.paintKey(layoutId);
    if (was && was.pkg === pkg && was.id === layoutId && was.key === key) return;
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
    this.#layouts.set(container, { pkg, id: layoutId, key });
  }

  /**
   * EMF or WMF pictures converted: repaint the layout previews that show them. A slide
   * thumbnail repaints when its summary changes; a layout no slide uses has no summary.
   */
  picturesChanged(): void {
    for (const [container, id] of [...this.#layoutBoxes]) {
      if (container.isConnected) this.renderLayout(id, container);
      else this.#layoutBoxes.delete(container);
    }
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
    this.#shown.delete(container);
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
    this.#layoutBoxes.clear();
    this.#shown.clear();
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
    const key = keyOf(session, index);
    const was = this.#painted.get(container);
    if (was?.session === session && was.key === key) return;
    this.#paint(session, index, container);
  }

  /**
   * Paint slide `index` into `container`: at most `budget` new items. A paint that leaves items
   * out marks the container as not painted and returns false: the idle repaints go on with it.
   */
  #paint(
    session: DocumentSession,
    index: number,
    container: HTMLElement,
    budget = Infinity
  ): boolean {
    const key = keyOf(session, index);
    const part = session.info.slides[index]?.partName;
    const lay = session.layout(index);
    if (!key || !part || !lay) return true;
    const opts: Omit<PaintOptions, 'idPrefix'> = {
      images: session.imageUrl,
      imageSize: session.imageSize,
      hitTargets: false,
      frameLabel: this.#frameLabel,
    };
    let painter = this.#painters.get(container);
    if (!painter?.paints({ ...opts, idPrefix: painter.idPrefix })) {
      this.#count += 1;
      painter = new SlidePainter(container.ownerDocument, {
        ...opts,
        idPrefix: `${this.#host.idPrefix}t${this.#count}`,
      });
      this.#painters.set(container, painter);
    }
    // Converted pictures, or new faces: every item paints again.
    painter.invalidate(session.paintKey(part));
    const svg = painter.paint(lay, budget);
    if (svg.getAttribute('width') !== '100%') {
      svg.setAttribute('width', '100%');
      svg.setAttribute('height', '100%');
    }
    if (container.firstChild !== svg || container.childNodes.length !== 1)
      container.replaceChildren(svg);
    // Items left out: the slide is not painted yet, and its next repaint goes on with it.
    this.#painted.set(container, { session, part, key: painter.pending ? '' : key });
    this.#keep(container, { index, part });
    return !painter.pending;
  }

  /** Note a painted container as the newest; past `MAX_PAINTED`, empty the oldest out of view. */
  #keep(container: HTMLElement, ask: Ask): void {
    // A container may paint before the host places it (a rail that mounts its items in
    // batches): it counts now, and the end of the task forgets it if it is still off the page.
    this.#shown.delete(container);
    this.#shown.set(container, ask);
    if (this.#shown.size <= MAX_PAINTED || this.#evicting) return;
    this.#evicting = true;
    // At the end of the task, after this task's paints: all layout reads, then the writes.
    queueMicrotask(() => {
      this.#evicting = false;
      this.#evict();
    });
  }

  #evict(): void {
    let excess = this.#shown.size - MAX_PAINTED;
    if (excess <= 0) return;
    const out: [HTMLElement, Ask][] = [];
    for (const [container, ask] of this.#shown) {
      if (out.length >= excess) break;
      if (!container.isConnected) {
        this.#shown.delete(container);
        excess -= 1;
      } else if (!this.#queue.inView(container)) out.push([container, ask]);
    }
    for (const [container, ask] of out) {
      if (!this.#queue.park(container, ask)) continue;
      this.#shown.delete(container);
      this.#pending.delete(container);
      this.#painted.delete(container);
      this.#painters.delete(container);
      container.replaceChildren();
    }
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
      const key = keyOf(session, index);
      if (was?.session === session && was.key === key) continue;
      if (!this.#paint(session, index, container, all ? Infinity : IDLE_PAINT_BUDGET))
        this.#pending.set(container, index);
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
