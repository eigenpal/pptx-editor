// Slide thumbnails for the rail.
//
// A container is painted once, then again only when its slide renders differently
// (`renderKey`). A container that changes slide (or document) paints at once, so it never shows
// the wrong slide. A slide that changed in place (typing, a drag, undo, a theme edit) repaints
// when the editor is idle: after a quiet time, and at least once a second while edits go on.
// The painted picture is always the slide as it is when the repaint runs.

import type { SlideLayout } from '../layout/index.ts';
import { paintSlide } from '../output/index.ts';
import type { PptxPackage } from '../store/index.ts';
import type { DocumentSession } from './document-session.ts';
import { renderKey } from './slide-dependencies.ts';

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
/** Idle callbacks run at the latest after this. */
const IDLE_TIMEOUT_MS = 500;

type IdleDeadline = { timeRemaining(): number } | undefined;
type IdleWindow = {
  requestIdleCallback?: (cb: (d: IdleDeadline) => void, o?: { timeout: number }) => number;
  cancelIdleCallback?: (handle: number) => void;
};

export class ThumbnailPainter {
  readonly #host: ThumbnailHost;
  #painted = new WeakMap<HTMLElement, Painted>();
  #layouts = new WeakMap<HTMLElement, { pkg: PptxPackage; id: string; fonts: number }>();
  /** Containers waiting for an idle repaint, with the slide index asked for last. */
  #pending = new Map<HTMLElement, number>();
  #timer: ReturnType<typeof setTimeout> | null = null;
  #idle: { cancel(): void } | null = null;
  /** When the oldest waiting repaint was asked for (0: none waits). */
  #since = 0;
  #count = 0;

  constructor(host: ThumbnailHost) {
    this.#host = host;
  }

  /** Paint slide `index` into `container`, now or when idle (see the file comment). */
  render(index: number, container: HTMLElement): void {
    const session = this.#host.session();
    if (!session) return;
    const key = renderKey(session.store.package, index, session.fontGeneration);
    const part = session.info.slides[index]?.partName;
    if (!key || !part) return;
    const was = this.#painted.get(container);
    if (was?.session === session && was.key === key) {
      this.#pending.delete(container);
      return;
    }
    if (!was || was.session !== session || was.part !== part) {
      this.#pending.delete(container);
      this.#paint(session, index, container);
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
    this.#drain(undefined, true);
  }

  /** Drop waiting repaints (the editor is destroyed or loads another document). */
  clear(): void {
    this.#cancel();
    this.#pending.clear();
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
    const w = globalThis as IdleWindow;
    if (typeof w.requestIdleCallback === 'function') {
      const handle = w.requestIdleCallback((d) => this.#drain(d, false), {
        timeout: IDLE_TIMEOUT_MS,
      });
      this.#idle = { cancel: () => w.cancelIdleCallback?.(handle) };
    } else {
      const handle = setTimeout(() => this.#drain(undefined, false), 0);
      this.#idle = { cancel: () => clearTimeout(handle) };
    }
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
