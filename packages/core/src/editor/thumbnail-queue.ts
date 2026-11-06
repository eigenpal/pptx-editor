// Which slide thumbnails paint first.
//
// The rail and the slide sorter mount one container per slide. Of the first paints asked for
// in one task, the first `BURST` paint at once: the top of a list that just mounted, where
// nothing is laid out yet. Only a container asked for the first time takes a burst place: a
// container asked for again (a slide that finished opening, a rail that renders again and
// releases its items first) is checked like the others, so it paints only in view. The others are checked at the end of the task, all layout reads
// first: a container in its scroll box (or within `OVERSCAN` of it) paints before the frame;
// the others wait, and paint when they come into view or when the browser is idle. A waiting
// container that showed another slide is emptied, so no container shows a wrong slide. A
// container that is not on the page when it is asked for (an export, a print, a paint before
// the host places it) cannot be out of view: it paints at once.
//
// An `IntersectionObserver` on the scroll box watches the waiting containers only, so a rail
// with every thumbnail painted costs the browser no intersection work while it edits. Where
// there is no observer (happy-dom, a server), the waiting containers paint when idle.
//
// A waiting container off the page (a rail kept alive while hidden) is held weakly and looked
// for again every `RECHECK_MS`, at most `RECHECK_TRIES` times: it paints when it is back. The
// adapters release a container when they unmount it (`renderThumbnail`'s disposer), so a closed rail
// leaves nothing waiting, no observer, and no timer.

/** First paints in one task that paint at once, before any layout read: one screen of slides. */
const BURST = 12;
/** How far outside its scroll box a container counts as in view, in CSS pixels. */
const OVERSCAN = 200;
/** Idle callbacks run at the latest after this. */
const IDLE_TIMEOUT_MS = 500;
/** Paint time for each slice where the browser reports none (`requestIdleCallback`). */
const SLICE_MS = 8;
/** How often the queue looks again for waiting containers that are off the page. */
const RECHECK_MS = 1000;
/** How many times the queue looks for a container off the page before it forgets it. */
const RECHECK_TRIES = 30;

/** A first paint asked for: the slide index, and the slide part it showed then. */
export interface Ask {
  readonly index: number;
  readonly part: string;
}

/** The idle period a callback runs in; `undefined` when the browser gives none. */
export type IdleDeadline = { timeRemaining(): number; readonly didTimeout?: boolean } | undefined;

type IdleWindow = {
  requestIdleCallback?: (cb: (d: IdleDeadline) => void, o?: { timeout: number }) => number;
  cancelIdleCallback?: (handle: number) => void;
};

/** Run `cb` when the browser is idle (`requestIdleCallback`), else after a timeout. */
export function idleCallback(cb: (d: IdleDeadline) => void): { cancel(): void } {
  const w = globalThis as IdleWindow;
  if (typeof w.requestIdleCallback === 'function') {
    const handle = w.requestIdleCallback(cb, { timeout: IDLE_TIMEOUT_MS });
    return { cancel: () => w.cancelIdleCallback?.(handle) };
  }
  const handle = setTimeout(() => cb(undefined), 0);
  return { cancel: () => clearTimeout(handle) };
}

type Observer = { observe(e: Element): void; unobserve(e: Element): void; disconnect(): void };
type ObserverEntry = { readonly target: Element; readonly isIntersecting: boolean };
type ObserverClass = new (
  cb: (entries: readonly ObserverEntry[]) => void,
  options: { root: Element | null; rootMargin: string }
) => Observer;

const SCROLLS = /auto|scroll|overlay/;

/** What the queue asks of the painter. */
export interface QueueHost {
  /** Paint the slide asked for into `container` now (unless something painted it since). */
  paint(container: HTMLElement, ask: Ask): void;
  /** Empty `container`: it waits, and showed another slide. */
  blank(container: HTMLElement): void;
}

/** The first-paint queue of one thumbnail painter. */
export class ThumbnailQueue {
  readonly #host: QueueHost;
  /** Containers to check at the end of the task, and whether each shows another slide. */
  readonly #checking = new Map<HTMLElement, Ask & { stale: boolean }>();
  /** Containers waiting for their first paint, with the slide asked for last. */
  readonly #waiting = new Map<HTMLElement, Ask>();
  /** Waiting containers off the page, held weakly, with the times the queue looked for each. */
  readonly #away = new Map<WeakRef<HTMLElement>, { ask: Ask; tries: number }>();
  /** Waiting containers that came into view; they paint before the others. */
  readonly #soon = new Set<HTMLElement>();
  /** One observer for each scroll box, and how many containers it watches. */
  readonly #observers = new Map<Element | null, { observer: Observer; count: number }>();
  /** The scroll box of each container, and of each ancestor walked to find one. */
  #rootOf = new WeakMap<Element, Element | null>();
  /** The containers an observer watches, with the scroll box it watches them in. */
  readonly #watched = new Map<Element, Element | null>();
  #burst = 0;
  /** Every container asked for since the last `clear`: only a new one takes a burst place. */
  #asked = new WeakSet<HTMLElement>();
  #idle: { cancel(): void } | null = null;
  #soonTimer: ReturnType<typeof setTimeout> | null = null;
  #recheck: ReturnType<typeof setTimeout> | null = null;

  constructor(host: QueueHost) {
    this.#host = host;
  }

  /**
   * Whether slide `index` should paint into `container` now. When not, the container is
   * checked at the end of the task: in view, it paints then; else it waits (and is emptied
   * when it is `stale`: it shows another slide).
   */
  paintNow(container: HTMLElement, ask: Ask, stale: boolean): boolean {
    this.#forgetAway(container);
    // A container not on the page (an export, a print, a paint before it is placed) cannot be
    // out of view: it paints at once.
    if (!container.isConnected) {
      this.#checking.delete(container);
      this.#waiting.delete(container);
      return true;
    }
    const first = !this.#asked.has(container);
    this.#asked.add(container);
    if (first && this.#takeBurst()) return true;
    if (this.#checking.size === 0) queueMicrotask(() => this.#settle());
    const was = this.#checking.get(container);
    this.#checking.set(container, { ...ask, stale: stale || !!was?.stale });
    return false;
  }

  /**
   * The container painted, needs no paint, or left for good (the disposer of `renderThumbnail`): it waits no
   * more, and no observer watches it.
   */
  drop(container: HTMLElement): void {
    this.#checking.delete(container);
    this.#waiting.delete(container);
    this.#soon.delete(container);
    this.#unwatch(container);
    this.#forgetAway(container);
    // Nothing left to paint: no idle callback or timer stays behind.
    if (this.#waiting.size === 0 && this.#soon.size === 0) {
      this.#idle?.cancel();
      this.#idle = null;
      if (this.#soonTimer !== null) clearTimeout(this.#soonTimer);
      this.#soonTimer = null;
    }
  }

  /** Whether anything waits, is watched, or is scheduled (for tests). @internal */
  get busy(): boolean {
    return (
      this.#checking.size + this.#waiting.size + this.#away.size + this.#observers.size > 0 ||
      this.#idle !== null ||
      this.#soonTimer !== null ||
      this.#recheck !== null
    );
  }

  #forgetAway(container: HTMLElement): void {
    if (this.#away.size === 0) return;
    for (const ref of this.#away.keys()) if (ref.deref() === container) this.#away.delete(ref);
    if (this.#away.size === 0 && this.#recheck !== null) {
      clearTimeout(this.#recheck);
      this.#recheck = null;
    }
  }

  /** Paint every waiting container now. */
  flush(): void {
    this.#settle();
    this.#cancel();
    const all = [...this.#waiting];
    this.#waiting.clear();
    this.#soon.clear();
    for (const [container, ask] of all) {
      this.#unwatch(container);
      this.#host.paint(container, ask);
    }
  }

  /** Forget every container (another document, or the editor is destroyed). */
  clear(): void {
    this.#cancel();
    this.#checking.clear();
    this.#waiting.clear();
    this.#away.clear();
    this.#soon.clear();
    for (const { observer } of this.#observers.values()) observer.disconnect();
    this.#observers.clear();
    this.#watched.clear();
    this.#rootOf = new WeakMap();
    this.#asked = new WeakSet();
  }

  #cancel(): void {
    this.#idle?.cancel();
    this.#idle = null;
    if (this.#soonTimer !== null) clearTimeout(this.#soonTimer);
    this.#soonTimer = null;
    if (this.#recheck !== null) clearTimeout(this.#recheck);
    this.#recheck = null;
  }

  #takeBurst(): boolean {
    if (this.#burst === 0) queueMicrotask(() => (this.#burst = 0));
    this.#burst += 1;
    return this.#burst <= BURST;
  }

  /** At the end of the task: read every container's box, then paint the ones in view. */
  #settle(): void {
    const all = [...this.#checking];
    this.#checking.clear();
    const shown = all.map(([container]) => this.#inView(container));
    all.forEach(([container, { stale, ...ask }], i) => {
      if (shown[i]) {
        this.#waiting.delete(container);
        this.#unwatch(container);
        this.#host.paint(container, ask);
        return;
      }
      if (stale) this.#host.blank(container);
      this.#waiting.set(container, ask);
      this.#watch(container);
    });
    if (this.#waiting.size > 0) this.#schedule();
  }

  /** Whether a container is laid out in its scroll box (or the window), within `OVERSCAN`. */
  #inView(container: HTMLElement): boolean {
    const view = container.ownerDocument?.defaultView;
    if (!view || !container.isConnected) return false;
    const box = container.getBoundingClientRect();
    if (box.width === 0 || box.height === 0) return false;
    const root = this.#scrollBox(container, view);
    const r = root
      ? root.getBoundingClientRect()
      : { left: 0, top: 0, right: view.innerWidth, bottom: view.innerHeight };
    return (
      box.right >= r.left - OVERSCAN &&
      box.left <= r.right + OVERSCAN &&
      box.bottom >= r.top - OVERSCAN &&
      box.top <= r.bottom + OVERSCAN
    );
  }

  /** Observe a waiting container, on the observer of its scroll box. */
  #watch(container: HTMLElement): void {
    if (this.#watched.has(container)) return;
    const view = container.ownerDocument?.defaultView as
      (Window & { IntersectionObserver?: ObserverClass }) | null;
    const Observer = view?.IntersectionObserver;
    if (!view || typeof Observer !== 'function' || !container.isConnected) return;
    const root = this.#scrollBox(container, view);
    let entry = this.#observers.get(root);
    if (!entry) {
      const observer = new Observer((entries) => this.#onView(entries), {
        root,
        rootMargin: `${OVERSCAN}px`,
      });
      entry = { observer, count: 0 };
      this.#observers.set(root, entry);
    }
    entry.count += 1;
    entry.observer.observe(container);
    this.#watched.set(container, root);
  }

  /** Stop observing a container; an observer with nothing left to watch goes. */
  #unwatch(container: Element): void {
    if (!this.#watched.has(container)) return;
    const root = this.#watched.get(container) ?? null;
    this.#watched.delete(container);
    const entry = this.#observers.get(root);
    if (!entry) return;
    entry.observer.unobserve(container);
    entry.count -= 1;
    if (entry.count > 0) return;
    entry.observer.disconnect();
    this.#observers.delete(root);
  }

  /** Waiting containers came into view: paint them, within a time slice; the rest soon. */
  #onView(entries: readonly ObserverEntry[]): void {
    for (const { target, isIntersecting } of entries) {
      const container = target as HTMLElement;
      if (!this.#waiting.has(container)) this.#unwatch(container);
      else if (!container.isConnected) this.#leave(container);
      else if (isIntersecting) this.#soon.add(container);
      else this.#schedule();
    }
    this.#paintSoon();
  }

  #paintSoon(): void {
    const start = Date.now();
    for (const container of [...this.#soon]) {
      if (Date.now() - start >= SLICE_MS) break;
      this.#soon.delete(container);
      const ask = this.#waiting.get(container);
      if (!ask || !container.isConnected) continue;
      this.#waiting.delete(container);
      this.#unwatch(container);
      this.#host.paint(container, ask);
    }
    if (this.#soon.size > 0 && this.#soonTimer === null)
      this.#soonTimer = setTimeout(() => {
        this.#soonTimer = null;
        this.#paintSoon();
      }, 0);
  }

  /** The nearest ancestor that scrolls, or null for the viewport. */
  #scrollBox(container: HTMLElement, view: Window): Element | null {
    const cached = this.#walk(container, view, true);
    // A container moved to another list: walk again without the remembered boxes.
    return cached && !cached.contains(container) ? this.#walk(container, view, false) : cached;
  }

  #walk(container: HTMLElement, view: Window, useCache: boolean): Element | null {
    const walked: Element[] = [];
    let found: Element | null = null;
    for (let el = container.parentElement; el; el = el.parentElement) {
      const known = useCache ? this.#rootOf.get(el) : undefined;
      if (known !== undefined) {
        found = known;
        break;
      }
      walked.push(el);
      const style = view.getComputedStyle(el);
      if (SCROLLS.test(`${style.overflowX} ${style.overflowY}`)) {
        found = el;
        break;
      }
    }
    // Each ancestor walked remembers the scroll box of what it holds.
    for (const el of walked) this.#rootOf.set(el, found);
    return found;
  }

  #schedule(): void {
    if (!this.#idle) this.#idle = idleCallback((d) => this.#drain(d));
  }

  /** Paint waiting containers while idle time lasts; continue in the next idle period. */
  #drain(deadline: IdleDeadline): void {
    this.#idle = null;
    const start = Date.now();
    const left = (): number => deadline?.timeRemaining() ?? SLICE_MS - (Date.now() - start);
    // A callback that ran on its timeout paints one slide; one with no time left paints none.
    let painted = 0;
    let more = false;
    for (const [container, ask] of [...this.#waiting]) {
      if (!container.isConnected) {
        this.#leave(container);
        continue;
      }
      if (left() < 2 && !(painted === 0 && deadline?.didTimeout)) {
        more = true;
        break;
      }
      this.#waiting.delete(container);
      this.#soon.delete(container);
      this.#unwatch(container);
      this.#host.paint(container, ask);
      painted += 1;
    }
    if (more) this.#schedule();
  }

  /** A waiting container left the page: hold it weakly, and look for it again later. */
  #leave(container: HTMLElement): void {
    const ask = this.#waiting.get(container);
    this.#waiting.delete(container);
    this.#soon.delete(container);
    this.#unwatch(container);
    if (!ask) return;
    this.#away.set(new WeakRef(container), { ask, tries: 0 });
    this.#recheckLater();
  }

  #recheckLater(): void {
    if (this.#recheck !== null || this.#away.size === 0) return;
    this.#recheck = setTimeout(() => {
      this.#recheck = null;
      let back = false;
      for (const [ref, entry] of [...this.#away]) {
        const container = ref.deref();
        if (container?.isConnected) {
          this.#away.delete(ref);
          this.#waiting.set(container, entry.ask);
          this.#watch(container);
          back = true;
        } else if (!container || ++entry.tries >= RECHECK_TRIES) this.#away.delete(ref);
      }
      if (back) this.#schedule();
      this.#recheckLater();
    }, RECHECK_MS);
  }
}
