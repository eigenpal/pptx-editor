// The background half of a progressive open: the first slide shows and takes input, and the
// other slides parse in short tasks. After DOCX Editor's `surface-progressive-open.ts` and
// `docx-editor-opening-layout.ts`: a large open ends its first task as soon as the first view
// can paint, the rest runs in slices of about 12 ms on message-channel tasks, a slice ends early
// when input waits, the screen hears about progress at most every 300 ms, and a reader of the
// whole document finishes the rest at once.
//
// What runs in the background here is the PARSE of the slide and notes slide parts
// (`store/package/deferred-parts.ts`): layout is already lazy, one slide at a time
// (`DocumentSession.layout`). A deferred part parses on first read wherever it is read, so
// every command, search, undo, and save is correct before the background work ends; the
// background only does that work early. The slides the rail shows parse first.
//
// The queue holds the deferred PART OBJECTS of the package that opened, not names: an undo can
// bring back a slide that an edit removed, and its part must be parsed too before the open
// counts as settled.

import {
  deferredPartRejection,
  deferredRejectionCount,
  isDeferredPending,
  type PackagePart,
  type PackageRejection,
  type PptxPackage,
  type PresentationInfo,
} from '../store/index.ts';

/**
 * The content size past which an open is progressive: the bytes of the slides and notes
 * slides after the first one, uncompressed. DOCX Editor's open yield uses the same 512 KiB
 * (`OPEN_PAINT_YIELD_CONTENT_BYTES`). A smaller deck parses whole before it shows, as before.
 */
export const PROGRESSIVE_OPEN_CONTENT_BYTES = 512 * 1024;

/** The work one background task aims to do, in milliseconds: room for a frame and for input. */
export const OPENING_SLICE_MS = 12;

/**
 * The shortest time between two progress notices. Each notice rebuilds the slide summaries
 * and re-renders the rail, so it follows the clock; the parse itself does not.
 */
export const OPENING_PUBLISH_INTERVAL_MS = 300;

type XmlRejection = Extract<PackageRejection, { reason: 'xml' }>;

/** What the background tasks need from the editor. */
export interface OpeningHooks {
  /** Whether this open is still the editor's (no later load, no destroy). */
  current(): boolean;
  /** The pending slide count moved: publish a snapshot. */
  publish(): void;
  /** Every deferred part parsed. */
  settled(): void;
  /** A deferred part did not parse: the deck is refused. */
  failed(rejection: XmlRejection): void;
}

/** The background parse of one opened package. */
export class ProgressiveOpen {
  readonly #queue: PackagePart[];
  readonly #byName = new Map<string, PackagePart>();
  #cursor = 0;
  #wanted: PackagePart[] = [];
  #settled = false;
  #rejection: XmlRejection | null = null;
  #rejections: number;
  #pendingMemo: { pkg: PptxPackage | null; parsed: number; count: number } = {
    pkg: null,
    parsed: -1,
    count: 0,
  };
  #parsed = 0;

  /**
   * `pkg` is the package that opened; `order` the deferred part names in the order they
   * should parse (the slides after the start slide, each with its notes, then the ones before).
   */
  constructor(pkg: PptxPackage, order: readonly string[]) {
    const queued = new Set<PackagePart>();
    this.#queue = [];
    const add = (part: PackagePart | undefined): void => {
      if (!part || queued.has(part) || !isDeferredPending(part)) return;
      queued.add(part);
      this.#queue.push(part);
      this.#byName.set(part.partName, part);
    };
    for (const name of order) add(pkg.parts.get(name));
    // Any other deferred part (a notes slide no slide names) parses last.
    for (const part of pkg.parts.values()) add(part);
    this.#rejections = deferredRejectionCount();
  }

  /** Whether every deferred part of the opened package has parsed. */
  get settled(): boolean {
    return this.#settled;
  }

  /** The refusal that ended this open, if a deferred part did not parse. */
  get rejection(): XmlRejection | null {
    return this.#rejection;
  }

  /**
   * How many slides of `info` (the current package's slide list) have not parsed yet. Memoized
   * by package and by the parse count, so a snapshot pays for it once per change.
   */
  pendingSlides(pkg: PptxPackage, info: PresentationInfo): number {
    if (this.#settled) return 0;
    const memo = this.#pendingMemo;
    if (memo.pkg === pkg && memo.parsed === this.#parsed) return memo.count;
    let count = 0;
    for (const slide of info.slides)
      if (isDeferredPending(pkg.parts.get(slide.partName))) count += 1;
    this.#pendingMemo = { pkg, parsed: this.#parsed, count };
    return count;
  }

  /** Parse the deferred part `partName` before the others (a thumbnail in view asked for it). */
  want(partName: string): void {
    const part = this.#byName.get(partName);
    if (part && isDeferredPending(part) && !this.#wanted.includes(part)) this.#wanted.push(part);
  }

  /**
   * A part parsed somewhere (a reader touched it). Call after a commit or a read that may
   * have parsed parts, so the pending count and a failed parse are seen.
   */
  noteReads(): XmlRejection | null {
    if (deferredRejectionCount() === this.#rejections) return null;
    this.#rejections = deferredRejectionCount();
    for (const part of this.#queue) {
      const rejection = deferredPartRejection(part);
      if (rejection) return (this.#rejection = rejection);
    }
    return null;
  }

  /**
   * Parse for about `budgetMs`, ending early when `inputPending` says input waits; at least
   * one part per call. Answers `done` once every part parsed, and `wanted` when a part a
   * thumbnail asked for parsed in this call.
   */
  step(
    budgetMs: number,
    now: () => number,
    inputPending: () => boolean
  ): { readonly done: boolean; readonly wanted: boolean } {
    const deadline = now() + budgetMs;
    let wanted = false;
    let first = true;
    while (!this.#settled && !this.#rejection) {
      if (!first && (now() >= deadline || inputPending())) break;
      first = false;
      const fromWanted = this.#wanted.length > 0;
      const part = fromWanted ? this.#wanted.shift()! : this.#queue[this.#cursor++];
      if (!part) {
        this.#settled = true;
        break;
      }
      if (!isDeferredPending(part) || part.kind !== 'xml') continue;
      void part.xml;
      this.#parsed += 1;
      if (fromWanted) wanted = true;
      const rejection = deferredPartRejection(part);
      if (rejection) this.#rejection = rejection;
    }
    if (this.#rejection) return { done: true, wanted };
    return { done: this.#settled, wanted };
  }

  /** Parse every part now, for a reader of the whole deck (save, a collaboration room). */
  complete(): XmlRejection | null {
    while (
      !this.step(
        Number.POSITIVE_INFINITY,
        () => 0,
        () => false
      ).done
    );
    return this.#rejection ?? this.noteReads();
  }
}

/** Queue `run` as a new task: a message port, since nested timers wait 4 ms between steps. */
export function queueTask(run: () => void): void {
  if (typeof MessageChannel !== 'function') {
    setTimeout(run, 0);
    return;
  }
  const channel = new MessageChannel();
  channel.port1.onmessage = () => {
    channel.port1.close();
    run();
  };
  channel.port2.postMessage(null);
}

/**
 * Whether the browser holds input the page has not handled yet (Chromium's
 * `navigator.scheduling.isInputPending`). Every other engine answers false: the slice then
 * ends on its time budget alone.
 */
export function browserInputPending(): boolean {
  const scheduling = (
    globalThis.navigator as
      (Navigator & { scheduling?: { isInputPending?: () => boolean } }) | undefined
  )?.scheduling;
  return scheduling?.isInputPending?.() ?? false;
}

/** Run `open` in background tasks until it settles, fails, or stops being current. */
export function continueOpening(open: ProgressiveOpen, hooks: OpeningHooks): void {
  const clock = globalThis.performance ?? Date;
  let publishedAt = clock.now();
  const slice = (): void => {
    if (!hooks.current()) return;
    const failed = open.noteReads();
    if (failed) return hooks.failed(failed);
    const { done, wanted } = open.step(OPENING_SLICE_MS, () => clock.now(), browserInputPending);
    if (open.rejection) return hooks.failed(open.rejection);
    if (done) return hooks.settled();
    if (wanted || clock.now() - publishedAt >= OPENING_PUBLISH_INTERVAL_MS) {
      publishedAt = clock.now();
      hooks.publish();
    }
    queueTask(slice);
  };
  queueTask(slice);
}

/**
 * The order the deferred parts of `info` parse in: each slide after `start` with its notes
 * slide, then the slides before it.
 */
export function openingOrder(info: PresentationInfo, start: number): string[] {
  const out: string[] = [];
  const n = info.slides.length;
  for (let k = 1; k < n; k += 1) {
    const slide = info.slides[(start + k) % n]!;
    out.push(slide.partName);
    if (slide.notesPart) out.push(slide.notesPart);
  }
  return out;
}
