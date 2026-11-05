// Slide show navigation: the Slide Show view as plain state. No DOM: an adapter paints
// the slide on show (`slideshow-view.ts`) and feeds keys and clicks in.
//
// Rules kept here:
// - "Next" and "Previous" skip hidden slides (`p:sld/@show="0"`).
// - "From current slide" and "number, then Enter" show the slide asked for, hidden or not.
// - After the last slide comes a black end screen; the next advance ends the show.
// - B (or `.`) and W (or `,`) toggle a black or white screen; the next advance returns to the
//   slide without moving.

/** One slide as the show sees it. @public */
export interface SlideshowSlide {
  /** `p:sld/@show="0"`: next and previous skip the slide. */
  readonly hidden: boolean;
}

/** What the show displays: the slide, a black or white screen, or the end screen. @public */
export type SlideshowScreen = 'slide' | 'black' | 'white' | 'end';

/** The state of a running show. @public */
export interface SlideshowState {
  /** The slide on show; on the end screen, the slide shown last. */
  readonly index: number;
  /** All slides, hidden ones too. */
  readonly slideCount: number;
  /** What is on the screen. */
  readonly screen: SlideshowScreen;
  /** Digits typed so far for "number, then Enter". */
  readonly typed: string;
  /** True once the show has ended (Escape, or an advance past the end screen). */
  readonly done: boolean;
}

/** A key press as the show reads it (a `KeyboardEvent` fits). @public */
export interface SlideshowKey {
  /** The `KeyboardEvent.key` value. */
  readonly key: string;
  /** Modifier state, as on a `KeyboardEvent`. */
  readonly ctrlKey?: boolean;
  /** Modifier state, as on a `KeyboardEvent`. */
  readonly metaKey?: boolean;
  /** Modifier state, as on a `KeyboardEvent`. */
  readonly altKey?: boolean;
  /** Modifier state, as on a `KeyboardEvent`. */
  readonly shiftKey?: boolean;
}

/** Where a show starts: the first slide shown, or a slide index. @public */
export type SlideshowStart = 'start' | number;

/**
 * The show a key starts, by the slide show shortcuts: F5 and Ctrl/Cmd+Shift+Enter from the
 * start; Shift+F5 and Ctrl/Cmd+Enter from the current slide. Null for any other key. @public
 */
export function slideshowShortcut(event: SlideshowKey): 'start' | 'current' | null {
  const mod = !!(event.ctrlKey || event.metaKey);
  if (event.altKey) return null;
  if (event.key === 'F5' && !mod) return event.shiftKey ? 'current' : 'start';
  if (event.key === 'Enter' && mod) return event.shiftKey ? 'start' : 'current';
  return null;
}

const NEXT_KEYS = new Set(['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Spacebar', 'n', 'N']);
const PREVIOUS_KEYS = new Set(['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p', 'P']);
const BLACK_KEYS = new Set(['b', 'B', '.']);
const WHITE_KEYS = new Set(['w', 'W', ',']);
/** Slide numbers have at most this many digits (a deck holds far fewer slides). */
const MAX_TYPED = 5;

/**
 * Navigation state of one slide show. Create one per show; it ends itself (`done`) on Escape
 * or after the end screen. @public
 */
export class SlideshowController {
  #slides: readonly SlideshowSlide[];
  #index: number;
  #screen: SlideshowScreen = 'slide';
  #typed = '';
  #done = false;
  #state: SlideshowState | null = null;
  readonly #listeners = new Set<() => void>();

  constructor(slides: readonly SlideshowSlide[], from: SlideshowStart = 'start') {
    this.#slides = slides;
    if (from === 'start') {
      const first = this.#visibleAfter(-1);
      this.#index = first ?? 0;
      if (first === undefined) this.#screen = 'end';
    } else {
      this.#index = this.#clamp(from);
      if (slides.length === 0) this.#screen = 'end';
    }
  }

  /** The current state. The same object until something changes. */
  get state(): SlideshowState {
    this.#state ??= {
      index: this.#index,
      slideCount: this.#slides.length,
      screen: this.#screen,
      typed: this.#typed,
      done: this.#done,
    };
    return this.#state;
  }

  /** Call `listener` after every change. Returns the unsubscribe function. */
  subscribe(listener: () => void): () => void {
    this.#listeners.add(listener);
    return () => this.#listeners.delete(listener);
  }

  /** Advance: leave a blank screen, else the next shown slide, else the end screen, else end. */
  next(): void {
    if (this.#done) return;
    if (this.#screen === 'black' || this.#screen === 'white')
      return this.#set(this.#index, 'slide');
    if (this.#screen === 'end') return this.exit();
    const next = this.#visibleAfter(this.#index);
    if (next === undefined) this.#set(this.#index, 'end');
    else this.#set(next, 'slide');
  }

  /** Go back: leave a blank or end screen, else the previous shown slide. */
  previous(): void {
    if (this.#done) return;
    if (this.#screen !== 'slide') {
      if (this.#slides.length > 0) this.#set(this.#index, 'slide');
      return;
    }
    const previous = this.#visibleBefore(this.#index);
    if (previous !== undefined) this.#set(previous, 'slide');
  }

  /** The first shown slide. */
  first(): void {
    const first = this.#visibleAfter(-1);
    if (!this.#done && first !== undefined) this.#set(first, 'slide');
  }

  /** The last shown slide. */
  last(): void {
    const last = this.#visibleBefore(this.#slides.length);
    if (!this.#done && last !== undefined) this.#set(last, 'slide');
  }

  /** Show slide `index` (0-based), hidden or not, as for a typed number. */
  goTo(index: number): void {
    if (this.#done || this.#slides.length === 0 || !Number.isFinite(index)) return;
    this.#set(this.#clamp(index), 'slide');
  }

  /** Toggle a black or white screen over the slide. */
  blank(color: 'black' | 'white'): void {
    if (this.#done || this.#screen === 'end') return;
    this.#set(this.#index, this.#screen === color ? 'slide' : color);
  }

  /** End the show. */
  exit(): void {
    if (this.#done) return;
    this.#done = true;
    this.#typed = '';
    this.#emit();
  }

  /** The deck changed while the show runs: keep the slide on show where it still exists. */
  update(slides: readonly SlideshowSlide[]): void {
    const same =
      slides.length === this.#slides.length &&
      slides.every((s, i) => s.hidden === this.#slides[i].hidden);
    this.#slides = slides;
    if (same) return;
    if (slides.length === 0) return this.#set(0, 'end');
    this.#set(this.#clamp(this.#index), this.#screen);
  }

  /**
   * Apply a key press. Returns true when the show used the key. Keys with Ctrl, Cmd, or Alt
   * are left to the host.
   */
  key(event: SlideshowKey): boolean {
    if (this.#done || event.ctrlKey || event.metaKey || event.altKey) return false;
    const { key } = event;
    if (/^[0-9]$/.test(key)) {
      if (this.#typed.length < MAX_TYPED) this.#typed += key;
      this.#emit();
      return true;
    }
    const typed = this.#typed;
    this.#typed = '';
    if (key === 'Enter' && typed) {
      const n = Number(typed);
      if (n >= 1) this.goTo(n - 1);
      else this.#emit();
      return true;
    }
    if (key === 'Enter' || NEXT_KEYS.has(key)) this.next();
    else if (PREVIOUS_KEYS.has(key)) this.previous();
    else if (key === 'Home') this.first();
    else if (key === 'End') this.last();
    else if (key === 'Escape' || key === 'Esc') this.exit();
    else if (BLACK_KEYS.has(key)) this.blank('black');
    else if (WHITE_KEYS.has(key)) this.blank('white');
    else {
      if (typed) this.#emit();
      return false;
    }
    if (typed && this.#typed === '') this.#emit();
    return true;
  }

  #clamp(index: number): number {
    return Math.max(0, Math.min(Math.trunc(index), this.#slides.length - 1));
  }

  #visibleAfter(index: number): number | undefined {
    for (let i = index + 1; i < this.#slides.length; i += 1) if (!this.#slides[i].hidden) return i;
    return undefined;
  }

  #visibleBefore(index: number): number | undefined {
    for (let i = index - 1; i >= 0; i -= 1) if (!this.#slides[i].hidden) return i;
    return undefined;
  }

  #set(index: number, screen: SlideshowScreen): void {
    if (index === this.#index && screen === this.#screen && this.#state?.typed === this.#typed)
      return;
    this.#index = index;
    this.#screen = screen;
    this.#emit();
  }

  #emit(): void {
    this.#state = null;
    for (const l of [...this.#listeners]) l();
  }
}
