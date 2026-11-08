// The framework-free part of the slide show screen: full screen, the start keys, the show keys,
// the screen size, and links on the slide. The React and Vue `Slideshow` parts hold only their
// state and markup, and call these helpers, so the two adapters cannot drift. Navigation state
// is `SlideshowController` (slideshow.ts); paint is `paintSlideshowSlide` (slideshow-view.ts).
//
// Every export is adapter glue, not a stable API: `@internal`.

import { parseSlideLink } from '../automation/commands/slide-links.ts';
import type { DocumentSession } from './document-session.ts';
import type { PptxEditorInstance } from './pptx-editor.ts';
import { SlideshowController, slideshowShortcut, type SlideshowState } from './slideshow.ts';
import { slideshowLinkAt } from './slideshow-view.ts';

/** The control bar hides after the pointer rests this long. @internal */
export const SLIDESHOW_IDLE_MS = 2500;

/**
 * Per document, the full-screen request or exit still in flight. A show that starts right after
 * another ended waits for it: otherwise the old show's asynchronous exit takes the new show out
 * of full screen, and the new show reads that as the user leaving, and ends.
 */
const fullscreenSettling = new WeakMap<Document, Promise<unknown>>();

/**
 * Leave full screen and resolve once the document has really left it. Escape also makes the
 * browser leave full screen on its own, asynchronously; waiting for `fullscreenElement` to be
 * null covers both exits, so neither can land on the next show. The timeout keeps a browser
 * that never reports the change from blocking the next show.
 */
function leaveFullscreen(doc: Document, el: Element): Promise<void> {
  if (!doc.fullscreenElement) return Promise.resolve();
  return new Promise((resolve) => {
    const finish = (): void => {
      doc.removeEventListener('fullscreenchange', check);
      clearTimeout(timer);
      resolve();
    };
    const check = (): void => {
      if (!doc.fullscreenElement) finish();
    };
    const timer = setTimeout(finish, 1000);
    doc.addEventListener('fullscreenchange', check);
    if (doc.fullscreenElement === el && typeof doc.exitFullscreen === 'function')
      void Promise.resolve(doc.exitFullscreen()).catch(() => undefined);
    check();
  });
}

/** How many times each show element entered the screen: a late focus return of an earlier entry
 * (a development double mount) must not take focus from the show that runs now. */
const screenEntries = new WeakMap<HTMLElement, number>();

function afterFullscreenSettles(doc: Document, step: () => unknown): void {
  const next = (fullscreenSettling.get(doc) ?? Promise.resolve()).then(step).catch(() => undefined);
  fullscreenSettling.set(doc, next);
}

/** A slide item of the rail or the sorter, and the list that holds it. */
const SLIDE_ITEM = '.pptx-rail__item, .pptx-sorter__item';
const SLIDE_LIST = '.pptx-rail__list, .pptx-sorter__list';

/**
 * Show `el` full screen when the browser allows it, and focus it; the fixed overlay covers the
 * window otherwise. Returns the function that leaves full screen and gives focus back to the
 * element that had it.
 *
 * `onLeave` runs once when the document leaves full screen after `el` held it: the browser
 * takes the first Escape to leave full screen and does not give that key to the page, so the
 * show ends there too and one Escape is enough. Only an exit after this show itself held full
 * screen counts, so the late exit of an earlier show cannot end a show that just started.
 * @internal
 */
export function enterSlideshowScreen(el: HTMLElement, onLeave?: () => void): () => void {
  const doc = el.ownerDocument;
  const active = doc.activeElement;
  const before = active instanceof HTMLElement && active !== doc.body ? active : null;
  // When the element that had focus is gone (a menu row) or cannot take it, the slide canvas
  // of the show's editor takes it, never the page body.
  const canvas = el.closest('.pptx-editor')?.querySelector<HTMLElement>('.pptx-surface') ?? null;
  const entry = (screenEntries.get(el) ?? 0) + 1;
  screenEntries.set(el, entry);
  // A slide of the rail or the sorter had focus: the show ends on the slide shown last, which
  // is the current slide now, so its item takes focus in place of the item the show began on.
  const back = (): HTMLElement | null => {
    if (!before?.matches(SLIDE_ITEM) || before.getAttribute('aria-current') === 'true')
      return before;
    const list = before.closest<HTMLElement>(SLIDE_LIST);
    return list?.querySelector<HTMLElement>(`:is(${SLIDE_ITEM})[aria-current="true"]`) ?? before;
  };
  const giveBack = (): void => {
    if (screenEntries.get(el) !== entry) return;
    const now = doc.activeElement;
    if (now && now !== doc.body && now !== before && !el.contains(now)) return;
    const target = back();
    if (target?.isConnected && target !== now) target.focus({ preventScroll: true });
    const after = doc.activeElement;
    if ((!after || after === doc.body || el.contains(after)) && canvas?.isConnected)
      canvas.focus({ preventScroll: true });
  };
  el.focus({ preventScroll: true });
  let ended = false;
  let held = false;
  const onFullscreenChange = (): void => {
    if (doc.fullscreenElement === el) held = true;
    else if (held && !ended) {
      held = false;
      onLeave?.();
    }
  };
  doc.addEventListener('fullscreenchange', onFullscreenChange);
  // Refused or failed requests leave the overlay in place; `afterFullscreenSettles` swallows
  // the rejection.
  if (typeof el.requestFullscreen === 'function')
    afterFullscreenSettles(doc, () =>
      ended || !el.isConnected ? undefined : el.requestFullscreen()
    );
  return () => {
    ended = true;
    doc.removeEventListener('fullscreenchange', onFullscreenChange);
    // Focus goes back now, and again once full screen is left and the show is out of the page:
    // while the show holds full screen or focus, the page under it can refuse focus.
    afterFullscreenSettles(doc, () => leaveFullscreen(doc, el).then(giveBack));
    before?.focus({ preventScroll: true });
    if (typeof requestAnimationFrame === 'function') requestAnimationFrame(giveBack);
    else setTimeout(giveBack, 0);
  };
}

/**
 * A new show for `editor`, from the first shown slide or from the current slide. Null without an
 * open document with slides. Viewing mode can present. @internal
 */
export function startSlideshowController(
  editor: PptxEditorInstance | null,
  from: 'start' | 'current'
): SlideshowController | null {
  const snap = editor?.snapshot();
  if (!snap || snap.status !== 'ready' || snap.slides.length === 0) return null;
  return new SlideshowController(snap.slides, from === 'start' ? 'start' : snap.currentSlide);
}

/**
 * The show start keys (F5, Shift+F5, Ctrl/Cmd+Enter), while focus is in the editor scope of
 * `anchor` or on nothing, but not in a comment field (an element with `data-comment-field`).
 * The listener stays registered while a show runs and reads `running`, so a start key pressed
 * right after a show ends is never lost. Returns the function that stops.
 * @internal
 */
export function listenSlideshowStartKeys(
  anchor: () => Element | null | undefined,
  running: () => boolean,
  start: (from: 'start' | 'current') => void
): () => void {
  const onKey = (e: KeyboardEvent): void => {
    if (running()) return;
    const from = slideshowShortcut(e);
    const scope = anchor()?.closest('.pptx-editor');
    if (!from || !scope) return;
    const doc = scope.ownerDocument;
    const target = e.target instanceof Node ? e.target : null;
    const free = !target || target === doc.body || target === doc.documentElement;
    if (!free && !scope.contains(target)) return;
    // A comment field posts on Ctrl/Cmd+Enter.
    if (target instanceof Element && target.closest('[data-comment-field]')) return;
    e.preventDefault();
    e.stopPropagation();
    start(from);
  };
  window.addEventListener('keydown', onKey, true);
  return () => window.removeEventListener('keydown', onKey, true);
}

/**
 * Every key goes to the show: the editor underneath never sees one while it runs. Tab moves
 * through the control bar, and Enter and Space press its focused button. Returns the function
 * that stops. @internal
 */
export function listenSlideshowKeys(
  controller: SlideshowController,
  root: () => Element | null | undefined
): () => void {
  const onKey = (e: KeyboardEvent): void => {
    e.stopImmediatePropagation();
    const target = e.target instanceof HTMLElement ? e.target : null;
    const onButton = target?.tagName === 'BUTTON' && !!root()?.contains(target);
    if (e.key === 'Tab' || (onButton && (e.key === 'Enter' || e.key === ' '))) return;
    if (controller.key(e) || e.key === 'F5') e.preventDefault();
  };
  window.addEventListener('keydown', onKey, true);
  return () => window.removeEventListener('keydown', onKey, true);
}

/** The show screen size, in CSS pixels and the device pixel ratio. @internal */
export interface SlideshowScreenSize {
  readonly w: number;
  readonly h: number;
  readonly dpr: number;
}

/**
 * Measure the show screen now and on every resize. `onSize` gets only real changes. Returns the
 * function that stops. @internal
 */
export function observeSlideshowScreen(
  el: HTMLElement,
  onSize: (size: SlideshowScreenSize) => void
): () => void {
  const view = el.ownerDocument.defaultView;
  let last: SlideshowScreenSize | null = null;
  const measure = (): void => {
    const w = el.clientWidth || view?.innerWidth || 0;
    const h = el.clientHeight || view?.innerHeight || 0;
    const dpr = view?.devicePixelRatio || 1;
    if (last && last.w === w && last.h === h && last.dpr === dpr) return;
    last = { w, h, dpr };
    onSize(last);
  };
  measure();
  const RO = view?.ResizeObserver;
  const ro = RO ? new RO(measure) : null;
  ro?.observe(el);
  view?.addEventListener('resize', measure);
  return () => {
    ro?.disconnect();
    view?.removeEventListener('resize', measure);
  };
}

/**
 * The link under a pointer on the slide on show, through `sanitizeHref`, or null. `slide` is the
 * painted slide box; `width` and `height` are the slide size in EMU. @internal
 */
export function slideshowLinkUnder(
  session: DocumentSession | null,
  state: Pick<SlideshowState, 'index' | 'screen'>,
  slide: Element | null | undefined,
  width: number,
  height: number,
  clientX: number,
  clientY: number
): string | null {
  if (!session || state.screen !== 'slide') return null;
  const box = slide?.getBoundingClientRect();
  if (!box || box.width <= 0 || box.height <= 0) return null;
  const x = ((clientX - box.left) / box.width) * width;
  const y = ((clientY - box.top) / box.height) * height;
  if (x < 0 || y < 0 || x > width || y > height) return null;
  return slideshowLinkAt(session, state.index, x, y);
}

/**
 * Follow a link the show resolved (`slideshowLinkUnder`): a slide link moves the show, and any
 * other link opens in a new tab. @internal
 */
export function followSlideshowLink(
  controller: SlideshowController,
  doc: Document,
  href: string
): void {
  const target = parseSlideLink(href);
  if (!target) return openSlideshowLink(doc, href);
  if ('slide' in target) return controller.goTo(target.slide);
  controller[target.jump]();
}

/** Open a link the show resolved through `sanitizeHref`, in a new tab without an opener. @internal */
export function openSlideshowLink(doc: Document, href: string): void {
  const a = doc.createElement('a');
  a.setAttribute('href', href);
  a.setAttribute('target', '_blank');
  a.setAttribute('rel', 'noopener noreferrer');
  a.click();
}
