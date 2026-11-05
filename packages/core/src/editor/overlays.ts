// The framework-free part of the overlays every adapter draws over the slide canvas: the find
// bar, the link popup, the edit notice, and the font notice. The React and Vue parts hold only
// their own state and markup, and call these helpers for the requests, the keys, the texts, and
// where focus goes, so the two adapters cannot drift.
//
// Every export is adapter glue, not a stable API: `@internal`.

import type { Translate } from '@pptx-editor.dev/i18n';
import type { EditorRejection, EditorRemoval, FindMatch, FindOptions } from '../contracts/index.ts';
import { translateKeyOrText } from './adapter-support.ts';
import { comparePoints } from '../automation/commands/text-model.ts';
import type { PptxEditorInstance } from './pptx-editor.ts';
import { surfaceRootOf } from './surface-registry.ts';
import type { PopoverAnchorRect } from './chrome-popover.ts';
import { queryLocalFontSources } from './local-fonts.ts';
import { withFontOrigin } from './host-fonts.ts';
import { sanitizeHref } from '../store/index.ts';
import {
  createHyperlinkPopupController,
  openHyperlinkPopup,
  type HyperlinkPopupController,
  type HyperlinkPopupState,
} from './link-popup.ts';

/** The event a find request travels on. */
const FIND_REQUEST = 'pptx-find-request';
/** The event a link request travels on (Ctrl/Cmd+K dispatches it from the surface too). */
const LINK_REQUEST = 'pptx-link-request';

/**
 * Open the find bar of `editor`: to find, or with the replace row when `replace`. The request
 * names the editor, so a bar answers only its own editor. @internal
 */
export function dispatchFindRequest(editor: object | null, replace = false): void {
  if (!editor || typeof document === 'undefined') return;
  document.dispatchEvent(new CustomEvent(FIND_REQUEST, { detail: { editor, replace } }));
}

/**
 * The element an event started on. In a shadow root, `target` at the document is the shadow
 * host, so the composed path gives the real element inside it.
 */
function eventOrigin(e: Event): EventTarget | null {
  return (typeof e.composedPath === 'function' ? e.composedPath()[0] : null) ?? e.target;
}

/** For each editor: whether the last press or focus on the page was in it. */
const lastWorkedIn = new WeakMap<object, boolean>();

/**
 * Listen for the requests a find bar answers: `dispatchFindRequest` for `editor`, and
 * Ctrl/Cmd+F (find) or Ctrl/Cmd+H (replace) while focus is in the bar's editor scope. With focus
 * on nothing (a closed menu or a removed control drops it to `<body>`), the key opens the bar
 * only when the last press or focus on the page was in this editor. Elsewhere on the host page,
 * the browser's own find keeps the key. `scope` finds the bar's `.pptx-editor`. In the slide
 * sorter, which shows no slide to find on, the request shows why through the edit notice
 * instead. Returns the function that stops. @internal
 */
export function listenFindRequests(
  editor: PptxEditorInstance,
  scope: () => Element | null | undefined,
  open: (replace: boolean) => void
): () => void {
  const answer = (replace: boolean): void => {
    if (editor.snapshot().presentationView === 'slideSorter') editor.reject('disabled.findSorter');
    else open(replace);
  };
  const onRequest = (e: Event): void => {
    const d = (e as CustomEvent<{ editor?: unknown; replace?: boolean } | null>).detail;
    if (d?.editor === editor) answer(!!d.replace);
  };
  const inside = (target: EventTarget | null): boolean => {
    const own = scope()?.closest('.pptx-editor');
    return !!own && target instanceof Node && own.contains(target);
  };
  // Whether the person last worked in this editor: a press or focus there, not on the host page.
  // Kept for the editor, so a listener an adapter adds again (the bar opened or closed) keeps it.
  if (!lastWorkedIn.has(editor)) lastWorkedIn.set(editor, inside(document.activeElement));
  const onActivity = (e: Event): void => {
    lastWorkedIn.set(editor, inside(eventOrigin(e)));
  };
  const onKey = (e: KeyboardEvent): void => {
    const key = e.key.toLowerCase();
    if (!(e.metaKey || e.ctrlKey) || e.altKey || (key !== 'f' && key !== 'h')) return;
    const onNothing = e.target === document.body || e.target === document.documentElement;
    if (!inside(eventOrigin(e)) && !(onNothing && lastWorkedIn.get(editor))) return;
    e.preventDefault();
    e.stopPropagation();
    answer(key === 'h');
  };
  document.addEventListener(FIND_REQUEST, onRequest);
  document.addEventListener('keydown', onKey, true);
  document.addEventListener('pointerdown', onActivity, true);
  document.addEventListener('focusin', onActivity, true);
  return () => {
    document.removeEventListener(FIND_REQUEST, onRequest);
    document.removeEventListener('keydown', onKey, true);
    document.removeEventListener('pointerdown', onActivity, true);
    document.removeEventListener('focusin', onActivity, true);
  };
}

/**
 * What the find bar shows for a query: the matches, the shown match, and the counter. @internal
 */
export interface FindBarView {
  readonly matches: readonly FindMatch[];
  /** The index of the shown match, or -1. */
  readonly index: number;
  /** The counter text: empty with no query. */
  readonly counter: string;
  /** The shown match is in the speaker notes. */
  readonly inNotes: boolean;
}

/** The find bar's matches and counter for `query`. Nothing while the bar is closed. @internal */
export function findBarView(
  t: Translate,
  editor: PptxEditorInstance | null,
  open: boolean,
  query: string,
  options: FindOptions
): FindBarView {
  const matches = open && editor && query ? editor.find.matches(query, options) : [];
  const index = matches.length && editor ? editor.find.position(query, options) : -1;
  // With matches but none shown (after Replace all whose text still matches), the counter
  // gives the count alone: "0 of N" would read as no match.
  const counter = !query
    ? ''
    : matches.length === 0
      ? t('find.noResults')
      : index < 0
        ? t('find.total', { total: matches.length })
        : t('find.counter', { current: index + 1, total: matches.length });
  return { matches, index, counter, inNotes: matches[index]?.where === 'notes' };
}

/**
 * Run `reset` each time `editor` opens another document: a find bar closes and forgets the
 * old search, which names text of the old deck. Returns the function that stops. @internal
 */
export function onNewDocument(editor: PptxEditorInstance, reset: () => void): () => void {
  let session = editor.session;
  return editor.subscribe(() => {
    if (editor.session === session) return;
    session = editor.session;
    if (session) reset();
  });
}

/**
 * Whether focus goes back to the find field after `action`: a Replace all that leaves no match
 * disables the buttons, and a disabled button that has focus drops it to the page. @internal
 */
export function findFieldTakesFocus(
  editor: PptxEditorInstance | null,
  action: FindBarAction,
  query: string,
  options: FindOptions
): boolean {
  return (
    !!editor &&
    action !== 'next' &&
    action !== 'previous' &&
    !editor.find.matches(query, options).length
  );
}

/** The find bar's actions. Each returns the status line to show, or null. @internal */
export type FindBarAction = 'next' | 'previous' | 'replace' | 'replaceAll';

/** Run one find bar action on `editor`, and return the status line, or null. @internal */
export function runFindBarAction(
  t: Translate,
  editor: PptxEditorInstance | null,
  action: FindBarAction,
  query: string,
  replacement: string,
  options: FindOptions
): string | null {
  if (!editor || !query) return null;
  const refusal = (r: { reason: string; detail?: string }): string =>
    translateKeyOrText(t, r.reason, r.detail ? { detail: r.detail } : undefined);
  if (action === 'replaceAll') {
    const r = editor.find.replaceAll(query, replacement, options);
    return r.ok ? t('find.replaced', { count: r.count }) : refusal(r);
  }
  const r =
    action === 'next'
      ? editor.find.next(query, options)
      : action === 'previous'
        ? editor.find.previous(query, options)
        : editor.find.replace(query, replacement, options);
  return r.ok ? null : refusal(r);
}

/**
 * Find as you type: select the first match of `query` at or after the start of the selection,
 * so the counter shows "1 of N" before Enter. A shown match that still matches stays, and a
 * longer query keeps the match it extends. With no match, the selection does not change.
 * @internal
 */
export function findAsYouType(
  editor: PptxEditorInstance | null,
  query: string,
  options: FindOptions
): void {
  if (!editor || !query || editor.find.position(query, options) >= 0) return;
  if (editor.find.matches(query, options).length === 0) return;
  const sel = editor.snapshot().selection;
  if (sel.kind === 'text') {
    // From the start of the selection: the match the shorter query showed can be the one now.
    const start = comparePoints(sel.anchor, sel.focus) <= 0 ? sel.anchor : sel.focus;
    editor.setSelection({ ...sel, anchor: start, focus: start });
  }
  editor.find.next(query, options);
}

/** The element that has focus now, to give it back when the bar closes. @internal */
export function findBarOpener(): HTMLElement | null {
  const active = typeof document === 'undefined' ? null : document.activeElement;
  return active instanceof HTMLElement ? active : null;
}

/**
 * Give focus back when the find bar closes. A match the bar selected is text on the slide: the
 * caret goes there, ready to type. Otherwise focus goes back to the opener, or to the slide.
 * @internal
 */
export function focusAfterFindBar(
  editor: PptxEditorInstance | null,
  bar: Element | null | undefined,
  scopeFrom: Element | null | undefined,
  opener: HTMLElement | null
): void {
  const scope = scopeFrom?.closest('.pptx-editor');
  const text = editor?.snapshot().selection.kind === 'text';
  const target = text
    ? scope?.querySelector<HTMLElement>('.pptx-surface__input')
    : opener && opener.isConnected && opener !== document.body && !bar?.contains(opener)
      ? opener
      : scope?.querySelector<HTMLElement>('.pptx-surface');
  target?.focus({ preventScroll: true });
}

/** The custom property that holds the find bar's height, set on the bar's parent. */
const FIND_RESERVE_PROPERTY = '--pptx-find-reserve';

/** True when two boxes overlap. */
const overlaps = (a: DOMRect, b: DOMRect): boolean =>
  a.left < b.right && b.left < a.right && a.top < b.bottom && b.top < a.bottom;

/**
 * Keep the find bar off the slide, and move the slide only when the bar would cover it. When
 * the open bar overlaps the slide (`.pptx-surface__stage` beside it), its height goes into
 * `--pptx-find-reserve` on its parent, and the stylesheet pads the canvas beside it
 * (`.pptx-canvas` before an open `.pptx-find`) by that much, so the fitted slide sits below the
 * bar and a zoomed slide scrolls out from under it. A bar that sits in the canvas margin
 * reserves nothing, so opening it does not move the slide. Once the bar reserves room, it keeps
 * it until it closes, so the slide does not move back and forth. A hidden bar reserves nothing.
 * Returns the cleanup. @internal
 */
export function reserveFindBarSpace(bar: HTMLElement | null): () => void {
  const parent = bar?.parentElement;
  if (!bar || !parent) return () => {};
  const view = bar.ownerDocument.defaultView;
  const observer = view?.ResizeObserver ? new view.ResizeObserver(() => later()) : null;
  let stage: HTMLElement | null = null;
  let reserved = 0;
  // Without a slide to measure, the bar reserves its height, as it would cover the slide.
  const covers = (): boolean => {
    const slide = parent.querySelector<HTMLElement>('.pptx-surface__stage');
    if (slide !== stage) {
      if (stage) observer?.unobserve(stage);
      // A zoom changes the slide's box: the bar checks again.
      if (slide) observer?.observe(slide);
      stage = slide;
    }
    return !slide || overlaps(bar.getBoundingClientRect(), slide.getBoundingClientRect());
  };
  const update = (): void => {
    if (bar.hidden) reserved = 0;
    else if (reserved > 0 || covers()) reserved = bar.offsetHeight;
    parent.style.setProperty(FIND_RESERVE_PROPERTY, `${reserved}px`);
  };
  update();
  // The write resizes the canvas, which the slide surface observes: it waits for the next
  // frame, so no observer sees a size change inside another observer's callback (a loop error).
  let frame = 0;
  const later = (): void => {
    if (!view?.requestAnimationFrame) return update();
    if (!frame) frame = view.requestAnimationFrame(() => ((frame = 0), update()));
  };
  observer?.observe(bar);
  return () => {
    observer?.disconnect();
    if (frame) view?.cancelAnimationFrame(frame);
    parent.style.removeProperty(FIND_RESERVE_PROPERTY);
  };
}

/** A request for the link popup. @internal */
export interface LinkRequest {
  readonly editor?: unknown;
  readonly anchor?: HTMLElement | null;
}

/**
 * Ask the root of `editor` to open its link popup, as Ctrl/Cmd+K does. The popup opens under
 * `anchor`, else under the toolbar's link button. @internal
 */
export function requestLinkPopup(editor: unknown, anchor?: HTMLElement | null): void {
  if (typeof document === 'undefined') return;
  document.dispatchEvent(
    new CustomEvent<LinkRequest>(LINK_REQUEST, { detail: { editor, anchor } })
  );
}

/** Whether an element is laid out (not inside `display: none` or a hidden More panel). */
function shown(el: HTMLElement | null | undefined): el is HTMLElement {
  return !!el && el.isConnected && el.getClientRects().length > 0;
}

/**
 * Where the link popup opens: the request's anchor, the visible toolbar link button, the open
 * menu bar trigger, or the editor's slide. A menu row never anchors it: the row unmounts with
 * its menu at once. @internal
 */
export function linkPopupAnchor(request: LinkRequest, editor: object): HTMLElement | null {
  if (shown(request.anchor) && !request.anchor.closest('.pptx-menu')) return request.anchor;
  const surface = surfaceRootOf(editor);
  const scope = surface?.closest<HTMLElement>('.pptx-editor') ?? null;
  const button = scope?.querySelector<HTMLElement>('[data-slot="text.link"]');
  if (shown(button) && !button.closest('.pptx-menu')) return button;
  // The open menu bar trigger of this editor only: another editor on the page has its own.
  const trigger = scope?.querySelector<HTMLElement>('.pptx-menubar__trigger[aria-expanded="true"]');
  if (shown(trigger)) return trigger;
  return surface;
}

/**
 * The box of the selected text on the slide, in client pixels, for the link popup to open at:
 * the drawn selection highlight, else the caret. Null when no text is selected or the surface
 * shows none, and the popup then opens at its anchor element. @internal
 */
export function linkSelectionRect(editor: object): PopoverAnchorRect | null {
  const surface = surfaceRootOf(editor);
  if (!surface) return null;
  const boxes = [...surface.querySelectorAll('.pptx-surface__text-selection')]
    .map((mark) => mark.getBoundingClientRect())
    .filter((box) => box.width > 0 || box.height > 0);
  if (boxes.length === 0) {
    const caret = surface.querySelector('.pptx-surface__caret')?.getBoundingClientRect();
    if (caret && (caret.width > 0 || caret.height > 0)) boxes.push(caret);
  }
  if (boxes.length === 0) return null;
  return {
    left: Math.min(...boxes.map((box) => box.left)),
    right: Math.max(...boxes.map((box) => box.right)),
    top: Math.min(...boxes.map((box) => box.top)),
    bottom: Math.max(...boxes.map((box) => box.bottom)),
  };
}

/**
 * Listen for link requests for `editor` (the toolbar, a menu row, Ctrl/Cmd+K). `open` gets the
 * anchor; `enabled` says whether the link command can run now. Returns the function that stops.
 * @internal
 */
export function listenLinkRequests(
  editor: object,
  enabled: () => boolean,
  open: (anchor: HTMLElement | null) => void
): () => void {
  // Requests for this root's editor only: another editor on the page opens its own.
  const onRequest = (e: Event): void => {
    const request = (e as CustomEvent<LinkRequest | null>).detail ?? {};
    if (request.editor !== editor || !enabled()) return;
    open(linkPopupAnchor(request, editor));
  };
  document.addEventListener(LINK_REQUEST, onRequest);
  return () => document.removeEventListener(LINK_REQUEST, onRequest);
}

/**
 * Whether focus should go back to the slide after the link popup closed: it is on nothing, in
 * the link popup, or on this editor's own chrome. Focus that moved elsewhere on purpose stays there. @internal
 */
export function linkPopupReturnsFocus(editor?: object): boolean {
  const active = document.activeElement;
  // The slide itself (Escape gives focus to the popup's anchor, which can be the slide).
  const slide = editor ? surfaceRootOf(editor) : null;
  if (slide && active && slide.contains(active)) return true;
  return (
    !active ||
    active === document.body ||
    active.closest('.pptx-menubar, .pptx-surface, .pptx-toolbar, .pptx-link-popup') !== null
  );
}

/** The message for a refusal: its reason, with the detail where the reason names one. @internal */
export function editRejectionMessage(t: Translate, rejection: EditorRejection): string {
  const detail = rejection.detail ?? '';
  return translateKeyOrText(t, rejection.reason, { reason: detail, detail });
}

/** Whether the edit notice shows `rejection`: a refusal not dismissed, and not a no-op. @internal */
export function editNoticeShows(
  rejection: EditorRejection | null | undefined,
  dismissed: EditorRejection | null
): rejection is EditorRejection {
  // A command with nothing to do (Ctrl+Z with no history) is not a refusal worth a notice.
  return !!rejection && !rejection.noop && rejection !== dismissed;
}

/**
 * Whether the open link popup already shows `rejection`: its Apply was refused, and the popup
 * says why under the field. The edit notice then stays hidden, so the refusal shows once.
 * @internal
 */
export function linkPopupShowsRejection(
  popup: HyperlinkPopupState | null | undefined,
  rejection: EditorRejection | null | undefined
): boolean {
  return (
    !!popup &&
    !!rejection &&
    popup.mode !== 'closed' &&
    popup.error !== null &&
    popup.error === rejection.reason
  );
}

/**
 * How long the edit notice offers to undo a removal, in milliseconds: long enough to find and
 * reach the action. A pointer on the notice or the focus in it holds it longer. @internal
 */
export const REMOVAL_NOTICE_MS = 10_000;

/** The message for a removal: what was deleted, and how many when more than one. @internal */
export function editRemovalMessage(t: Translate, removal: EditorRemoval): string {
  const key =
    removal.kind === 'slides'
      ? removal.count === 1
        ? 'notice.slideDeleted'
        : 'notice.slidesDeleted'
      : removal.count === 1
        ? 'notice.shapeDeleted'
        : 'notice.shapesDeleted';
  return translateKeyOrText(t, key, { count: String(removal.count) });
}

/**
 * Whether the edit notice shows `removal`: a removal that is not dismissed and has not timed
 * out. A refusal shows first: it is newer, or the removal was cleared. @internal
 */
export function removalNoticeShows(
  removal: EditorRemoval | null | undefined,
  closed: number | null
): removal is EditorRemoval {
  return !!removal && removal.serial !== closed;
}

/**
 * The families the font notice lists: those that render with a substitute that can change line
 * breaks. Metric-compatible substitutes are not listed. @internal
 */
export function fontNoticeFamilies(
  substituted: readonly { readonly family: string; readonly metricCompatible: boolean }[]
): string[] {
  return substituted.filter((f) => !f.metricCompatible).map((f) => f.family);
}

/**
 * Add the fonts installed on this computer for `families`, after the host's own fonts, and
 * return the translation key of the outcome. @internal
 */
export async function applyLocalFonts(
  editor: PptxEditorInstance,
  families: readonly string[]
): Promise<'fonts.localDenied' | 'fonts.localNone' | 'fonts.localAdded'> {
  const result = await queryLocalFontSources(families);
  if (!result.ok) return result.reason === 'denied' ? 'fonts.localDenied' : 'fonts.localNone';
  // The local faces go after the host's own fonts; the host's origins stay in force.
  await editor.setFonts(withFontOrigin(editor.fonts, { sources: result.sources }));
  return 'fonts.localAdded';
}

// ---------------------------------------------------------------------------------------------
// The link popup's state, after DOCX Editor's `useHyperlinkPopup`. One store per root holds
// whether the popup shows, what it shows, and the draft. The packaged popup, a host's own
// panel (`useHyperlinkPopup`), and the root's request listener read and change the same store,
// so the React and Vue adapters only subscribe to it.

/** The link a popup shows. @public */
export interface HyperlinkPopupLink {
  /** The target as the file holds it. */
  readonly url: string;
  /** The target that `sanitizeHref` allows, or null for a target that never opens. */
  readonly href: string | null;
}

/** The actions of the link popup, as DOCX Editor's `UseHyperlinkPopupResult`. @public */
export interface HyperlinkPopupActions {
  /** Open the link at the caret in reading mode, or the URL field when there is none. A null link requests editing. Caller-supplied targets never replace the selected target. */
  readonly open: (
    link?: { readonly url: string; readonly href?: string | null } | null,
    anchor?: HTMLElement | null
  ) => void;
  /**
   * Open for the selection, as Ctrl/Cmd+K and the toolbar's link control do: in editing mode,
   * with the link of the selection in the field when it has one.
   */
  readonly openAtCaret: (anchor?: HTMLElement | null) => void;
  /** Hide the popup. */
  readonly close: () => void;
  /** Copy the safe target. False when there is nothing safe to copy. */
  readonly copy: () => Promise<boolean>;
  /** Change to editing mode. */
  readonly beginEdit: () => void;
  /** Change the draft target. */
  readonly setUrl: (url: string) => void;
  /** Apply the draft to the selection. False when the engine refused it. */
  readonly commitEdit: () => boolean;
  /** Remove the link from the selection. The text stays. */
  readonly unlink: () => boolean;
  /** Open the safe target in a new tab. False for a target that never opens. */
  readonly openTarget: () => boolean;
}

/** One link popup: its state, a subscription, and its actions. @internal */
export interface HyperlinkPopupStore extends HyperlinkPopupActions {
  readonly controller: () => HyperlinkPopupController | null;
  readonly getState: () => HyperlinkPopupState;
  readonly subscribe: (listener: () => void) => () => void;
  /** The editor the actions run on. A new editor closes the popup. */
  readonly setEditor: (editor: PptxEditorInstance | null) => void;
}

/** The closed popup. @internal */
export const CLOSED_HYPERLINK_POPUP: HyperlinkPopupState = Object.freeze({
  mode: 'closed',
  href: null,
  anchor: null,
  url: '',
  draft: '',
  copied: false,
  error: null,
  canEdit: false,
});

/** The popup link for a target, or null for no target. @internal */
export function hyperlinkPopupLink(url: string | null | undefined): HyperlinkPopupLink | null {
  if (!url) return null;
  const safe = sanitizeHref(url);
  return { url, href: safe.ok ? safe.href : null };
}

/**
 * One store around the shared link controller. A root and its custom panels use the same
 * reading state, draft, safety checks, and selection lifetime. @internal
 */
export function createHyperlinkPopupStore(): HyperlinkPopupStore {
  let editor: PptxEditorInstance | null = null;
  let controller: HyperlinkPopupController | null = null;
  let stop: (() => void) | undefined;
  const listeners = new Set<() => void>();
  const notify = (): void => {
    for (const listener of [...listeners]) listener();
  };
  return {
    controller: () => controller,
    getState: () => controller?.getState() ?? CLOSED_HYPERLINK_POPUP,
    subscribe: (listener) => {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    setEditor: (next) => {
      if (next === editor) return;
      stop?.();
      controller?.dispose();
      editor = next;
      controller = next ? createHyperlinkPopupController(next) : null;
      stop = controller?.subscribe(notify);
      notify();
    },
    open: (link, anchor) => {
      if (controller && editor) openHyperlinkPopup(controller, editor, link, anchor);
    },
    openAtCaret: (anchor) => controller?.openEditing(anchor ?? null),
    close: () => controller?.close(),
    copy: () => controller?.copy() ?? Promise.resolve(false),
    beginEdit: () => controller?.beginEdit(),
    setUrl: (url) => controller?.setUrl(url),
    commitEdit: () => controller?.commitEdit() ?? false,
    unlink: () => controller?.unlink() ?? false,
    openTarget: () => controller?.openTarget() ?? false,
  };
}

/** The actions of a store, without its state and subscription. @internal */
export function hyperlinkPopupActions(store: HyperlinkPopupStore): HyperlinkPopupActions {
  const { open, openAtCaret, close, copy, beginEdit, setUrl, commitEdit, unlink, openTarget } =
    store;
  return { open, openAtCaret, close, copy, beginEdit, setUrl, commitEdit, unlink, openTarget };
}

/**
 * Open `store` for each link request of `editor` (the toolbar, a menu row, Ctrl/Cmd+K), while
 * `enabled` says that the link command can run. Returns the function that stops. @internal
 */
export function listenHyperlinkPopup(
  store: HyperlinkPopupStore,
  editor: object,
  enabled: () => boolean
): () => void {
  return listenLinkRequests(editor, enabled, (anchor) => store.openAtCaret(anchor));
}
