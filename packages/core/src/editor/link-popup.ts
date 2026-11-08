// The link popup's behavior, framework-free, after DOCX Editor's `useHyperlinkPopup`
// (packages/react/src/editor/useHyperlinkPopup.ts at 6907dc963): one state per editor, two modes.
// `reading` shows the link under the caret after a click on it, with Copy, Edit, and Unlink.
// `editing` shows the URL field with Apply and Cancel, for Ctrl/Cmd+K, the link control,
// Insert > Link, and Edit. The React and Vue adapters keep only their glue (context, refs, and
// effects), so the two cannot drift.
//
// The engine decides WHAT IS SAFE: the target opens and copies only as `sanitizeHref` returns
// it, and Apply and Unlink run the one `setLink` command, which refuses an unsafe URL. A
// popup never edits the package itself. The display text is not a field: `setLink` links the
// text as it stands (the selection, or the word at the caret), and a link's text is edited on
// the slide.
//
// The panel opens under the slide's text input, which follows the caret and is never repainted,
// so the panel keeps its place and its scope when the slide repaints. Reading mode takes no
// focus: F6 moves focus into it (`listenLinkPanelKey`), and the root announces it in a polite
// live region (`linkPopupAnnouncement`). A viewer can read, copy, and open a link; Edit and
// Unlink show only when the link command can run. A `popups.hyperlink` renderer gets reading
// mode only when it asks for it (`definePopup(Component, { reading: true })`), so a renderer
// written as an edit form keeps opening for Ctrl/Cmd+K alone.
//
// Differences from DOCX Editor's hook, on purpose: no `setText` (no display-text field, above),
// and `error` is the refusal's translation key or text, not a boolean. `open(link, anchor)` takes
// DOCX Editor's arguments, but the panel shows the link at the caret: a `null` link asks for the
// URL field, and any other value only asks for reading mode.

import type { CustomShowSummary, Editor, EditorSelection } from '../contracts/index.ts';
import { sanitizeHref } from '../store/index.ts';
import type { Translate } from '@pptx-editor.dev/i18n';
import { customShowLinkLabel, slideLinkIndex, slideLinkLabel } from './link-fields.ts';
import { surfaceInputOf, surfaceRootOf } from './surface-registry.ts';

/** What the link popup shows. @public */
export type HyperlinkPopupMode =
  /** Not shown. */
  | 'closed'
  /** The link under the caret: its target, with Copy, Edit, and Unlink. */
  | 'reading'
  /** The URL field, for a new link or a change to one. */
  | 'editing';

/** The link popup's state. @public */
export interface HyperlinkPopupState {
  readonly mode: HyperlinkPopupMode;
  /** The link of the selection when the popup opened, as the file stores it, or `''`. */
  readonly url: string;
  /** `url` as `sanitizeHref` passes it, or null when it is not safe to open or copy. */
  readonly href: string | null;
  /** The URL field's draft, in edit mode. */
  readonly draft: string;
  /** The translation key (or engine text) of why the last Apply was refused, or null. */
  readonly error: string | null;
  /** True after Copy, until the next change: for a "Copied" confirmation. */
  readonly copied: boolean;
  /** Whether the link command can run now; a read-only editor trims Edit and Unlink. */
  readonly canEdit: boolean;
  /** The element the panel opens under: the caret for a click, the opener otherwise. */
  readonly anchor: HTMLElement | null;
}

/** The link popup's behavior, shared by the adapters. @internal */
export interface HyperlinkPopupController {
  getState(): HyperlinkPopupState;
  /** Be told after every change. Returns the function that stops. */
  subscribe(listener: () => void): () => void;
  /** Open in reading mode over the link at the caret; nothing happens with no link there. */
  openReading(anchor: HTMLElement | null): void;
  /** Open the URL field, seeded with the link of the selection. */
  openEditing(anchor: HTMLElement | null): void;
  /** Switch from reading to editing, at the same place. */
  beginEdit(): void;
  setUrl(url: string): void;
  /** Apply the draft. False when the engine refused it; the popup then stays open. */
  commitEdit(): boolean;
  /** Remove the link, keeping its text. False when the engine refused. */
  unlink(): boolean;
  /** Copy the safe target. False when there is none, it is a slide link, or the clipboard refused. */
  copy(): Promise<boolean>;
  /**
   * Open the safe target in a new tab, with no opener; a slide link shows its slide in the
   * editor instead. False when there is none.
   */
  openTarget(): boolean;
  close(): void;
  /** Stop following the editor. */
  dispose(): void;
}

const CLOSED: HyperlinkPopupState = Object.freeze({
  mode: 'closed' as const,
  url: '',
  href: null,
  draft: '',
  error: null,
  copied: false,
  canEdit: false,
  anchor: null,
});

/** The link of the selection now, or `''`. */
function linkOf(editor: Editor): string {
  return editor.snapshot().format?.link ?? '';
}

/** Whether the selection is a caret in text, not a range. @internal */
export function isTextCaret(selection: EditorSelection): boolean {
  return (
    selection.kind === 'text' &&
    selection.anchor.paragraph === selection.focus.paragraph &&
    selection.anchor.offset === selection.focus.offset
  );
}

/** A key for the selection, so a move of the caret can be told from a repaint. */
function selectionKey(selection: EditorSelection): string {
  return JSON.stringify(selection);
}

/** The safe target of `url`, or null. */
function safeHref(url: string): string | null {
  if (!url) return null;
  const safe = sanitizeHref(url);
  // A relative path (a file or a program the file links to) names nothing a browser can open:
  // it would resolve against the page. Only an address with a scheme, or a `#` form, opens.
  if (!safe.ok || !(safe.href.startsWith('#') || /^[a-z][a-z0-9+.-]*:/i.test(safe.href)))
    return null;
  return safe.href;
}

/** The link popup controller of one editor. @internal */
export function createHyperlinkPopupController(
  editor: Editor,
  clipboard: () => { writeText(text: string): Promise<void> } | undefined = () =>
    typeof navigator === 'undefined' ? undefined : navigator.clipboard
): HyperlinkPopupController {
  let state = CLOSED;
  // The selection when reading mode opened: a move of the caret closes the panel, as DOCX
  // Editor's closes when the caret leaves the link. Typing moves the caret too.
  let openedAt: string | null = null;
  const listeners = new Set<() => void>();
  const set = (next: HyperlinkPopupState): void => {
    state = next;
    for (const l of [...listeners]) l();
  };
  // Whether the engine takes a link here, or the removal of one: a read-only editor takes
  // neither, also over a link, so a viewer gets no Edit, no Unlink, and no edit form.
  const canEdit = (): boolean =>
    editor.can({ type: 'setLink', url: null }).ok ||
    editor.can({ type: 'setLink', url: 'https://example.com' }).ok;
  const stop = editor.subscribe(() => {
    if (state.mode !== 'reading') return;
    const snap = editor.snapshot();
    const link = snap.format?.link ?? '';
    if (selectionKey(snap.selection) !== openedAt || !link) set(CLOSED);
    // The link changed under the same caret (a peer's edit, an undo): show the new target.
    else if (link !== state.url) set({ ...state, url: link, href: safeHref(link), copied: false });
  });
  const close = (): void => {
    openedAt = null;
    if (state.mode !== 'closed') set(CLOSED);
  };
  return {
    getState: () => state,
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    openReading(anchor) {
      const snap = editor.snapshot();
      const url = snap.format?.link ?? '';
      if (!url || !isTextCaret(snap.selection)) return;
      openedAt = selectionKey(snap.selection);
      set({
        ...CLOSED,
        mode: 'reading',
        url,
        href: safeHref(url),
        canEdit: canEdit(),
        anchor,
      });
    },
    openEditing(anchor) {
      // A refused edit opens nothing: the form could only fail.
      if (!canEdit()) return;
      openedAt = null;
      const url = linkOf(editor);
      set({
        ...CLOSED,
        mode: 'editing',
        url,
        href: safeHref(url),
        draft: url,
        canEdit: canEdit(),
        anchor,
      });
    },
    beginEdit() {
      if (state.mode !== 'reading' || !state.canEdit) return;
      openedAt = null;
      set({ ...state, mode: 'editing', draft: state.url, error: null, copied: false });
    },
    setUrl(url) {
      if (state.mode !== 'editing') return;
      set({ ...state, draft: url, error: null });
    },
    commitEdit() {
      if (state.mode !== 'editing') return false;
      const r = editor.exec({ type: 'setLink', url: state.draft });
      if (r.ok) close();
      else set({ ...state, error: r.reason });
      return r.ok;
    },
    unlink() {
      if (state.mode === 'closed') return false;
      const r = editor.exec({ type: 'setLink', url: null });
      if (r.ok) close();
      else set({ ...state, error: r.reason });
      return r.ok;
    },
    async copy() {
      const href = state.href;
      const target = clipboard();
      // A slide link names a place in this deck, not an address to paste elsewhere.
      if (!href || !target || href.startsWith('#')) return false;
      try {
        await target.writeText(href);
      } catch {
        return false;
      }
      if (state.mode !== 'closed' && href === state.href) set({ ...state, copied: true });
      return true;
    },
    openTarget() {
      const href = state.href;
      if (href?.startsWith('#')) {
        // A slide link shows its slide in the editor. A custom show link has no one slide.
        const snap = editor.snapshot();
        const index = slideLinkIndex(href, snap.currentSlide, snap.slides.length);
        if (index === null) return false;
        close();
        editor.goToSlide(index);
        return true;
      }
      const doc =
        state.anchor?.ownerDocument ?? (typeof document === 'undefined' ? null : document);
      if (!href || !doc) return false;
      // A link element, not a window call: the target opens as a click on it would, with no
      // opener and no referrer.
      const a = doc.createElement('a');
      a.setAttribute('href', href);
      a.setAttribute('target', '_blank');
      a.setAttribute('rel', 'noopener noreferrer');
      a.click();
      return true;
    },
    close,
    dispose() {
      stop();
      listeners.clear();
    },
  };
}

/**
 * `useHyperlinkPopup().open(link, anchor)` in both adapters, in DOCX Editor's argument order:
 * reading mode over the link at a text caret, else the URL field; `link === null` asks for the
 * URL field at once. The panel reads the link at the caret, so a given `link` only asks for
 * reading mode. `anchor` defaults to the caret. @internal
 */
export function openHyperlinkPopup(
  controller: HyperlinkPopupController,
  editor: object,
  link: { readonly url: string } | null | undefined,
  anchor: HTMLElement | null | undefined
): void {
  const at = anchor ?? linkCaretAnchor(editor);
  if (link === null) return controller.openEditing(at);
  controller.openReading(at);
  if (controller.getState().mode === 'closed') controller.openEditing(at);
}

/** How far a press may move and still be a click, in CSS pixels: more for a finger. */
const CLICK_SLOP = 4;
const TOUCH_SLOP = 10;

/**
 * The element the reading panel opens under: the slide's text input, which sits on the caret
 * and is never repainted, else the slide. The painted caret is rebuilt on every repaint, and a
 * panel anchored to it would lose its place and its scope. @internal
 */
export function linkCaretAnchor(editor: object): HTMLElement | null {
  const root = surfaceRootOf(editor);
  const input = surfaceInputOf(editor);
  return input && input.getBoundingClientRect().height > 0 ? input : root;
}

/**
 * Open the reading panel when a click puts the caret inside a link, as DOCX Editor's surface
 * classifies a click on a link: a press that moved, a modifier click, or a range is not one,
 * and a keyboard move of the caret never opens it. `open` runs after the editor took the click.
 * Returns the function that stops. @internal
 */
export function listenLinkClicks(
  editor: Editor,
  open: (anchor: HTMLElement | null) => void,
  enabled: () => boolean = () => true,
  doc: Document | undefined = typeof document === 'undefined' ? undefined : document
): () => void {
  if (!doc) return () => {};
  const view = doc.defaultView;
  // On the document, with the surface found at press time: the slide mounts after the root
  // that listens, and remounts when the view changes.
  const onSlide = (e: Event): boolean => {
    const root = surfaceRootOf(editor);
    return !!root && root.contains(e.target as Node | null);
  };
  let down: { x: number; y: number; slop: number } | null = null;
  let frame = 0;
  const onDown = (e: PointerEvent): void => {
    const plain = e.button === 0 && !e.shiftKey && !e.altKey && !e.ctrlKey && !e.metaKey;
    const slop = e.pointerType === 'touch' ? TOUCH_SLOP : CLICK_SLOP;
    down = plain && onSlide(e) ? { x: e.clientX, y: e.clientY, slop } : null;
  };
  const onUp = (e: PointerEvent): void => {
    const start = down;
    down = null;
    if (!start || e.button !== 0) return;
    if (Math.hypot(e.clientX - start.x, e.clientY - start.y) > start.slop) return;
    // After the editor placed the caret and painted it.
    if (frame) view?.cancelAnimationFrame(frame);
    const check = (): void => {
      frame = 0;
      const snap = editor.snapshot();
      if (!enabled() || !snap.format?.link || !isTextCaret(snap.selection)) return;
      open(linkCaretAnchor(editor));
    };
    if (view?.requestAnimationFrame) frame = view.requestAnimationFrame(check);
    else check();
  };
  doc.addEventListener('pointerdown', onDown, true);
  doc.addEventListener('pointerup', onUp, true);
  return () => {
    if (frame) view?.cancelAnimationFrame(frame);
    doc.removeEventListener('pointerdown', onDown, true);
    doc.removeEventListener('pointerup', onUp, true);
  };
}

/**
 * F6 moves focus from the slide text into an open reading panel, so a keyboard user reaches its
 * actions: the panel takes no focus when it opens, and Tab in slide text types a tab. Only
 * while focus is in this editor's slide. Returns the function that stops. @internal
 */
export function listenLinkPanelKey(
  editor: object,
  reading: () => boolean,
  doc: Document | undefined = typeof document === 'undefined' ? undefined : document
): () => void {
  if (!doc) return () => {};
  const onKey = (e: KeyboardEvent): void => {
    if (e.key !== 'F6' || e.shiftKey || e.altKey || e.ctrlKey || e.metaKey || !reading()) return;
    const root = surfaceRootOf(editor);
    if (!root?.contains(doc.activeElement)) return;
    const scope = root.closest('.pptx-editor') ?? doc;
    const panel = scope.querySelector<HTMLElement>('.pptx-link-popup');
    const first = panel?.querySelector<HTMLElement>('a[href], button:not([disabled]), input');
    if (!first) return;
    e.preventDefault();
    first.focus();
  };
  doc.addEventListener('keydown', onKey, true);
  return () => doc.removeEventListener('keydown', onKey, true);
}

/** What the root's polite live region says for a state: the open link, or the copy. @internal */
export function linkPopupAnnouncement(
  state: HyperlinkPopupState,
  t: Translate,
  shows: readonly CustomShowSummary[] = []
): string {
  if (state.mode !== 'reading') return '';
  // A slide link says its slide ("Slide 3"), and a custom show link its show, not a `#` form.
  const url = slideLinkLabel(t, state.url) || customShowLinkLabel(t, state.url, shows) || state.url;
  return state.copied ? t('link.copied') : t('link.announce', { url });
}

/** The mark a `popups.hyperlink` renderer carries when it shows reading mode too. */
const READING = 'pptxHyperlinkReading';

/** Mark a `popups.hyperlink` renderer as one that shows reading mode. @internal */
export function markReadingPopup<F extends object>(renderer: F): F {
  return Object.assign(renderer, { [READING]: true });
}

/**
 * Whether a click into linked text opens reading mode for this `popups.hyperlink` entry: the
 * packaged panel does, a host renderer only when marked (`definePopup(C, { reading: true })`).
 * @internal
 */
export function hyperlinkEntryReads(entry: unknown): boolean {
  if (entry === undefined) return true;
  return (
    typeof entry === 'function' && (entry as unknown as Record<string, unknown>)[READING] === true
  );
}

/** What a `popups.hyperlink` renderer receives: the twin of the adapters' props. @internal */
export interface HyperlinkPopupRequest {
  readonly anchor: HTMLElement | null;
  readonly url: string;
  readonly mode?: 'reading' | 'editing';
  readonly onClose: () => void;
}

/**
 * A controller for a packaged popup rendered from its props alone, outside the root that owns
 * the shared state (a host that renders `HyperlinkPopup` itself). With an editor it runs the
 * real commands; with none it shows the props and its actions do nothing. Closing calls
 * `onClose`. @internal
 */
export function hyperlinkPopupFromRequest(
  editor: Editor | null,
  request: HyperlinkPopupRequest
): HyperlinkPopupController {
  if (editor) {
    const real = createHyperlinkPopupController(editor);
    if (request.mode === 'reading') real.openReading(request.anchor);
    if (real.getState().mode === 'closed') real.openEditing(request.anchor);
    if (real.getState().mode !== 'closed') {
      real.subscribe(() => {
        if (real.getState().mode === 'closed') request.onClose();
      });
      return real;
    }
    // The engine refuses an edit here (a viewer, no text selected): the host asked for the
    // popup, so it shows the props, and its actions do nothing.
    real.dispose();
  }
  const state: HyperlinkPopupState = {
    ...CLOSED,
    mode: request.mode ?? 'editing',
    url: request.url,
    href: safeHref(request.url),
    draft: request.url,
    anchor: request.anchor,
  };
  return {
    getState: () => state,
    subscribe: () => () => {},
    openReading: () => {},
    openEditing: () => {},
    beginEdit: () => {},
    setUrl: () => {},
    commitEdit: () => false,
    unlink: () => false,
    copy: async () => false,
    openTarget: () => false,
    close: () => request.onClose(),
    dispose: () => {},
  };
}
