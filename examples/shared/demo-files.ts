// The demo's file handling, shared by the React demo (examples/vite) and the Vue demo
// (examples/vue): the decks it opens, the download helper, and the slide exports. It imports only
// `@pptx-editor.dev/core`, so it has no framework code. Everything loads and saves in the
// browser. Only the PDF export sends the deck to the demo's own server, when the user picks it.

import {
  exportSlidePng,
  exportSlideSvg,
  exportSlidesPngZip,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { canonicalPackageFingerprint, type PptxPackage } from '@pptx-editor.dev/core/store';

export const PPTX_MIME =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation';

/** The width of exported slide pictures, in pixels (Full HD). */
export const PICTURE_WIDTH = 1920;

/** What File > Export makes besides the deck. */
export type SlideExport = 'pdf' | 'png' | 'zip' | 'svg';

/** The demo's PDF export function: the root `api/convert.ts`, also served by the dev servers. */
export const PDF_ENDPOINT = '/api/convert';

/**
 * The decks the demo opens: the sample presentation (`showcase.pptx`, every feature, with demo
 * notes) by default, and the basic five-slide deck (`sample.pptx`) that the e2e tests use, with
 * `?deck=basic`. Only the sample presentation is in the File menu.
 */
export const EXAMPLES = {
  showcase: { file: 'showcase.pptx', title: 'Sample presentation' },
  basic: { file: 'sample.pptx', title: 'Basic sample' },
} as const;
export type Example = keyof typeof EXAMPLES;

/** The blank deck that New opens. */
export const BLANK = { file: 'blank.pptx', title: 'Untitled presentation' } as const;

/** The deck the URL asks for: `?deck=basic`, or the sample presentation. */
export function initialExample(search: string): Example {
  return new URLSearchParams(search).get('deck') === 'basic' ? 'basic' : 'showcase';
}

/** Where the demo keeps the theme picked in File > Page setup. */
const THEME_KEY = 'pptx-editor-demo-theme';

type Theme = 'light' | 'dark';
const isTheme = (v: unknown): v is Theme => v === 'light' || v === 'dark';

/** The theme picked before, or null. Storage can be off (a private window): then null. */
function storedTheme(): Theme | null {
  try {
    const v = globalThis.localStorage?.getItem(THEME_KEY);
    return isTheme(v) ? v : null;
  } catch {
    return null;
  }
}

/**
 * The theme the demo starts in: the URL's (`?theme=dark` or `?theme=light`), else the theme
 * picked before, else light. Dark is only ever picked by hand.
 */
export function initialTheme(search: string): Theme {
  const asked = new URLSearchParams(search).get('theme');
  return isTheme(asked) ? asked : (storedTheme() ?? 'light');
}

/** Keep the theme picked in File > Page setup for the next visit. */
export function rememberTheme(theme: Theme): void {
  try {
    globalThis.localStorage?.setItem(THEME_KEY, theme);
  } catch {
    // Storage is off: the next visit starts light again.
  }
}

/** `?chrome=ribbon` shows the ribbon and the status bar in place of the flat toolbar. */
export function ribbonChrome(search: string): boolean {
  return new URLSearchParams(search).get('chrome') === 'ribbon';
}

/**
 * A deck download fails when no bytes arrive for this long. The limit is on progress, not on
 * the whole download, so a slow connection that keeps sending bytes still finishes.
 */
export const DECK_STALL_MS = 10_000;

/** A deck download that sent no bytes for `DECK_STALL_MS`. */
export class DeckStallError extends Error {
  constructor(file: string) {
    super(`No bytes of ${file} arrived for ${DECK_STALL_MS} ms`);
    this.name = 'DeckStallError';
  }
}

/** Options of `fetchDeck`. */
export interface FetchDeckOptions {
  /** Fail when no bytes arrive for this long. Defaults to `DECK_STALL_MS`. */
  readonly stallMs?: number;
  /** Cancels the download. */
  readonly signal?: AbortSignal;
}

/**
 * Fetch a deck that the demo serves. Rejects on a network error, an HTTP error, a stall (no
 * response or no body bytes for `stallMs`), or a cancel.
 */
export async function fetchDeck(
  base: string,
  file: string,
  { stallMs = DECK_STALL_MS, signal }: FetchDeckOptions = {}
): Promise<Uint8Array> {
  const controller = new AbortController();
  signal?.throwIfAborted();
  signal?.addEventListener('abort', () => controller.abort(signal.reason), { once: true });
  let timer: ReturnType<typeof setTimeout> | undefined;
  // The stall rejects the wait itself too: a host whose body stream ignores the abort
  // cannot hold the open.
  const stalled = new Promise<never>((_, reject) => {
    controller.signal.addEventListener('abort', () => reject(controller.signal.reason));
  });
  stalled.catch(() => undefined);
  const arm = (): void => {
    clearTimeout(timer);
    timer = setTimeout(() => controller.abort(new DeckStallError(file)), stallMs);
  };
  arm();
  try {
    const res = await Promise.race([
      fetch(`${base}${file}`, { signal: controller.signal }),
      stalled,
    ]);
    if (!res.ok) throw new Error(`HTTP ${res.status} for ${file}`);
    if (!res.body) return new Uint8Array(await Promise.race([res.arrayBuffer(), stalled]));
    const reader = res.body.getReader();
    const chunks: Uint8Array[] = [];
    let size = 0;
    try {
      for (;;) {
        arm();
        const next = await Promise.race([reader.read(), stalled]);
        if (next.done) break;
        chunks.push(next.value);
        size += next.value.byteLength;
      }
    } catch (cause) {
      void reader.cancel(cause).catch(() => undefined);
      throw cause;
    }
    const bytes = new Uint8Array(size);
    let at = 0;
    for (const chunk of chunks) (bytes.set(chunk, at), (at += chunk.byteLength));
    return bytes;
  } finally {
    clearTimeout(timer);
  }
}

/**
 * Download the deck the demo opens first. A deck that the user opens before it arrives cancels
 * it, so a late download never replaces the user's deck, and its failure shows no notice.
 * Returns the cancel.
 */
export function fetchFirstDeck(
  base: string,
  file: string,
  handlers: {
    readonly loaded: (bytes: Uint8Array) => void;
    readonly failed: (cause: unknown) => void;
  }
): () => void {
  const controller = new AbortController();
  fetchDeck(base, file, { signal: controller.signal }).then(
    (bytes) => {
      if (!controller.signal.aborted) handlers.loaded(bytes);
    },
    (cause: unknown) => {
      if (!controller.signal.aborted) handlers.failed(cause ?? new Error('fetch failed'));
    }
  );
  return () => controller.abort();
}

/** Say that a deck the demo serves did not load, in the editor's notice. */
export function refuseDeckFetch(
  editor: PptxEditorInstance | null | undefined,
  cause: unknown
): void {
  console.error('[demo] The presentation did not load', cause);
  editor?.reject('app.deckFetchFailed');
}

/** Save bytes or a blob as a file through a temporary link. */
export function download(data: Uint8Array | Blob, name: string, type = PPTX_MIME): void {
  const url = URL.createObjectURL(
    data instanceof Blob ? data : new Blob([data.slice().buffer], { type })
  );
  const a = document.createElement('a');
  a.setAttribute('href', url);
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/** Whether a dropped file is a deck the demo opens. */
export function isPptxFile(file: File | undefined): file is File {
  return !!file && /\.pptx$/i.test(file.name);
}

/** The title of an opened file: its name without the extension. */
export function fileTitle(file: File): string {
  return file.name.replace(/\.pptx$/i, '');
}

/** Save the open deck as a `.pptx` file. */
export function exportDeck(editor: PptxEditorInstance | null | undefined, name: string): void {
  const bytes = editor?.save();
  if (bytes) download(bytes, `${name}.pptx`);
}

/** The largest deck the demo's PDF function accepts (`MAX_UPLOAD` in `api/convert.ts`). */
export const PDF_MAX_UPLOAD = 4 * 1024 * 1024;

/** Why the PDF export failed, as the `app.*` key of the editor's notice. */
export class PdfExportFailure extends Error {
  constructor(
    readonly key: 'app.pdfTooLarge' | 'app.pdfTooComplex' | 'app.pdfBusy' | 'app.pdfExportFailed'
  ) {
    super(key);
    this.name = 'PdfExportFailure';
  }
}

/**
 * The notice for a failure of the PDF function: its HTTP status and the `error` name of its JSON
 * body. Only the name is read, never the server's message.
 */
function failureFor(status: number, error: unknown): PdfExportFailure {
  if (status === 413) return new PdfExportFailure('app.pdfTooLarge');
  // Too many slides, too much paint, a PDF past its size limit, or past the conversion deadline:
  // a retry fails the same way.
  if (status === 507 || (status === 408 && error === 'PdfWorkLimitError'))
    return new PdfExportFailure('app.pdfTooComplex');
  if (status === 503) return new PdfExportFailure('app.pdfBusy');
  return new PdfExportFailure('app.pdfExportFailed');
}

/**
 * Convert the saved deck to a PDF on the demo's server (`api/convert.ts`). The deck leaves the
 * browser only here, after the user picks the export. Throws `PdfExportFailure`.
 */
export async function convertToPdf(deck: Uint8Array): Promise<Blob> {
  if (deck.byteLength > PDF_MAX_UPLOAD) throw new PdfExportFailure('app.pdfTooLarge');
  let response: Response;
  try {
    response = await fetch(PDF_ENDPOINT, {
      method: 'POST',
      headers: { 'Content-Type': PPTX_MIME },
      body: deck.slice().buffer,
      // Longer than the function can run (`maxDuration` in `api/convert.ts`).
      signal: AbortSignal.timeout(130_000),
    });
  } catch {
    // Offline, no server, or past the timeout.
    throw new PdfExportFailure('app.pdfExportFailed');
  }
  const type = response.headers.get('content-type') ?? '';
  if (response.ok && type.startsWith('application/pdf')) return response.blob();
  const body: unknown = type.includes('application/json')
    ? await response.json().catch(() => null)
    : null;
  throw failureFor(response.status, (body as { error?: unknown } | null)?.error);
}

/** Whether a PDF export is running. A second pick does nothing until it ends. */
let pdfRunning = false;
const pdfListeners = new Set<(busy: boolean) => void>();

function setPdfRunning(busy: boolean): void {
  pdfRunning = busy;
  for (const listener of pdfListeners) listener(busy);
}

/**
 * Follow whether a PDF export is running, so the Export button shows that the PDF is being made.
 * Calls `listener` now and on each change. Returns the cleanup.
 */
export function watchPdfExport(listener: (busy: boolean) => void): () => void {
  pdfListeners.add(listener);
  listener(pdfRunning);
  return () => pdfListeners.delete(listener);
}

/**
 * Export from the editor's current deck: the whole deck as PDF (made on the demo's server), the
 * current slide as PNG or SVG, or every slide as PNG files in a ZIP. The pictures come from the
 * editor's own paint and never leave the browser. A failure shows as the editor's notice.
 */
export async function exportSlides(
  editor: PptxEditorInstance | null | undefined,
  kind: SlideExport,
  name: string
): Promise<void> {
  const session = editor?.session;
  if (!editor || !session) return;
  const slide = editor.snapshot().currentSlide;
  const options = { width: PICTURE_WIDTH };
  try {
    if (kind === 'pdf') {
      if (pdfRunning) return;
      // `save` is null for a damaged slide; the editor already shows why.
      const deck = editor.save();
      if (!deck) return;
      setPdfRunning(true);
      try {
        download(await convertToPdf(deck), `${name}.pdf`);
      } finally {
        setPdfRunning(false);
      }
    } else if (kind === 'png')
      download(await exportSlidePng(session, slide, options), `${name} ${slide + 1}.png`);
    else if (kind === 'zip')
      download(await exportSlidesPngZip(session, options), `${name}.zip`, 'application/zip');
    else {
      // EMF and WMF pictures convert in the background: wait for those of the slide.
      await session.picturesReady(slide);
      download(
        new Blob([exportSlideSvg(session, slide, options)], { type: 'image/svg+xml' }),
        `${name} ${slide + 1}.svg`
      );
    }
  } catch (e) {
    console.error(e);
    if (kind === 'pdf')
      editor.reject(e instanceof PdfExportFailure ? e.key : 'app.pdfExportFailed');
  }
}

/**
 * Up to this window width the side panel is a sheet under the slide (shell.css), so it starts
 * closed: open, it takes half of a phone's height. The panel toggle still opens it.
 */
export const NARROW_PANEL_QUERY = '(max-width: 900px)';

const narrowQuery = (): MediaQueryList | null =>
  typeof window !== 'undefined' && typeof window.matchMedia === 'function'
    ? window.matchMedia(NARROW_PANEL_QUERY)
    : null;

/** Where the demo keeps whether the side panel and the speaker notes are open. */
const PANEL_KEY = 'pptx-editor-demo-panel';
const NOTES_KEY = 'pptx-editor-demo-notes';

/** A stored open state, or null. Storage can be off (a private window): then null. */
function storedOpen(key: string): boolean | null {
  try {
    const v = globalThis.localStorage?.getItem(key);
    return v === 'open' ? true : v === 'closed' ? false : null;
  } catch {
    return null;
  }
}

function rememberOpen(key: string, open: boolean): void {
  try {
    globalThis.localStorage?.setItem(key, open ? 'open' : 'closed');
  } catch {
    // Storage is off: the next visit starts with the defaults.
  }
}

/**
 * Whether the side panel starts open: as the user left it, else open, except on a narrow window,
 * where it is a sheet that takes half the height.
 */
export function panelStartsOpen(): boolean {
  return storedOpen(PANEL_KEY) ?? !narrowQuery()?.matches;
}

/**
 * Keep the side panel open or closed for the next visit: call it for the user's own toggle. A
 * narrow window does not keep it: there the panel is a sheet that you open for a task and close
 * after it, and that close must not close the panel of a wide window.
 */
export function rememberPanel(open: boolean): void {
  if (!narrowQuery()?.matches) rememberOpen(PANEL_KEY, open);
}

/** Whether the speaker notes start open: as the user left them, else closed. */
export function notesStartOpen(): boolean {
  return storedOpen(NOTES_KEY) ?? false;
}

/** Keep the speaker notes open or closed for the next visit. */
export function rememberNotes(open: boolean): void {
  rememberOpen(NOTES_KEY, open);
}

/**
 * Close the side panel when the window becomes narrow (it becomes a sheet), and give it back
 * the user's choice when the window becomes wide again. Returns the cleanup.
 */
export function followPanelWidth(setOpen: (open: boolean) => void): () => void {
  const query = narrowQuery();
  if (!query) return () => {};
  const changed = (event: MediaQueryListEvent): void =>
    setOpen(!event.matches && (storedOpen(PANEL_KEY) ?? true));
  query.addEventListener('change', changed);
  return () => query.removeEventListener('change', changed);
}

/** The heights the panel sheet snaps to, as parts of the workspace height. */
const SHEET_SNAPS = [0.4, 0.6, 0.9] as const;
/** A release faster than this (px per ms) goes on to the next snap height, or closes. */
const FLICK_SPEED = 0.5;

/** The sheet height: the CSS variable, as a part of the workspace height. */
function sheetPart(sheet: HTMLElement): number {
  const host = sheet.parentElement?.getBoundingClientRect().height ?? 0;
  return host > 0 ? sheet.getBoundingClientRect().height / host : SHEET_SNAPS[0];
}

/**
 * The grip of the side panel sheet (shell.css, a narrow window): drag it to set the sheet's
 * height, and let go to snap to 40%, 60%, or 90% of the workspace. A quick flick goes to the
 * next height in its direction; a flick or drag down from the lowest height closes the panel.
 * Arrow Up and Arrow Down step through the heights, and a tap goes to the next height. The
 * height does not animate: the slide fits again at each height. Returns the cleanup.
 */
export function attachSheetGrip(
  grip: HTMLElement,
  options: { readonly sheet: () => HTMLElement | null; readonly close: () => void }
): () => void {
  let drag: { y: number; t: number; start: number; host: number } | null = null;
  let last = { y: 0, t: 0 };
  // A drag sets the height once a frame at most: the slide fits again at each height.
  let frame = 0;
  const apply = (sheet: HTMLElement, part: number): void => {
    sheet.style.setProperty('--demo-sheet-h', `${(part * 100).toFixed(2)}%`);
  };
  const set = (sheet: HTMLElement, part: number, now: boolean): void => {
    cancelAnimationFrame(frame);
    frame = 0;
    if (now) apply(sheet, part);
    else frame = requestAnimationFrame(() => apply(sheet, part));
  };
  const snapTo = (sheet: HTMLElement, part: number, velocity: number): void => {
    // A flick moves one snap height in its direction; a slow release takes the nearest.
    const nearest = SHEET_SNAPS.reduce((a, b) => (Math.abs(b - part) < Math.abs(a - part) ? b : a));
    let at = SHEET_SNAPS.indexOf(nearest);
    if (velocity < -FLICK_SPEED) at = Math.min(at + 1, SHEET_SNAPS.length - 1);
    if (velocity > FLICK_SPEED) at -= 1;
    if (at < 0 || part < SHEET_SNAPS[0] / 2) {
      options.close();
      return;
    }
    set(sheet, SHEET_SNAPS[at]!, true);
  };
  const down = (e: PointerEvent): void => {
    const sheet = options.sheet();
    if (!sheet || !e.isPrimary || drag) return;
    const host = sheet.parentElement?.getBoundingClientRect().height ?? 0;
    if (host <= 0) return;
    try {
      grip.setPointerCapture(e.pointerId);
    } catch {
      // The pointer is gone already (a fast tap): the drag still works over the grip.
    }
    drag = { y: e.clientY, t: e.timeStamp, start: sheetPart(sheet), host };
    last = { y: e.clientY, t: e.timeStamp };
  };
  const move = (e: PointerEvent): void => {
    const sheet = options.sheet();
    if (!drag || !sheet) return;
    let part = drag.start + (drag.y - e.clientY) / drag.host;
    // Past the highest snap the sheet slows down, as a drag at a hard edge does.
    const top = SHEET_SNAPS[SHEET_SNAPS.length - 1];
    if (part > top) part = top + (part - top) * 0.25;
    set(sheet, Math.max(0.1, part), false);
    last = { y: e.clientY, t: e.timeStamp };
  };
  const up = (e: PointerEvent): void => {
    const sheet = options.sheet();
    if (!drag || !sheet) return;
    const ms = Math.max(1, e.timeStamp - last.t);
    const velocity = (e.clientY - last.y) / ms;
    const moved = Math.abs(e.clientY - drag.y);
    drag = null;
    // A tap on the grip goes to the next height, and back to the first from the top.
    if (moved < 4) {
      const at = SHEET_SNAPS.findIndex((s) => s > sheetPart(sheet) + 0.02);
      set(sheet, SHEET_SNAPS[at < 0 ? 0 : at]!, true);
      return;
    }
    snapTo(sheet, sheetPart(sheet), velocity);
  };
  const key = (e: KeyboardEvent): void => {
    const sheet = options.sheet();
    if (!sheet || (e.key !== 'ArrowUp' && e.key !== 'ArrowDown')) return;
    e.preventDefault();
    const part = sheetPart(sheet);
    const raise = e.key === 'ArrowUp';
    const next = raise
      ? SHEET_SNAPS.find((s) => s > part + 0.02)
      : [...SHEET_SNAPS].reverse().find((s) => s < part - 0.02);
    if (next !== undefined) set(sheet, next, true);
  };
  grip.addEventListener('pointerdown', down);
  grip.addEventListener('pointermove', move);
  grip.addEventListener('pointerup', up);
  grip.addEventListener('pointercancel', up);
  grip.addEventListener('keydown', key);
  return () => {
    grip.removeEventListener('pointerdown', down);
    grip.removeEventListener('pointermove', move);
    grip.removeEventListener('pointerup', up);
    grip.removeEventListener('pointercancel', up);
    grip.removeEventListener('keydown', key);
  };
}

/**
 * The Escape handler of an overlay pane: close the open overlay, and give focus back to its
 * toggle when focus was in it. A menu, a field, or a popover that used this Escape keeps it,
 * and an overlay that closes takes it, so one Escape does one thing.
 */
function closeOnEscape(
  overlay: () => HTMLElement | null,
  toggle: () => HTMLElement | null,
  close: () => void
): (e: KeyboardEvent) => void {
  return (e) => {
    if (e.key !== 'Escape' || e.defaultPrevented || e.isComposing) return;
    const el = overlay();
    if (!el) return;
    e.preventDefault();
    const hadFocus = el.contains(document.activeElement);
    close();
    if (hadFocus) toggle()?.focus();
  };
}

/**
 * Close the side panel while it opens over the slide (a narrow window, shell.css) on Escape, as
 * the overlay rail does: focus goes back to the panel toggle when it was in the panel. A wide
 * window keeps the panel in the row, so Escape leaves it open. Returns the cleanup.
 */
export function dismissPanelOverlay(options: {
  /** The demo's root element (`.demo-root`), which holds the panel and its toggle. */
  readonly root: () => HTMLElement | null;
  readonly close: () => void;
}): () => void {
  if (typeof document === 'undefined') return () => {};
  const overlay = (): HTMLElement | null => {
    const el = options.root()?.querySelector<HTMLElement>('.pptx-panel') ?? null;
    return el && getComputedStyle(el).position === 'absolute' ? el : null;
  };
  const toggle = (): HTMLElement | null =>
    options.root()?.querySelector<HTMLElement>('.demo-panel-toggle') ?? null;
  const onKey = closeOnEscape(overlay, toggle, options.close);
  document.addEventListener('keydown', onKey);
  return () => document.removeEventListener('keydown', onKey);
}

/**
 * Close the slide rail while it opens over the slide (a narrow window, shell.css), as a
 * dismissible overlay does: on Escape, on a press outside it and its toggle, and after a click
 * picks a slide. Escape gives focus back to the toggle when focus was in the rail; a picked
 * slide gives it to the slide. A wide window keeps the rail in the row, so nothing closes it.
 * Returns the cleanup.
 */
export function dismissRailOverlay(options: {
  /** The demo's root element (`.demo-root`), which holds the rail and its toggle. */
  readonly root: () => HTMLElement | null;
  readonly editor: () => PptxEditorInstance | null;
  readonly close: () => void;
}): () => void {
  if (typeof document === 'undefined') return () => {};
  const rail = (): HTMLElement | null =>
    options.root()?.querySelector<HTMLElement>('.pptx-rail:not(.pptx-rail--collapsed)') ?? null;
  const toggle = (): HTMLElement | null =>
    options.root()?.querySelector<HTMLElement>('.demo-rail-toggle') ?? null;
  // Only the overlay closes: shell.css positions the rail over the slide on a narrow window.
  const overlay = (): HTMLElement | null => {
    const el = rail();
    return el && getComputedStyle(el).position === 'absolute' ? el : null;
  };
  const inside = (el: HTMLElement, target: EventTarget | null): boolean =>
    target instanceof Node && el.contains(target);
  const onKey = closeOnEscape(overlay, toggle, options.close);
  const onPress = (e: PointerEvent): void => {
    const el = overlay();
    if (!el || inside(el, e.target) || inside(toggle() ?? el, e.target)) return;
    // A press in a menu that the rail opened (its context menu) keeps the rail.
    if (e.target instanceof Element && e.target.closest('.pptx-menu')) return;
    options.close();
  };
  const onClick = (e: MouseEvent): void => {
    const el = overlay();
    if (!el || !(e.target instanceof Element)) return;
    // A plain click picks one slide; Shift and Ctrl/Cmd clicks build a selection and stay.
    if (e.shiftKey || e.ctrlKey || e.metaKey || e.button !== 0) return;
    const item = e.target.closest('.pptx-rail__item');
    if (!item || !el.contains(item)) return;
    options.close();
    options.editor()?.focus();
  };
  document.addEventListener('keydown', onKey);
  document.addEventListener('pointerdown', onPress, true);
  document.addEventListener('click', onClick);
  return () => {
    document.removeEventListener('keydown', onKey);
    document.removeEventListener('pointerdown', onPress, true);
    document.removeEventListener('click', onClick);
  };
}

/**
 * The page title the demo starts with (the `<title>` of its `index.html`): the deck name goes
 * before it. It is read once per page load and kept on the document, because a hot reload
 * evaluates this module again after `document.title` already holds a deck name.
 */
function pageTitle(): string {
  const root = document.documentElement;
  root.dataset.pageTitle ??= document.title;
  return root.dataset.pageTitle;
}

/** Show the deck name in the browser tab: "Sample presentation - PPTX Editor". */
export function showDeckTitle(deck: string): void {
  if (typeof document === 'undefined') return;
  const base = pageTitle();
  const name = deck.trim();
  document.title = name ? `${name} - ${base}` : base;
}

/**
 * Whether a drag carries files. The file-drop hint and the drop act only on a file drag, so a
 * slide dragged in the rail or the sorter reaches its own drop target.
 */
export function carriesFiles(data: DataTransfer | null | undefined): boolean {
  return data != null && Array.from(data.types).includes('Files');
}

/**
 * Whether a `dragleave` left the drop zone itself. A drag that moves from the zone onto one of
 * its children also fires `dragleave`, and the drop hint must not flicker then.
 */
export function leftDropZone(zone: EventTarget | null, next: EventTarget | null): boolean {
  return !(zone instanceof Node && next instanceof Node && zone.contains(next));
}

/** The unsaved edits of the open deck: what the demo asks about before it replaces the deck. */
export interface UnsavedEdits {
  /** Whether the open deck has edits that were not exported since it opened. */
  dirty(): boolean;
  /** The deck was exported: its edits so far are saved. */
  markSaved(): void;
  /** Stop following the editor, and remove the page-leave guard. */
  dispose(): void;
}

/**
 * Follow the unsaved edits of `editor`'s deck: a new deck starts clean, and an export marks the
 * edits saved. The deck is compared by content with the deck as it was opened or last saved
 * (`canonicalPackageFingerprint`), so undo back to that state is clean again; the revision
 * cannot tell, because it goes up on every commit, undo and redo too. While edits are unsaved,
 * closing or reloading the page asks first (`beforeunload`).
 */
export function trackUnsavedEdits(editor: PptxEditorInstance): UnsavedEdits {
  const current = (): PptxPackage | null => editor.session?.store.package ?? null;
  let session = editor.session;
  let saved = current();
  let savedPrint: string | null = null;
  let last: { pkg: PptxPackage; print: string } | null = null;
  const print = (pkg: PptxPackage): string => {
    if (last?.pkg === pkg) return last.print;
    last = { pkg, print: canonicalPackageFingerprint(pkg) };
    return last.print;
  };
  const mark = (): void => {
    saved = current();
    savedPrint = null;
  };
  const follow = (): void => {
    if (editor.session === session) return;
    session = editor.session;
    mark();
  };
  const off = editor.subscribe(follow);
  const dirty = (): boolean => {
    follow();
    const pkg = current();
    if (!pkg || !saved || pkg === saved) return false;
    // The fingerprint is taken only when asked: before the page closes or another deck opens.
    savedPrint ??= print(saved);
    return print(pkg) !== savedPrint;
  };
  const onLeave = (e: BeforeUnloadEvent): void => {
    if (!dirty()) return;
    e.preventDefault();
    // Some browsers show the prompt only for a return value.
    e.returnValue = '';
  };
  const win = typeof window === 'undefined' ? null : window;
  win?.addEventListener('beforeunload', onLeave);
  return {
    dirty,
    markSaved() {
      follow();
      mark();
    },
    dispose() {
      off();
      win?.removeEventListener('beforeunload', onLeave);
    },
  };
}

/**
 * Open `bytes` in `editor` in place of its deck. Resolves true once the deck opened, so the demo
 * renames the deck only then: a refused file keeps the open deck and its name, and the editor's
 * open-failure notice says why.
 */
export async function openDeck(
  editor: PptxEditorInstance | null | undefined,
  bytes: Uint8Array
): Promise<boolean> {
  return !!editor && (await editor.load(bytes)).ok;
}

/** Say why a dropped file does not open: only `.pptx` files do. */
export function refuseDroppedFile(editor: PptxEditorInstance | null | undefined, file: File): void {
  editor?.reject('app.dropRefused', file.name);
}

/**
 * The other demo's address with this page's query (`?theme=dark&chrome=ribbon&deck=…`), so the
 * React | Vue switcher keeps the theme, the chrome, and the deck.
 */
export function adapterHref(base: string, search: string): string {
  return search && search !== '?' ? `${base}${search.startsWith('?') ? '' : '?'}${search}` : base;
}
