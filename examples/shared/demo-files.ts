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

/** Fetch a deck that the demo serves. */
export async function fetchDeck(base: string, file: string): Promise<Uint8Array> {
  const res = await fetch(`${base}${file}`);
  return new Uint8Array(await res.arrayBuffer());
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

/** The notice for an HTTP status of the PDF function. */
function failureFor(status: number): PdfExportFailure {
  if (status === 413) return new PdfExportFailure('app.pdfTooLarge');
  // Too many slides, too much paint, or a PDF past its size limit.
  if (status === 507) return new PdfExportFailure('app.pdfTooComplex');
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
      signal: AbortSignal.timeout(90_000),
    });
  } catch {
    // Offline, no server, or past the timeout.
    throw new PdfExportFailure('app.pdfExportFailed');
  }
  const type = response.headers.get('content-type') ?? '';
  if (response.ok && type.startsWith('application/pdf')) return response.blob();
  throw failureFor(response.status);
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
    else
      download(
        new Blob([exportSlideSvg(session, slide, options)], { type: 'image/svg+xml' }),
        `${name} ${slide + 1}.svg`
      );
  } catch (e) {
    console.error(e);
    if (kind === 'pdf')
      editor.reject(e instanceof PdfExportFailure ? e.key : 'app.pdfExportFailed');
  }
}

/**
 * Below this window width the side panel starts closed, so the slide keeps a usable zoom: at
 * 1024 px the 340 px panel leaves the slide near 30%. The panel toggle still opens it.
 */
export const NARROW_PANEL_QUERY = '(max-width: 1199px)';

const narrowQuery = (): MediaQueryList | null =>
  typeof window !== 'undefined' && typeof window.matchMedia === 'function'
    ? window.matchMedia(NARROW_PANEL_QUERY)
    : null;

/** Whether the side panel starts open: closed on a narrow window. */
export function panelStartsOpen(): boolean {
  return !narrowQuery()?.matches;
}

/**
 * Close the side panel when the window becomes narrow, and open it when the window becomes wide
 * again. Between those crossings the user's toggle stands. Returns the cleanup.
 */
export function followPanelWidth(setOpen: (open: boolean) => void): () => void {
  const query = narrowQuery();
  if (!query) return () => {};
  const changed = (event: MediaQueryListEvent): void => setOpen(!event.matches);
  query.addEventListener('change', changed);
  return () => query.removeEventListener('change', changed);
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
