// A demo app in happy-dom, driven the way a person drives it in a browser: clicks, keys,
// pointer drags, file picks, and drops. The React and Vue demo tests (`demo-*.test.tsx` in
// `packages/react`, `demo-*.test.ts` in `packages/vue`) mount their app through a `DemoHost`
// and assert on the store, the snapshot, and the painted SVG. They replace the Playwright demo
// specs; only what needs a real browser's layout or canvas stays in `e2e/`.
//
// The driver plays each event in a browser's order, with the browser's default actions: text
// input through `beforeinput`, Tab focus moves, Enter on a button, and the clipboard events of
// Ctrl/Cmd+C, X, and V. Hit-testing stays the engine's own (`surface-hit.ts`): a press at a
// slide point reaches the stage with that point's client coordinates (`demo-geometry.ts`).
// Layout uses the HarfBuzz measurer with the packaged substitute fonts, as the demo does.

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { installGeometry, restoreGeometry, slidePoint, stage } from './demo-geometry.ts';

export { IN, SLIDE_H, SLIDE_W, slidePoint, wordPoint } from './demo-geometry.ts';

/** The editor's stylesheet, then the demo shell's: the order of the demos' `styles.css`. */
const STYLESHEETS = [
  '../../packages/core/src/styles/editor.css',
  '../../examples/shared/shell.css',
];

// The demos read their base path (`import.meta.env.BASE_URL`, Vite's) when they load.
process.env.BASE_URL = '/';

const realFetch = globalThis.fetch;
const realCreateObjectURL = URL.createObjectURL;
const realRevokeObjectURL = URL.revokeObjectURL;

/** A file the demo saved through its download link. */
export interface Download {
  readonly name: string;
  readonly blob: Blob;
}

let packagedAssets: URL | null = null;

/** The demo's fetches: decks from its `public/` folder, fonts from the font package. */
async function demoFetch(publicDir: URL, input: RequestInfo | URL): Promise<Response> {
  const url = String(input instanceof Request ? input.url : input);
  if (url.startsWith('file://')) {
    packagedAssets ??= new URL('./', url);
    return new Response(readFileSync(fileURLToPath(url)));
  }
  const name = url.split('/').pop() ?? '';
  if (name.endsWith('.pptx')) return new Response(readFileSync(new URL(name, publicDir)));
  // The pinned catalog (a CDN in a browser): the same faces from the font package, offline.
  if (packagedAssets && name.endsWith('.ttf')) {
    try {
      return new Response(readFileSync(fileURLToPath(new URL(name, packagedAssets))));
    } catch {
      return new Response(null, { status: 404 });
    }
  }
  return new Response(null, { status: 404 });
}

/** Everything a test needs from the mounted demo. */
export interface Demo {
  readonly editor: PptxEditorInstance;
  /** `pageerror` and `console.error` messages, as the e2e specs collected them. */
  readonly errors: string[];
  readonly downloads: Download[];
}

/** The keys `ControlOrMeta`, `Shift`, and `Alt` in a Playwright-style chord. */
interface Mods {
  ctrlKey: boolean;
  shiftKey: boolean;
  altKey: boolean;
  metaKey: boolean;
}

const NO_MODS: Mods = { ctrlKey: false, shiftKey: false, altKey: false, metaKey: false };

/** How a test file mounts its demo: the React or the Vue app, and how its updates flush. */
export interface DemoHost {
  /** The demo's `public/` folder, which serves its decks. */
  readonly publicDir: URL;
  /** Mount the demo app into the document. */
  mount(): void | Promise<void>;
  /** Unmount it. */
  unmount(): void;
  /** Run a step and flush the framework's updates (React's `act`). */
  run(step: () => void | Promise<void>): Promise<void>;
}

/** Options of `openDemo`. */
export interface DemoOptions {
  readonly width?: number;
  readonly height?: number;
  readonly query?: string;
  /** Load the demo's stylesheets, so their breakpoints apply at `width`. */
  readonly css?: boolean;
  /** Measure the toolbar this many pixels wide (see `installGeometry`). */
  readonly barWidth?: number;
}

let host: DemoHost | null = null;
/** Puts back the browser stubs of the open demo. */
let restoreStubs: (() => void) | null = null;

function run(step: () => void | Promise<void>): Promise<void> {
  return host ? host.run(step) : Promise.resolve(step()).then(() => undefined);
}

/** Open a demo at a window size, on the basic deck, and wait until it paints. */
export async function openDemoWith(next: DemoHost, options: DemoOptions = {}): Promise<Demo> {
  host = next;
  const { width = 1440, height = 900, query = '?deck=basic', css = false } = options;
  // The page title of the demo's `index.html`; the deck name goes before it.
  document.title = 'PPTX Editor';
  delete document.documentElement.dataset.pageTitle;
  // The demo's stylesheets (the editor's, then the shell's), for the tests that read a
  // breakpoint: happy-dom matches their media queries against the window size.
  if (css)
    for (const sheet of STYLESHEETS) {
      const style = document.createElement('style');
      style.textContent = readFileSync(new URL(sheet, import.meta.url), 'utf8');
      document.head.append(style);
    }
  // A browser does not focus an element that its styles hide (HTML's focusing steps need a
  // rendered element); happy-dom does. With the stylesheets on, model the browser.
  const focus = HTMLElement.prototype.focus;
  if (css)
    HTMLElement.prototype.focus = function (this: HTMLElement, options?: FocusOptions): void {
      for (let n: Element | null = this; n; n = n.parentElement)
        if (getComputedStyle(n).display === 'none') return;
      focus.call(this, options);
    };
  const win = window as unknown as {
    happyDOM: { setURL(u: string): void; setViewport(v: { width: number; height: number }): void };
  };
  win.happyDOM.setViewport({ width, height });
  win.happyDOM.setURL(`http://localhost:5297/${query}`);
  globalThis.fetch = ((input: RequestInfo | URL) =>
    demoFetch(next.publicDir, input)) as typeof fetch;
  installGeometry(options.barWidth);
  const errors: string[] = [];
  const downloads: Download[] = [];
  const blobs = new Map<string, Blob>();
  let n = 0;
  URL.createObjectURL = (obj: Blob | MediaSource) => {
    const url = `blob:demo/${++n}`;
    if (obj instanceof Blob) blobs.set(url, obj);
    return url;
  };
  URL.revokeObjectURL = () => {};
  // The demo saves through a detached `<a download>` that it clicks: record the file.
  const anchorClick = HTMLAnchorElement.prototype.click;
  HTMLAnchorElement.prototype.click = function (this: HTMLAnchorElement): void {
    const blob = blobs.get(this.getAttribute('href') ?? '');
    if (this.hasAttribute('download') && blob) downloads.push({ name: this.download, blob });
    else anchorClick.call(this);
  };
  const onError = (e: ErrorEvent): void => void errors.push(e.message);
  window.addEventListener('error', onError);
  const consoleError = console.error;
  console.error = (...args: unknown[]) => {
    errors.push(args.map(String).join(' '));
  };
  // happy-dom has no FontFace API: the measured faces register with the page, as in a browser.
  const scope = globalThis as unknown as { FontFace?: unknown };
  const hadFontFace = scope.FontFace !== undefined;
  const hadFontSet = (document as unknown as { fonts?: unknown }).fonts !== undefined;
  if (!hadFontFace)
    scope.FontFace = class {
      load = () => Promise.resolve(this);
    };
  if (!hadFontSet)
    Object.defineProperty(document, 'fonts', {
      configurable: true,
      value: { add: () => undefined, delete: () => true, ready: Promise.resolve() },
    });
  const restore = (): void => {
    HTMLElement.prototype.focus = focus;
    if (!hadFontFace) delete scope.FontFace;
    if (!hadFontSet) Reflect.deleteProperty(document, 'fonts');
    console.error = consoleError;
    window.removeEventListener('error', onError);
    HTMLAnchorElement.prototype.click = anchorClick;
    globalThis.fetch = realFetch;
    URL.createObjectURL = realCreateObjectURL;
    URL.revokeObjectURL = realRevokeObjectURL;
    restoreGeometry();
  };
  restoreStubs = restore;
  await next.mount();
  await until(() => {
    if (document.querySelectorAll('.pptx-rail__item').length !== 5) throw new Error('no deck');
    if (!document.querySelector('.pptx-surface__slide')) throw new Error('no paint');
  }, 30_000);
  const editor = (window as unknown as { __pptxEditor: PptxEditorInstance }).__pptxEditor;
  return { editor, errors, downloads };
}

/** Unmount the demo and put the browser stubs back. Call it in `afterEach`. */
export function closeDemo(): void {
  host?.unmount();
  host = null;
  document.head.replaceChildren();
  restoreStubs?.();
  restoreStubs = null;
  document.body.replaceChildren();
  // The pointer, the held keys, and the clipboard start fresh for the next test.
  hovered = null;
  Object.assign(held, NO_MODS);
  clipboard = dataTransfer();
}

/** Let effects, microtasks, and animation frames run. */
export async function settle(ms = 0): Promise<void> {
  await run(async () => {
    // An animation frame is a macrotask in happy-dom: two turns run a frame and its effects.
    await new Promise((r) => setTimeout(r, ms));
    await new Promise((r) => setTimeout(r, 0));
  });
}

/** Poll `check` until it stops throwing, as Playwright's `expect.poll` does. */
export async function until(check: () => void, timeout = 5_000): Promise<void> {
  const end = Date.now() + timeout;
  for (;;) {
    try {
      check();
      return;
    } catch (error) {
      if (Date.now() > end) throw error;
    }
    await settle(5);
  }
}

function focusable(el: Element | null): HTMLElement | null {
  for (let n = el; n; n = n.parentElement) {
    if (!(n instanceof HTMLElement)) continue;
    if (n.matches('input:not([disabled]), textarea, select, button:not([disabled]), a[href]'))
      return n;
    if (n.isContentEditable || n.getAttribute('contenteditable') === 'plaintext-only') return n;
    if (n.hasAttribute('tabindex')) return n;
  }
  return null;
}

function pointer(type: string, target: Element, at: { x: number; y: number }, init = {}) {
  const event = new PointerEvent(type, {
    bubbles: true,
    cancelable: true,
    composed: true,
    clientX: at.x,
    clientY: at.y,
    pointerId: 1,
    pointerType: 'mouse',
    isPrimary: true,
    ...init,
  });
  target.dispatchEvent(event);
  return event;
}

function mouse(type: string, target: Element, at: { x: number; y: number }, init = {}) {
  const event = new MouseEvent(type, {
    bubbles: true,
    cancelable: true,
    composed: true,
    clientX: at.x,
    clientY: at.y,
    ...init,
  });
  target.dispatchEvent(event);
  return event;
}

const held: Mods = { ...NO_MODS };

/**
 * A press and release on `target` at a client point, with the browser's event order: pointer
 * events, the compatibility mouse events unless `pointerdown` was cancelled, focus on the
 * focusable target unless `mousedown` was cancelled, then `click` or `contextmenu`.
 */
async function pressOn(
  target: Element,
  at: { x: number; y: number },
  options: { button?: 'left' | 'right'; modifiers?: Partial<Mods> } = {}
): Promise<void> {
  const button = options.button === 'right' ? 2 : 0;
  const init = { ...held, ...options.modifiers, button, buttons: button === 2 ? 2 : 1 };
  await run(async () => {
    const down = pointer('pointerdown', target, at, init);
    if (!down.defaultPrevented) {
      const md = mouse('mousedown', target, at, init);
      if (!md.defaultPrevented) {
        const f = focusable(target);
        if (f) f.focus();
        else if (document.activeElement instanceof HTMLElement) document.activeElement.blur();
      }
    }
    if (button === 2) mouse('contextmenu', target, at, init);
    pointer('pointerup', target, at, { ...init, buttons: 0 });
    if (!down.defaultPrevented) mouse('mouseup', target, at, { ...init, buttons: 0 });
    if (button === 0) mouse('click', target, at, { ...init, buttons: 0, detail: 1 });
  });
  await settle();
}

function centerOf(el: Element): { x: number; y: number } {
  const r = el.getBoundingClientRect();
  return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
}

/** Click an element, as `locator.click()` does. */
export function click(
  el: Element,
  options: { button?: 'left' | 'right'; modifiers?: Partial<Mods> } = {}
): Promise<void> {
  return pressOn(el, centerOf(el), options);
}

/** Double-click an element. */
export async function dblclick(el: Element): Promise<void> {
  await click(el);
  await click(el);
  await run(async () => void mouse('dblclick', el, centerOf(el), { detail: 2 }));
  await settle();
}

/** Click a slide point (EMU) on the canvas. */
export function clickSlide(
  x: number,
  y: number,
  options: { button?: 'left' | 'right'; modifiers?: Partial<Mods> } = {}
): Promise<void> {
  return pressOn(stage(), slidePoint(x, y), options);
}

/** Double-click a slide point (EMU) on the canvas. */
export async function dblclickSlide(x: number, y: number): Promise<void> {
  const at = slidePoint(x, y);
  await pressOn(stage(), at);
  await pressOn(stage(), at);
  await run(async () => void mouse('dblclick', stage(), at, { detail: 2 }));
  await settle();
}

/**
 * A pointer drag in client points: press at `from` on `target` (the stage by default), move in
 * `steps`, then release. `during` runs while the button is still down.
 */
export async function drag(
  from: { x: number; y: number },
  to: { x: number; y: number },
  options: { steps?: number; target?: Element; during?: () => void | Promise<void> } = {}
): Promise<void> {
  const { steps = 5 } = options;
  const target = options.target ?? stage();
  await run(async () => {
    const down = pointer('pointerdown', target, from, { ...held, button: 0, buttons: 1 });
    if (!down.defaultPrevented) mouse('mousedown', target, from, { ...held, button: 0 });
    for (let i = 1; i <= steps; i += 1) {
      const at = {
        x: from.x + ((to.x - from.x) * i) / steps,
        y: from.y + ((to.y - from.y) * i) / steps,
      };
      pointer('pointermove', target, at, { ...held, button: -1, buttons: 1 });
    }
  });
  await settle();
  if (options.during) await options.during();
  await run(async () => {
    pointer('pointerup', target, to, { ...held, button: 0, buttons: 0 });
    mouse('click', target, to, { ...held, button: 0 });
  });
  await settle();
}

let hovered: Element | null = null;

/** Move the mouse over an element, as `locator.hover()` does. */
export async function hover(el: Element): Promise<void> {
  const at = centerOf(el);
  const from = hovered;
  hovered = el;
  await run(async () => {
    if (from && from !== el) {
      pointer('pointerout', from, at, { relatedTarget: el });
      mouse('mouseout', from, at, { relatedTarget: el });
    }
    pointer('pointerover', el, at, { relatedTarget: from });
    pointer('pointerenter', el, at, { relatedTarget: from, bubbles: false });
    mouse('mouseover', el, at, { relatedTarget: from });
    mouse('mouseenter', el, at, { relatedTarget: from, bubbles: false });
    pointer('pointermove', el, at, { buttons: 0 });
    mouse('mousemove', el, at, { buttons: 0 });
  });
  await settle();
}

/** Move the mouse to a client point over the element there (no button down). */
export async function moveTo(el: Element, at: { x: number; y: number }): Promise<void> {
  await run(async () => {
    pointer('pointermove', el, at, { buttons: 0 });
    mouse('mousemove', el, at, { buttons: 0 });
  });
  await settle();
}

/** Press at a client point on the canvas stage. */
export function clickAt(
  at: { x: number; y: number },
  options: { button?: 'left' | 'right'; modifiers?: Partial<Mods> } = {}
): Promise<void> {
  return pressOn(stage(), at, options);
}

/** Press at a client point on an element (one that covers the window, as the slide show). */
export function clickOn(el: Element, at: { x: number; y: number }): Promise<void> {
  return pressOn(el, at);
}

/** A repository file (`fixtures/…`) as a browser `File`. */
export function repoFile(path: string, name = path.split('/').pop()!): File {
  return new File([readFileSync(new URL(`../../${path}`, import.meta.url))], name);
}

/** Pick a file in the demo's Open input, as `setInputFiles` does; wait for the load. */
export async function openFile(file: File, input?: HTMLInputElement): Promise<void> {
  const el =
    input ?? document.querySelector<HTMLInputElement>('input[type=file][accept*=".pptx"]')!;
  await run(async () => {
    Object.defineProperty(el, 'files', { configurable: true, value: [file] });
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
  });
  // The read and the load are asynchronous; the status leaves `loading` when they are done.
  await settle(10);
  await until(() => {
    if (document.querySelector('.pptx-loading')) throw new Error('loading');
  });
  await settle();
}

/** Take the focus away from an element, as a click elsewhere does. */
export async function blur(el: HTMLElement): Promise<void> {
  await run(async () => el.blur());
  await settle();
}

/** Run a step that changes the demo (an editor call, a `focus()`), then let it settle. */
export async function step(change: () => unknown): Promise<void> {
  await run(async () => void (await change()));
  await settle();
}

/**
 * A browser `DataTransfer` as a page sees it: a file drag lists the type `Files` (happy-dom's
 * own lists each file's MIME type instead).
 */
export function dataTransfer(files: readonly File[] = []): DataTransfer {
  const data = new Map<string, string>();
  return {
    dropEffect: 'none',
    effectAllowed: 'all',
    files: files as unknown as FileList,
    items: [] as unknown as DataTransferItemList,
    get types() {
      return [...(files.length ? ['Files'] : []), ...data.keys()];
    },
    setData: (k: string, v: string) => void data.set(k, v),
    getData: (k: string) => data.get(k) ?? '',
    clearData: () => data.clear(),
    setDragImage: () => {},
  } as unknown as DataTransfer;
}

/** Dispatch a drag event with a data transfer; returns whether the page cancelled it. */
export function dragEvent(
  target: Element,
  type: string,
  data: DataTransfer,
  init: { clientX?: number; clientY?: number; relatedTarget?: EventTarget | null } = {}
): boolean {
  const event = new MouseEvent(type, { bubbles: true, cancelable: true, composed: true, ...init });
  Object.defineProperty(event, 'dataTransfer', { value: data });
  target.dispatchEvent(event);
  return event.defaultPrevented;
}

/**
 * A native drag of one rail item onto another, as the browser fires it for a mouse drag:
 * `dragstart` on the source, `dragenter` and `dragover` on the target near its top (`before`)
 * or bottom (`after`) edge, `drop`, then `dragend`. The drag carries no files.
 */
export async function dragItem(
  from: HTMLElement,
  to: HTMLElement,
  where: 'before' | 'after',
  during?: () => void
): Promise<void> {
  const data = dataTransfer();
  const box = to.getBoundingClientRect();
  const at = { clientX: box.x + 40, clientY: where === 'after' ? box.bottom - 4 : box.top + 4 };
  await run(async () => void dragEvent(from, 'dragstart', data, centerInit(from)));
  await run(async () => {
    dragEvent(to, 'dragenter', data, at);
    dragEvent(to, 'dragover', data, at);
  });
  during?.();
  await run(async () => {
    dragEvent(to, 'drop', data, at);
    dragEvent(from, 'dragend', data, at);
  });
  await settle();
}

function centerInit(el: Element): { clientX: number; clientY: number } {
  const c = centerOf(el);
  return { clientX: c.x, clientY: c.y };
}

/** Hold or release a modifier for the pointer events that follow (`keyboard.down('Alt')`). */
export function hold(key: 'Alt' | 'Shift', down: boolean): void {
  if (key === 'Alt') held.altKey = down;
  else held.shiftKey = down;
}

// ---------------------------------------------------------------------------------------------
// Keyboard

const NAMED: Readonly<Record<string, string>> = { Space: ' ' };

function parseChord(chord: string): { key: string; mods: Mods } {
  const parts = chord.split('+');
  // A chord that ends in `+` (the plus key) splits into an empty last part.
  let key = parts.pop() || '+';
  const mods: Mods = { ...NO_MODS };
  for (const p of parts) {
    if (p === 'ControlOrMeta' || p === 'Control') mods.ctrlKey = true;
    else if (p === 'Meta') mods.metaKey = true;
    else if (p === 'Shift') mods.shiftKey = true;
    else if (p === 'Alt') mods.altKey = true;
  }
  key = NAMED[key] ?? key;
  if (mods.shiftKey && key.length === 1) key = key.toUpperCase();
  return { key, mods };
}

const CODE: Readonly<Record<string, string>> = { ' ': 'Space' };

function codeOf(key: string): string {
  if (CODE[key]) return CODE[key];
  if (/^[a-z]$/i.test(key)) return `Key${key.toUpperCase()}`;
  if (/^[0-9]$/.test(key)) return `Digit${key}`;
  return key;
}

function active(): HTMLElement {
  return (document.activeElement as HTMLElement | null) ?? document.body;
}

function isTextField(el: Element): el is HTMLInputElement | HTMLTextAreaElement {
  return (
    el instanceof HTMLTextAreaElement ||
    (el instanceof HTMLInputElement &&
      /^(text|search|url|email|number|tel|password|)$/.test(el.type))
  );
}

/** Set a field's value through the native setter, so React sees the change. */
function setFieldValue(el: HTMLInputElement | HTMLTextAreaElement, value: string): void {
  const proto = el instanceof HTMLTextAreaElement ? HTMLTextAreaElement : HTMLInputElement;
  Object.getOwnPropertyDescriptor(proto.prototype, 'value')!.set!.call(el, value);
}

function beforeInput(el: Element, inputType: string, data: string | null = null): boolean {
  const e = new InputEvent('beforeinput', {
    bubbles: true,
    cancelable: true,
    composed: true,
    inputType,
    data,
  });
  el.dispatchEvent(e);
  return !e.defaultPrevented;
}

function inputEvent(el: Element, inputType: string, data: string | null = null): void {
  el.dispatchEvent(new InputEvent('input', { bubbles: true, composed: true, inputType, data }));
}

/** Insert text at a field's selection, as the browser's default action does. */
function insertInField(el: HTMLInputElement | HTMLTextAreaElement, text: string): void {
  const start = el.selectionStart ?? el.value.length;
  const end = el.selectionEnd ?? start;
  const next = el.value.slice(0, start) + text + el.value.slice(end);
  setFieldValue(el, next);
  el.setSelectionRange(start + text.length, start + text.length);
  inputEvent(el, text ? 'insertText' : 'deleteContentBackward', text || null);
}

/** Insert text into an editable element at its selection, or at its end. */
function insertInEditable(el: HTMLElement, text: string): void {
  const sel = window.getSelection();
  const range = sel && sel.rangeCount > 0 ? sel.getRangeAt(0) : null;
  if (range && el.contains(range.startContainer)) {
    range.deleteContents();
    const node = document.createTextNode(text);
    range.insertNode(node);
    range.setStartAfter(node);
    range.collapse(true);
    sel!.removeAllRanges();
    sel!.addRange(range);
  } else {
    el.append(document.createTextNode(text));
    caretToEnd(el);
  }
  inputEvent(el, 'insertText', text);
}

function caretToEnd(el: HTMLElement): void {
  const sel = window.getSelection();
  if (!sel) return;
  const range = document.createRange();
  range.selectNodeContents(el);
  range.collapse(false);
  sel.removeAllRanges();
  sel.addRange(range);
}

function tabbables(): HTMLElement[] {
  const all = [
    ...document.querySelectorAll<HTMLElement>(
      'a[href], button, input, select, textarea, [tabindex], [contenteditable]'
    ),
  ];
  return all.filter((el) => {
    if (el.matches(':disabled') || el.tabIndex < 0) return false;
    if (el instanceof HTMLInputElement && el.type === 'hidden') return false;
    for (let n: HTMLElement | null = el; n; n = n.parentElement)
      if (n.hidden || n.style.display === 'none') return false;
    return true;
  });
}

/** Move the focus to the next (or previous) element in tab order, as Tab does. */
function tabFocus(back: boolean): void {
  const list = tabbables();
  const at = list.indexOf(active());
  const next = list[(at + (back ? -1 : 1) + list.length) % list.length];
  next?.focus();
}

/** The system clipboard: what a `copy` or `cut` event put there, for the next `paste`. */
let clipboard = dataTransfer();

const CLIPBOARD_KEYS: Readonly<Record<string, 'copy' | 'cut' | 'paste'>> = {
  c: 'copy',
  x: 'cut',
  v: 'paste',
};

/** Fire a clipboard event at an element, as Ctrl/Cmd+C, X, or V do with clipboard access. */
function clipboardEvent(el: HTMLElement, type: 'copy' | 'cut' | 'paste'): void {
  if (type !== 'paste') clipboard = dataTransfer();
  const event = new Event(type, { bubbles: true, cancelable: true, composed: true });
  Object.defineProperty(event, 'clipboardData', { value: clipboard });
  el.dispatchEvent(event);
}

/** The browser's default action of a key that the page did not cancel. */
function defaultAction(el: HTMLElement, key: string, mods: Mods): void {
  const plain = !mods.ctrlKey && !mods.metaKey && !mods.altKey;
  const clip = (mods.ctrlKey || mods.metaKey) && !mods.altKey && CLIPBOARD_KEYS[key];
  if (clip) return clipboardEvent(el, clip);
  const surfaceInput = el.classList.contains('pptx-surface__input');
  if (key === 'Tab' && plain) return tabFocus(mods.shiftKey);
  if (surfaceInput) {
    if (!plain) return;
    if (key.length === 1) beforeInput(el, 'insertText', key);
    else if (key === 'Enter')
      beforeInput(el, mods.shiftKey ? 'insertLineBreak' : 'insertParagraph');
    else if (key === 'Backspace') beforeInput(el, 'deleteContentBackward');
    else if (key === 'Delete') beforeInput(el, 'deleteContentForward');
    return;
  }
  if (isTextField(el)) {
    if (plain && key.length === 1 && beforeInput(el, 'insertText', key)) insertInField(el, key);
    else if (plain && key === 'Backspace') {
      const start = el.selectionStart ?? el.value.length;
      el.setSelectionRange(Math.max(0, start - 1), el.selectionEnd ?? start);
      insertInField(el, '');
    }
    return;
  }
  if (el.isContentEditable || el.getAttribute('contenteditable') === 'plaintext-only') {
    if (mods.ctrlKey && key === 'End') return caretToEnd(el);
    if (plain && key.length === 1 && beforeInput(el, 'insertText', key)) insertInEditable(el, key);
    return;
  }
  if (plain && key === 'Enter' && el.matches('button, a[href], [role="menuitem"]')) el.click();
}

/** Press a key chord on the focused element, as `page.keyboard.press` does. */
export async function press(chord: string): Promise<void> {
  const { key, mods } = parseChord(chord);
  const init = { key, code: codeOf(key), bubbles: true, cancelable: true, composed: true, ...mods };
  await run(async () => {
    const el = active();
    const down = new KeyboardEvent('keydown', init);
    el.dispatchEvent(down);
    if (!down.defaultPrevented) defaultAction(active() === el ? el : active(), key, mods);
    const target = active();
    const up = new KeyboardEvent('keyup', init);
    target.dispatchEvent(up);
    if (!up.defaultPrevented && key === ' ' && target.matches('button')) target.click();
  });
  await settle();
}

/** Type text one key at a time, as `page.keyboard.type` does. */
export async function type(text: string): Promise<void> {
  for (const ch of text) await press(ch === ' ' ? 'Space' : ch);
}

/** Replace a field's text, as `locator.fill` does. */
export async function fill(el: HTMLElement, text: string): Promise<void> {
  await run(async () => {
    el.focus();
    if (isTextField(el)) {
      setFieldValue(el, text);
      inputEvent(el, 'insertText', text);
    } else {
      el.textContent = text;
      caretToEnd(el);
      inputEvent(el, 'insertText', text);
    }
  });
  await settle();
}

/** A toolbar button by name, on the bar or in its More menu. */
export async function barButton(name: string | RegExp, nth = 0): Promise<HTMLElement> {
  const bar = document.querySelector('[role="toolbar"][aria-label="Formatting"]')!;
  const onBar = [...bar.querySelectorAll<HTMLElement>('button')].filter((b) => matchName(b, name));
  const visible = onBar.filter((b) => !b.closest('.pptx-toolbar__overflow'));
  if (visible[nth]) return visible[nth];
  let overflow = document.querySelector('.pptx-toolbar__overflow');
  if (!overflow) {
    const more = [...bar.querySelectorAll<HTMLElement>('button')].find((x) => matchName(x, 'More'));
    await click(more!);
    overflow = document.querySelector('.pptx-toolbar__overflow');
  }
  const inMore = [...overflow!.querySelectorAll<HTMLElement>('button')].filter((b) =>
    matchName(b, name)
  );
  return inMore[nth]!;
}

function accessibleName(el: Element): string {
  return (el.getAttribute('aria-label') ?? el.textContent ?? '').trim();
}

function matchName(el: Element, name: string | RegExp): boolean {
  const n = accessibleName(el);
  return typeof name === 'string' ? n === name : name.test(n);
}

/** The painted `<text>` contents of the current slide. */
export function paintedTexts(): string[] {
  return [...document.querySelectorAll('.pptx-surface__slide text')].map(
    (t) => t.textContent ?? ''
  );
}

/** The first painted text element of the current slide whose content matches. */
export function paintedText(match: RegExp): SVGTextElement | undefined {
  return [...document.querySelectorAll<SVGTextElement>('.pptx-surface__slide text')].find((t) =>
    match.test(t.textContent ?? '')
  );
}

/** Whether a control is disabled, as Playwright's `toBeDisabled` reads it (`aria-disabled` too). */
export function isDisabled(el: Element): boolean {
  return (el as HTMLButtonElement).disabled === true || el.getAttribute('aria-disabled') === 'true';
}
