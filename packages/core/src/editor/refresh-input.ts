// The user input that a document refresh must not overwrite: an IME composition, a pointer
// gesture on the slide, and speaker notes text that the notes pane did not commit yet.
//
// A result that arrives during such input waits for it, or commits it, and then finds the
// revision changed and refuses with `local-edits`. Without this, a reload replaces the text of
// a running composition, cancels a shape drag, or applies over a notes draft that the pane then
// writes over the returned notes.
//
// The tracker listens on the editor's document in the capture phase, so it sees each press and
// composition before the surface or a pane handles it. It is DOM state only: it changes nothing.
// A press counts as a gesture on the slide when it goes down in a `.pptx-surface` element of
// the editor. The tracker records that element at the press: the overlay draws its handles
// again on each move, so the pressed handle leaves the page during a drag, but the surface
// stays.
//
// `cancel()` ends each wait at once (`wakeRefreshWaits`): the waiting check then reads the
// cancellation and stops.
//
// To commit a notes draft, `commitPendingInput` sends one `pointerdown` event to the document
// itself, with pointer id `COMMIT_POINTER_ID`, which no real pointer uses. It does so only when
// the focused notes field holds text that differs from the committed notes, as `setNotes`
// normalizes it. The event has no target element, so no element of the page handles it; the
// listeners on the document see it, the pane's commit listener among them. The tracker
// does not record it as a press.

import { cleanText } from '../automation/commands/text-units.ts';
import { notesFieldText } from './notes-dom.ts';
import type { PptxEditorInstance } from './pptx-editor-types.ts';
import { COMPOSING_CLASS } from './surface-input.ts';
import { editorHostDocument, onEditorHostChange, surfaceRootOf } from './surface-registry.ts';

/** How long a refresh waits for input to end. Tests set shorter times. @internal */
export const REFRESH_INPUT_LIMITS = {
  /** The longest wait for an IME composition to end. */
  compositionMs: 5000,
  /** The longest wait for a pointer gesture to end (a lost `pointerup` ends it). */
  gestureMs: 30000,
};

/** The pointer id of the press that commits a notes draft: no real pointer has it. @internal */
export const COMMIT_POINTER_ID = 2147483647;

interface Tracker {
  /** The surface element under the press of each pointer that is down, or null. */
  readonly pressed: Map<number, Element | null>;
  /** The element an IME composes in, or null. */
  composing: Element | null;
  /** Called after each pointer release and composition end. */
  readonly waiters: Set<() => void>;
}

const trackers = new WeakMap<Document, Tracker>();

/** The `.pptx-surface` element that contains `target`, or null. */
function surfaceAt(target: EventTarget | null): Element | null {
  const node = target as Node | null;
  if (!node || typeof node.nodeType !== 'number') return null;
  const el = node.nodeType === 1 ? (node as Element) : node.parentElement;
  return el?.closest?.('.pptx-surface') ?? null;
}

/** The tracker of `doc`, made with its listeners on first use. */
function trackerOf(doc: Document): Tracker {
  const known = trackers.get(doc);
  if (known) return known;
  const t: Tracker = { pressed: new Map(), composing: null, waiters: new Set() };
  trackers.set(doc, t);
  const wake = () => {
    // The next task: the surface ends its own gesture in a later listener of the same event.
    setTimeout(() => {
      for (const waiter of [...t.waiters]) waiter();
    }, 0);
  };
  const up = (e: Event) => {
    t.pressed.delete((e as PointerEvent).pointerId ?? 0);
    wake();
  };
  const id = (e: Event) => (e as PointerEvent).pointerId ?? 0;
  doc.addEventListener(
    'pointerdown',
    (e) => {
      if (id(e) !== COMMIT_POINTER_ID) t.pressed.set(id(e), surfaceAt(e.target));
    },
    true
  );
  doc.addEventListener('pointerup', up, true);
  doc.addEventListener('pointercancel', up, true);
  // A release the page never saw (outside the window): a move with no button ends the press.
  doc.addEventListener(
    'pointermove',
    (e) => {
      if ((e as PointerEvent).buttons === 0 && t.pressed.delete(id(e))) wake();
    },
    true
  );
  doc.addEventListener('compositionstart', (e) => (t.composing = e.target as Element), true);
  doc.addEventListener(
    'compositionend',
    () => {
      t.composing = null;
      wake();
    },
    true
  );
  return t;
}

/** Track input in the document of `editor`, now and after the slide moves to another one. */
export function trackRefreshInput(editor: object): void {
  const install = () => {
    const doc = editorHostDocument(editor);
    if (doc) trackerOf(doc);
  };
  install();
  onEditorHostChange(editor, install);
}

/** The editor's root element: the adapter's `.pptx-editor`, else the slide's own element. */
function editorRoot(editor: object): HTMLElement | null {
  const root = surfaceRootOf(editor);
  return (root?.closest('.pptx-editor') as HTMLElement | null) ?? root;
}

const inside = (root: HTMLElement | null, target: EventTarget | null | undefined): boolean =>
  !!root &&
  !!target &&
  typeof (target as Node).nodeType === 'number' &&
  root.contains(target as Node);

/** Whether an IME composes in the editor: in the slide's text input or in a pane field. */
export function composing(editor: object): boolean {
  const root = surfaceRootOf(editor);
  if (root?.querySelector('.pptx-surface__input')?.classList.contains(COMPOSING_CLASS)) return true;
  const doc = root?.ownerDocument;
  const t = doc ? trackers.get(doc) : undefined;
  return !!t?.composing && inside(editorRoot(editor), t.composing);
}

/** Whether a pointer that went down on the slide is still down. */
export function gestureActive(editor: object): boolean {
  const root = surfaceRootOf(editor);
  const t = root ? trackers.get(root.ownerDocument) : undefined;
  if (!t) return false;
  for (const surface of t.pressed.values()) if (inside(root, surface)) return true;
  return false;
}

/** Run each wait's check now: `cancel()` calls it, so a cancelled wait ends at once. */
export function wakeRefreshWaits(editor: object): void {
  const doc = editorHostDocument(editor);
  const t = doc ? trackers.get(doc) : undefined;
  for (const waiter of [...(t?.waiters ?? [])]) waiter();
}

/**
 * Wait until `busy()` is false, at most `ms`. Resolves true when it ended, false at the limit.
 * Checks after each pointer release and composition end.
 */
export function waitUntil(editor: object, busy: () => boolean, ms: number): Promise<boolean> {
  if (!busy()) return Promise.resolve(true);
  const doc = editorHostDocument(editor);
  const t = doc ? trackerOf(doc) : null;
  return new Promise((resolve) => {
    let done = false;
    const end = (ok: boolean) => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      t?.waiters.delete(check);
      resolve(ok);
    };
    const check = () => {
      if (!busy()) end(true);
    };
    const timer = setTimeout(() => end(!busy()), ms);
    t?.waiters.add(check);
    // The slide's input reports a composition end on itself; a destroyed input ends nothing.
    const input = surfaceRootOf(editor)?.querySelector('.pptx-surface__input');
    input?.addEventListener('compositionend', () => setTimeout(check, 0), { once: true });
  });
}

/**
 * Commit speaker notes text that the notes pane holds as a draft. The pane commits a draft
 * before a press anywhere else acts (a `pointerdown` listener on its document, in React and in
 * Vue); this sends that press to the document itself, so no element of the page handles it.
 * Nothing happens when the focused notes field shows the committed notes, after the
 * normalization of `setNotes`.
 */
export function commitPendingInput(editor: PptxEditorInstance): void {
  const root = editorRoot(editor);
  const doc = root?.ownerDocument;
  const field = doc?.activeElement as HTMLElement | null | undefined;
  if (!doc || !field || !field.classList?.contains('pptx-notes__text') || !root!.contains(field))
    return;
  const text = cleanText(notesFieldText(field)).replace(/\u000B/g, '\n');
  if (text === editor.snapshot().notes) return;
  const win = doc.defaultView;
  const init = { pointerId: COMMIT_POINTER_ID, isPrimary: false };
  const press = win?.PointerEvent
    ? new win.PointerEvent('pointerdown', init)
    : Object.assign(new (win?.Event ?? Event)('pointerdown'), init);
  doc.dispatchEvent(press);
}
