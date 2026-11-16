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

interface Tracker {
  /** The press target of each pointer that is down. */
  readonly pressed: Map<number, EventTarget | null>;
  /** The element an IME composes in, or null. */
  composing: Element | null;
  /** Called after each pointer release and composition end. */
  readonly waiters: Set<() => void>;
}

const trackers = new WeakMap<Document, Tracker>();

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
  doc.addEventListener('pointerdown', (e) => t.pressed.set(id(e), e.target), true);
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
  for (const target of t.pressed.values()) if (inside(root, target)) return true;
  return false;
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
 * Nothing happens when the focused notes field shows the committed notes.
 */
export function commitPendingInput(editor: PptxEditorInstance): void {
  const root = editorRoot(editor);
  const doc = root?.ownerDocument;
  const field = doc?.activeElement as HTMLElement | null | undefined;
  if (!doc || !field || !field.classList?.contains('pptx-notes__text') || !root!.contains(field))
    return;
  if (notesFieldText(field) === editor.snapshot().notes) return;
  const EventCtor = doc.defaultView?.Event ?? Event;
  doc.dispatchEvent(new EventCtor('pointerdown'));
}
