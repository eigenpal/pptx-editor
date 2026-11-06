// Typed editor events for host chrome, after DOCX Editor's `EditorEvents` and `useEditorEvent`.
//
// The editor has one notification path: `subscribe`, which fires after every published state
// change. The events here are not a second path. Each one compares the snapshot before and
// after a notification and calls the handler when its part of the state moved. The React and
// Vue `useEditorEvent` hooks both call `onEditorEvent`, so they report the same events.

import type { DocumentOpenFailure, EditorRejection, EditorSnapshot } from '../contracts/index.ts';

/** The events `useEditorEvent` subscribes to, and what each handler receives. @public */
export interface PptxEditorEvents {
  /**
   * The presentation changed: an applied edit, an undo or redo, or a load that finished or
   * failed. Selection, zoom, and view changes do not fire it, and neither does a slide of a
   * long deck that finishes opening in the background (`selectionChange` reports those): a
   * host that saves on `change` does not save, and parse the whole deck, while it opens.
   */
  change: (snapshot: EditorSnapshot) => void;
  /**
   * The shape or text selection, the current slide, or the selected slides changed; or, while a
   * long deck opens, the count of slides still opening (`EditorSnapshot.pendingSlides`), as
   * DOCX Editor reports pages laid out in the background.
   */
  selectionChange: (snapshot: EditorSnapshot) => void;
  /** The editor refused an edit (`EditorSnapshot.lastRejection`). */
  rejection: (rejection: EditorRejection, snapshot: EditorSnapshot) => void;
  /** A file did not open (`EditorSnapshot.openFailure`). */
  openFailure: (failure: DocumentOpenFailure, snapshot: EditorSnapshot) => void;
}

/** The part of the editor `onEditorEvent` reads. */
interface EventSource {
  snapshot(): EditorSnapshot;
  subscribe(listener: () => void): () => void;
}

const sameSlides = (a: readonly number[], b: readonly number[]): boolean =>
  a.length === b.length && a.every((n, i) => n === b[i]);

/** Whether the event `event` happened between snapshots `prev` and `next`. */
function happened(event: keyof PptxEditorEvents, prev: EditorSnapshot, next: EditorSnapshot) {
  switch (event) {
    case 'change':
      // A load that finishes (or fails) counts; the start of a load does not. A slide that
      // finishes opening in the background is no change: the document did not change.
      return (
        prev.revision !== next.revision ||
        (prev.status !== next.status && next.status !== 'loading')
      );
    case 'selectionChange':
      return (
        prev.selection !== next.selection ||
        prev.currentSlide !== next.currentSlide ||
        !sameSlides(prev.selectedSlides, next.selectedSlides) ||
        (prev.pendingSlides ?? 0) !== (next.pendingSlides ?? 0)
      );
    case 'rejection':
      return !!next.lastRejection && next.lastRejection !== prev.lastRejection;
    case 'openFailure':
      return !!next.openFailure && next.openFailure !== prev.openFailure;
  }
}

/**
 * Call `handler` each time `event` happens on `editor`. Returns the unsubscribe function.
 * The adapters' `useEditorEvent` hooks use it; a host without an adapter can call it too.
 *
 * @example
 * ```ts
 * const off = onEditorEvent(editor, 'rejection', (rejection) => console.warn(rejection.reason));
 * ```
 *
 * @public
 */
export function onEditorEvent<E extends keyof PptxEditorEvents>(
  editor: EventSource,
  event: E,
  handler: PptxEditorEvents[E]
): () => void {
  let prev = editor.snapshot();
  return editor.subscribe(() => {
    const next = editor.snapshot();
    const before = prev;
    prev = next;
    if (next === before || !happened(event, before, next)) return;
    const call = handler as (...args: unknown[]) => void;
    if (event === 'rejection') call(next.lastRejection, next);
    else if (event === 'openFailure') call(next.openFailure, next);
    else call(next);
  });
}
