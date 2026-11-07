// The open state of the speaker notes, which the notes pane draws and host chrome (a View menu
// row, a Notes button) can toggle. `PptxEditor.Root` holds it, so they agree without the host
// wiring them. A notes pane given `open` is controlled by the host instead, as DOCX Editor's
// `Navigation` is: it mirrors that value into the root, and a toggle from the hook goes to its
// `onOpenChange`.

import { createContext, useCallback, useContext, useMemo, useRef, type CSSProperties } from 'react';
import { warnOutsideRoot } from '@pptx-editor.dev/core/editor';

/** The shared pane state. @internal */
export interface PaneState {
  /** Undefined until a notes pane or the host sets it. */
  readonly notesOpen: boolean | undefined;
  setNotesOpen(open: boolean): void;
  /** The mounted notes pane's handler: a controlled pane sends a change to its host. */
  readonly notesRequest: { current: ((open: boolean) => void) | null };
}

/** Provided by `PptxEditor.Root`. @internal */
export const PaneStateContext = createContext<PaneState | null>(null);

/** What `useNotesPane` returns. @public */
export interface UseNotesPaneReturn {
  /** Whether the notes pane shows its text. */
  readonly open: boolean;
  /** Show or hide the notes text. A controlled pane gets the change through `onOpenChange`. */
  readonly setOpen: (open: boolean) => void;
  /** Show the notes text when hidden, and hide it when shown. */
  readonly toggle: () => void;
}

/**
 * The open state of the speaker notes that `PptxEditor.NotesPane` shows, after DOCX Editor's
 * `useNavigationPane`. Use it to build a notes toggle of your own. Outside `PptxEditor.Root` it
 * warns and does nothing. @public
 */
export function useNotesPane(): UseNotesPaneReturn {
  const state = useContext(PaneStateContext);
  if (!state) warnOutsideRoot('useNotesPane');
  const open = state?.notesOpen ?? true;
  // Stable functions, as DOCX Editor's `useNavigationPane`: a host effect that lists them runs
  // only when the root changes. `toggle` reads the value at call time.
  const latest = useRef(open);
  latest.current = open;
  const setNotesOpen = state?.setNotesOpen;
  const request = state?.notesRequest;
  const setOpen = useCallback(
    (next: boolean): void => {
      if (!setNotesOpen || !request) return;
      if (request.current) request.current(next);
      else {
        // Only an uncontrolled root takes the value at once; a host may refuse a pane's change.
        latest.current = next;
        setNotesOpen(next);
      }
    },
    // The setter and the request slot are stable for the root's life, so the functions keep
    // their identity when the notes open and close.
    [setNotesOpen, request]
  );
  const toggle = useCallback(() => setOpen(!latest.current), [setOpen]);
  return useMemo(() => ({ open, setOpen, toggle }), [open, setOpen, toggle]);
}

/** A pane's inline style: its width from `paneWidth`, then the host's own styles. @internal */
export function paneStyle(width: number | undefined, style: CSSProperties | undefined) {
  if (width === undefined) return style;
  return { width, ...style };
}
