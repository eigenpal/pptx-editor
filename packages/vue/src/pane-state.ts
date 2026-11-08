// The open state of the speaker notes, which the notes pane draws and host chrome (a View menu
// row, a Notes button) can toggle: the Vue twin of the React adapter's `pane-state.ts`.
// `PptxEditorRoot` provides it, so they agree without the host wiring them. A notes pane given
// `open` is controlled by the host instead, as DOCX Editor's `Navigation` is: it writes that
// value into the root, and a toggle from the composable goes to its `update:open` handler.

import {
  computed,
  inject,
  provide,
  shallowRef,
  type ComputedRef,
  type InjectionKey,
  type ShallowRef,
} from 'vue';
import { warnOutsideRoot } from '@pptx-editor.dev/core/editor';

/** The shared pane state. @internal */
export interface PaneState {
  /** Undefined until a notes pane or the host sets it. */
  readonly notesOpen: ShallowRef<boolean | undefined>;
  setNotesOpen(open: boolean): void;
  /** The mounted notes pane's handler: a controlled pane sends a change to its host. */
  readonly notesRequest: { current: ((open: boolean) => void) | null };
}

/** Provided by `PptxEditorRoot`. @internal */
export const paneStateKey: InjectionKey<PaneState> = Symbol('pptxPaneState');

/** Create and provide the root's pane state. @internal */
export function providePaneState(): PaneState {
  const notesOpen = shallowRef<boolean | undefined>(undefined);
  const state: PaneState = {
    notesOpen,
    setNotesOpen: (open) => {
      notesOpen.value = open;
    },
    notesRequest: { current: null },
  };
  provide(paneStateKey, state);
  return state;
}

/** The root's pane state, or null outside a root. @internal */
export function usePaneState(): PaneState | null {
  return inject(paneStateKey, null);
}

/** What `useNotesPane` returns. @public */
export interface UseNotesPaneReturn {
  /** Whether the notes pane shows its text. */
  readonly open: ComputedRef<boolean>;
  /** Show or hide the notes text. A controlled pane gets the change through its handler. */
  readonly setOpen: (open: boolean) => void;
  /** Show the notes text when hidden, and hide it when shown. */
  readonly toggle: () => void;
}

/**
 * The open state of the speaker notes that the notes pane shows, as React's `useNotesPane` and
 * DOCX Editor's `useNavigationPane`. Use it to build a notes toggle of your own. Outside
 * `PptxEditorRoot` it warns and does nothing. @public
 */
export function useNotesPane(): UseNotesPaneReturn {
  const state = usePaneState();
  if (!state) warnOutsideRoot('useNotesPane');
  const open = computed(() => state?.notesOpen.value ?? true);
  const setOpen = (next: boolean): void => {
    if (!state) return;
    if (state.notesRequest.current) state.notesRequest.current(next);
    // Only an uncontrolled root takes the value at once; a host may refuse a pane's change.
    else state.setNotesOpen(next);
  };
  // Reads the value at call time: the root's ref changes at once, so two toggles in one tick
  // show and hide again, as the React twin's do.
  const toggle = (): void => setOpen(!open.value);
  return { open, setOpen, toggle };
}
