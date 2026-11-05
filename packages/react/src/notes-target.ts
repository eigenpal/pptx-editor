// The speaker notes field that has focus, per editor. While it has focus, the toolbar's bold,
// italic, underline, and strikethrough format the notes text. The registry lives in core
// (`chrome-forms.ts`), so the React and Vue adapters read the same one.

import { useSyncExternalStore } from 'react';
import {
  NOTES_SLOTS,
  notesTargetOf,
  setNotesTarget,
  subscribeNotesTargets,
  type NotesTarget,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';

export { NOTES_SLOTS, setNotesTarget, type NotesTarget };

/** The focused notes field of `editor`, or null. */
export function useNotesTarget(editor: PptxEditorInstance | null): NotesTarget | null {
  return useSyncExternalStore(
    subscribeNotesTargets,
    () => notesTargetOf(editor),
    () => null
  );
}
