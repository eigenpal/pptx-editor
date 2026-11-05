// The speaker notes field that has focus, per editor: the Vue twin of the React adapter's
// `notes-target.ts`. While it has focus, the toolbar's bold, italic, underline, and
// strikethrough format the notes text. The registry lives in core (`chrome-forms.ts`), so a
// notes field in either adapter reaches the same toolbar rule.

import {
  getCurrentScope,
  onScopeDispose,
  shallowRef,
  toValue,
  watch,
  type MaybeRefOrGetter,
  type ShallowRef,
} from 'vue';
import {
  NOTES_SLOTS,
  notesTargetOf,
  setNotesTarget,
  subscribeNotesTargets,
  type NotesTarget,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';

export { NOTES_SLOTS, setNotesTarget, type NotesTarget };

/** What `useNotesTarget` returns: the focused notes field, or null. @internal */
export type UseNotesTargetReturn = Readonly<ShallowRef<NotesTarget | null>>;

/** The focused notes field of `editor`, or null. @internal */
export function useNotesTarget(
  editor: MaybeRefOrGetter<PptxEditorInstance | null>
): UseNotesTargetReturn {
  const target = shallowRef<NotesTarget | null>(notesTargetOf(toValue(editor)));
  const read = (): void => {
    target.value = notesTargetOf(toValue(editor));
  };
  if (!getCurrentScope()) return target;
  watch(() => toValue(editor), read);
  onScopeDispose(subscribeNotesTargets(read));
  return target;
}
