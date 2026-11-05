// Font formatting as a hook, the twin of `useParagraphFormat`: the read side off the snapshot,
// the write side through the engine's `setFontFormat` command. The packaged Font dialog and any
// host chrome read the same `FontFormatRead`, so they cannot disagree.

import { useCallback } from 'react';
import {
  chromeCommandState,
  sameFontFormat,
  type FontFormatCommand,
  type FontFormatRead,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor } from './context.ts';

/** What `useFontFormat` returns. @public */
export interface UseFontFormatReturn {
  /** The font of the selected text, or of every run of the selected shapes; null with none. */
  readonly format: FontFormatRead | null;
  /** Whether the engine can write font formatting now (text or a shape with text selected, editable). */
  readonly isEnabled: boolean;
  /** Write the given fields as ONE undo step. Returns whether the engine accepted them. */
  readonly apply: (update: FontFormatCommand) => boolean;
}

/**
 * The font of the selected text (or of every run of the selected shapes), and the command to
 * change it. `apply` sends one `setFontFormat`, so
 * a dialog's worth of changes is one undo step.
 *
 * @example
 * ```tsx
 * function Loose() {
 *   const { isEnabled, apply } = useFontFormat();
 *   return <button disabled={!isEnabled} onClick={() => apply({ characterSpacingPt: 3 })}>Loose</button>;
 * }
 * ```
 *
 * @public
 */
export function useFontFormat(): UseFontFormatReturn {
  const editor = usePptxEditor();
  const format = useEditorState((s) => s.format?.font ?? null, sameFontFormat);
  // The engine's answer, not the dialog slot's: `popups.font: false` hides the dialog, never
  // the operation.
  const isEnabled = useEditorState(() => chromeCommandState(editor, 'text.dialog').isEnabled);
  const apply = useCallback(
    (update: FontFormatCommand): boolean =>
      editor !== null && editor.exec({ ...update, type: 'setFontFormat' }).ok,
    [editor]
  );
  return { format, isEnabled, apply };
}
