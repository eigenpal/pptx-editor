// Paragraph formatting as a hook: the read side off the snapshot, the write side through the
// engine's `setParagraphFormat` command. The packaged Paragraph dialog and any host chrome read
// the same `ParagraphFormatRead`, so they cannot disagree. After DOCX Editor's
// `useParagraphFormat` (docx-editor/packages/react/src/editor/useParagraphFormat.ts).

import { useCallback } from 'react';
import {
  chromeCommandState,
  sameParagraphFormat,
  type ParagraphFormatCommand,
  type ParagraphFormatRead,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor } from './context.ts';

/** What `useParagraphFormat` returns. @public */
export interface UseParagraphFormatReturn {
  /** The selected paragraphs' formatting, or null when no text is selected. */
  readonly format: ParagraphFormatRead | null;
  /** Whether the engine can write paragraph formatting now (text or a shape with text selected, editable). */
  readonly isEnabled: boolean;
  /** Write the given fields as ONE undo step. Returns whether the engine accepted them. */
  readonly apply: (update: ParagraphFormatCommand) => boolean;
}

/**
 * The selection's paragraph formatting, and the command to change it. `apply` sends one
 * `setParagraphFormat`, so a dialog's worth of changes is one undo step.
 *
 * @example
 * ```tsx
 * function SpaceAfterButton() {
 *   const { isEnabled, apply } = useParagraphFormat();
 *   return <button disabled={!isEnabled} onClick={() => apply({ spaceAfterPt: 12 })}>12 pt after</button>;
 * }
 * ```
 *
 * @public
 */
export function useParagraphFormat(): UseParagraphFormatReturn {
  const editor = usePptxEditor();
  const format = useEditorState((s) => s.format?.paragraph ?? null, sameParagraphFormat);
  // The engine's answer, not the dialog slot's: `popups.paragraph: false` hides the dialog,
  // never the operation.
  const isEnabled = useEditorState(() => chromeCommandState(editor, 'paragraph.dialog').isEnabled);
  const apply = useCallback(
    (update: ParagraphFormatCommand): boolean =>
      editor !== null && editor.exec({ ...update, type: 'setParagraphFormat' }).ok,
    [editor]
  );
  return { format, isEnabled, apply };
}
