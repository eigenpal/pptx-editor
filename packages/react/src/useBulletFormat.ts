// The bullet format as a hook, the twin of `useFontFormat`: the read side off the snapshot, the
// write side through the engine's `setBulletFormat` command. The packaged Bullets and Numbering
// dialog and any host chrome read the same `BulletFormatRead`, so they cannot disagree.

import { useCallback } from 'react';
import {
  sameBulletFormat,
  type BulletFormat,
  type BulletFormatRead,
} from '@pptx-editor.dev/core/editor';
import { useEditorCommand, useEditorState, usePptxEditor } from './context.ts';

/** What `useBulletFormat` returns. @public */
export interface UseBulletFormatReturn {
  /**
   * The bullet format of the selected paragraphs: null when they have no list, undefined when
   * they differ or nothing with text is selected.
   */
  readonly format: BulletFormatRead | null | undefined;
  /** Whether the engine can write a bullet format now (text or a shape with text selected, editable). */
  readonly isEnabled: boolean;
  /**
   * Write a bullet format (null turns the list off) as ONE undo step. An omitted field keeps
   * each paragraph's value; with no `kind`, the size and color change and each paragraph keeps
   * its own bullet or number. Returns whether the engine accepted it.
   */
  readonly apply: (format: BulletFormat | null) => boolean;
}

/**
 * The bullet format of the selected paragraphs, and the command to change it. `apply` sends one
 * `setBulletFormat`, so a dialog's worth of changes is one undo step.
 *
 * @example
 * ```tsx
 * function Bigger() {
 *   const { isEnabled, apply } = useBulletFormat();
 *   return <button disabled={!isEnabled} onClick={() => apply({ sizePercent: 150 })}>Bigger bullets</button>;
 * }
 * ```
 *
 * @public
 */
export function useBulletFormat(): UseBulletFormatReturn {
  const editor = usePptxEditor();
  const format = useEditorState((s) => s.format?.bulletFormat, sameBulletFormat);
  const { isEnabled } = useEditorCommand('list.bulletStyle');
  const apply = useCallback(
    (next: BulletFormat | null): boolean =>
      editor !== null && editor.exec({ type: 'setBulletFormat', format: next }).ok,
    [editor]
  );
  return { format, isEnabled, apply };
}
