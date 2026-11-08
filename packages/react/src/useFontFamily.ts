// The font family picker's behavior as a hook, after DOCX Editor's `useFontFamily`
// (docx-editor/packages/react/src/editor/toolbar/FontFamily.tsx): the current value, the
// offerable families, the enabled state, and the setter, with no UI. A host that wants its own
// picker takes the hook and ignores the `Toolbar.FontFamily` parts. PPTX Editor adds the groups
// of core's `font-picker.ts` (the slide's theme fonts, the deck's fonts, all fonts), so a host
// list can show the same groups as the packaged picker. The Vue twin is
// `packages/vue/src/useFontFamily.ts`. DOCX Editor commit adf0e178.

import { useCallback, useMemo } from 'react';
import {
  fontPickerGroups,
  isCurrentFontRow,
  type FontPickerGroup,
  type FontPickerRow,
} from '@pptx-editor.dev/core/editor';
import { useEditorCommand, useEditorState, usePptxEditor } from './context.ts';

const NO_FAMILIES: readonly string[] = Object.freeze([]);

/** What `useFontFamily` returns, as DOCX Editor's `UseFontFamilyResult`. @public */
export interface UseFontFamilyReturn {
  /** The family of the selected text, or null for a mixed selection or no selection. */
  readonly value: string | null;
  /**
   * The theme font the selected text uses (`major` for the headings font, `minor` for the body
   * font), or null when it names a family.
   */
  readonly theme: 'major' | 'minor' | null;
  /**
   * Apply a family, or a theme font reference (`+mj-lt`, `+mn-lt`) so the text follows a later
   * theme change. A refusal does nothing.
   */
  readonly setValue: (family: string) => void;
  /** Every family the editor offers: standard, configured, embedded, and the deck's own. */
  readonly options: readonly string[];
  /**
   * The packaged picker's groups for the current slide, empty groups left out: the theme fonts
   * (labeled Headings and Body), the other families the deck names, then all families.
   */
  readonly groups: readonly FontPickerGroup[];
  /** Whether a row is the font the selected text uses, for a check mark. */
  readonly isCurrent: (row: FontPickerRow) => boolean;
  /** Whether the engine can change the font now. */
  readonly isEnabled: boolean;
  /** Why the font cannot change: a catalog key, or the engine's own text. */
  readonly disabledReason: string | null;
}

/**
 * The font family picker's behavior, with no UI, as DOCX Editor's `useFontFamily`.
 *
 * @example
 * ```tsx
 * function FamilyList() {
 *   const { groups, setValue, isCurrent, isEnabled } = useFontFamily();
 *   return groups.flatMap((group) =>
 *     group.rows.map((row) => (
 *       <button key={row.value} disabled={!isEnabled || row.disabled} aria-pressed={isCurrent(row)}
 *         onClick={() => setValue(row.value)}>{row.family}</button>
 *     ))
 *   );
 * }
 * ```
 *
 * @public
 */
export function useFontFamily(): UseFontFamilyReturn {
  const editor = usePptxEditor();
  const cmd = useEditorCommand('font.family');
  // The families change with the deck (embedded and host fonts) and with the slide's theme.
  const revision = useEditorState((s) => s.revision);
  const slide = useEditorState((s) => s.currentSlide);
  const theme = useEditorState((s) => s.format?.fontFamilyTheme ?? null);
  const value = typeof cmd.value === 'string' ? cmd.value : null;
  const choices = useMemo(
    () => ({
      groups: fontPickerGroups(editor, slide, ''),
      options: editor?.fontFamilies() ?? NO_FAMILIES,
    }),
    // `revision` is the change signal for the deck's fonts.
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [editor, slide, revision]
  );
  const { execute } = cmd;
  const setValue = useCallback((family: string) => void execute(family), [execute]);
  const isCurrent = useCallback(
    (row: FontPickerRow) => !row.disabled && isCurrentFontRow(row, value, theme),
    [value, theme]
  );
  return useMemo(
    () => ({
      value,
      theme,
      setValue,
      options: choices.options,
      groups: choices.groups,
      isCurrent,
      isEnabled: cmd.isEnabled,
      disabledReason: cmd.disabledReason,
    }),
    [value, theme, setValue, choices, isCurrent, cmd.isEnabled, cmd.disabledReason]
  );
}
