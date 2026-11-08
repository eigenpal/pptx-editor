// The font family picker's behavior as a composable: the Vue twin of the React adapter's
// `useFontFamily` (packages/react/src/useFontFamily.ts), after DOCX Editor's Vue `useFontFamily`
// (docx-editor/packages/vue/src/editor/toolbar/useFontFamily.ts): the current value, the offerable
// families, the enabled state, and the setter, with no UI. PPTX Editor adds the groups of core's
// `font-picker.ts` (the slide's theme fonts, the deck's fonts, all fonts). DOCX Editor commit
// adf0e178.

import { computed, type ComputedRef } from 'vue';
import {
  fontPickerGroups,
  isCurrentFontRow,
  type FontPickerGroup,
  type FontPickerRow,
} from '@pptx-editor.dev/core/editor';
import { useEditorCommand, useEditorState, usePptxEditor } from './composables.ts';

const NO_FAMILIES: readonly string[] = Object.freeze([]);

/** What `useFontFamily` returns, as React's `UseFontFamilyReturn`. @public */
export interface UseFontFamilyReturn {
  /** The family of the selected text, or null for a mixed selection or no selection. */
  readonly value: ComputedRef<string | null>;
  /**
   * The theme font the selected text uses (`major` for the headings font, `minor` for the body
   * font), or null when it names a family.
   */
  readonly theme: ComputedRef<'major' | 'minor' | null>;
  /**
   * Apply a family, or a theme font reference (`+mj-lt`, `+mn-lt`) so the text follows a later
   * theme change. A refusal does nothing.
   */
  readonly setValue: (family: string) => void;
  /** Every family the editor offers: standard, configured, embedded, and the deck's own. */
  readonly options: ComputedRef<readonly string[]>;
  /**
   * The packaged picker's groups for the current slide, empty groups left out: the theme fonts
   * (labeled Headings and Body), the other families the deck names, then all families.
   */
  readonly groups: ComputedRef<readonly FontPickerGroup[]>;
  /** Whether a row is the font the selected text uses, for a check mark. */
  readonly isCurrent: (row: FontPickerRow) => boolean;
  /** Whether the engine can change the font now. */
  readonly isEnabled: ComputedRef<boolean>;
  /** Why the font cannot change: a catalog key, or the engine's own text. */
  readonly disabledReason: ComputedRef<string | null>;
}

/**
 * The font family picker's behavior, with no UI, as React's `useFontFamily` and DOCX Editor's
 * Vue composable of the same name. @public
 */
export function useFontFamily(): UseFontFamilyReturn {
  const editorRef = usePptxEditor();
  const cmd = useEditorCommand('font.family');
  // The families change with the deck (embedded and host fonts) and with the slide's theme.
  const revision = useEditorState((s) => s.revision);
  const slide = useEditorState((s) => s.currentSlide);
  const theme = useEditorState((s) => s.format?.fontFamilyTheme ?? null);
  const value = computed(() => (typeof cmd.value.value === 'string' ? cmd.value.value : null));
  const choices = computed(() => {
    // `revision` is the change signal for the deck's fonts.
    void revision.value;
    return {
      groups: fontPickerGroups(editorRef.value, slide.value, ''),
      options: editorRef.value?.fontFamilies() ?? NO_FAMILIES,
    };
  });
  return {
    value,
    theme: computed(() => theme.value),
    setValue: (family) => void cmd.execute(family),
    options: computed(() => choices.value.options),
    groups: computed(() => choices.value.groups),
    isCurrent: (row) => !row.disabled && isCurrentFontRow(row, value.value, theme.value),
    isEnabled: cmd.isEnabled,
    disabledReason: cmd.disabledReason,
  };
}
