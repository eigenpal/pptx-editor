// The Font group's controls of the chrome registry: Character Spacing, Change Case, the Font
// dialog slot, and the `script` group (superscript and subscript). Split from
// `chrome-controls.ts`, which places them in `CHROME_GROUPS`; type-only imports keep the
// modules free of an import cycle.

import type { ChromeControl, ChromeGroup } from './chrome-controls.ts';
import { ICON_PATHS } from './icon-paths.ts';
import { TEXT_CASES } from '../automation/commands/font-format.ts';

const icon = (name: string): readonly string[] => ICON_PATHS[name] ?? [];

/**
 * Home > Character Spacing, in points (`spc` / 100): Very Tight, Tight, Normal, Loose, and Very
 * Loose.
 */
const CHARACTER_SPACINGS = [
  ['veryTight', -3],
  ['tight', -1.5],
  ['normal', 0],
  ['loose', 3],
  ['veryLoose', 6],
] as const;

/** `text.characterSpacing`. @internal */
export const CHARACTER_SPACING_CONTROL: ChromeControl = {
  // Home > Character Spacing: the five presets, and More Spacing…, which
  // opens the Font dialog on its Character Spacing tab. Off the default bar: the Font dialog
  // sets it too.
  id: 'characterSpacing',
  defaultToolbar: false,
  shape: 'dropdown',
  labelKey: 'toolbar.characterSpacing',
  toggle: 'choice',
  paths: icon('format_letter_spacing'),
  options: [
    ...CHARACTER_SPACINGS.map(([name, points]) => ({
      value: String(points),
      labelKey: `toolbar.characterSpacings.${name}` as const,
    })),
    // The packaged toolbar opens it on the Character Spacing tab.
    { value: 'dialog', labelKey: 'toolbar.characterSpacingMore', dialog: 'font' },
  ],
  state: { kind: 'value' },
};

/** `text.changeCase`. @internal */
export const CHANGE_CASE_CONTROL: ChromeControl = {
  // Home > Change Case: rewrites the letters, keeps the formatting.
  id: 'changeCase',
  shape: 'dropdown',
  labelKey: 'toolbar.changeCase',
  paths: icon('match_case'),
  options: TEXT_CASES.map((c) => ({ value: c, labelKey: `toolbar.cases.${c}` as const })),
  state: { kind: 'value' },
};

/** `text.dialog`. @internal */
export const FONT_DIALOG_CONTROL: ChromeControl = {
  // The Font dialog: effects, offset, character spacing, and kerning. Off the default bar, as the
  // Paragraph dialog: Format > Font…, the text context menu, and More Spacing… open it.
  id: 'dialog',
  labelKey: 'toolbar.fontDialog',
  paths: icon('text_format'),
  defaultToolbar: false,
  state: { kind: 'dialog', dialog: 'font' },
};

/** The `script` group. @internal */
export const SCRIPT_GROUP: ChromeGroup = {
  // Superscript and subscript, DOCX Editor's `script` group. Off the default bar: the Font
  // dialog and the keyboard shortcuts set them.
  id: 'script',
  labelKey: 'toolbar.groups.script',
  controls: [
    {
      id: 'super',
      defaultToolbar: false,
      labelKey: 'toolbar.superscript',
      toggle: 'pressed',
      paths: icon('superscript'),
      state: { kind: 'command' },
    },
    {
      id: 'sub',
      defaultToolbar: false,
      labelKey: 'toolbar.subscript',
      toggle: 'pressed',
      paths: icon('subscript'),
      state: { kind: 'command' },
    },
  ],
};
