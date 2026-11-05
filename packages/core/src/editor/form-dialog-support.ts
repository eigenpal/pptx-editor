// The parts of the Paragraph, Font, Bullets and Numbering, Columns, and Background dialogs
// that do not depend on a framework: the message a field problem shows, the
// message a refused command shows, the options of the lists, the status line of the Background
// dialog, and which section a dialog opens on. The React and Vue dialogs call the same
// functions, so they cannot drift. Every export is adapter glue, not a stable API: `@internal`.

import type { Translate, TranslationKey } from '@pptx-editor.dev/i18n';
import type {
  BulletFormatRead,
  CommandResult,
  GradientFill,
  SlideBackground,
} from '../contracts/index.ts';
import { translateKeyOrText } from './adapter-support.ts';
import { BULLET_GALLERY, NUMBER_GALLERY } from './bullet-format.ts';
import {
  BULLET_TEXT_FONT,
  CURRENT_STYLE,
  bulletFontChoices,
  currentStylePreview,
  customBullet,
  customBulletPreview,
  type BulletsDialogFields,
  type BulletsDialogTab,
  type BulletsFieldError,
} from './bullets-dialog-fields.ts';
import type { DialogError } from './chrome-forms.ts';
import type { ColumnsFieldError } from './columns-dialog-fields.ts';
import { sameBackground } from './dialog-fields.ts';
import type { FontFieldError } from './font-dialog-fields.ts';
import type { ParagraphFieldError } from './paragraph-dialog-fields.ts';
import { presetGradient } from './palette.ts';

// --- Messages ------------------------------------------------------------------------------

/**
 * The message of a command a form dialog sent: null when the document took it, else the
 * engine's reason as the user reads it (or "no editor" when there is none). @internal
 */
export function formCommandRejection(
  t: Translate,
  result: CommandResult | null | undefined
): string | null {
  if (result?.ok) return null;
  if (!result) return t('disabled.noEditor');
  const detail = result.detail ?? '';
  return translateKeyOrText(t, result.reason, { reason: detail, detail });
}

/** The Paragraph dialog's message for a field problem, on the field to mark. @internal */
export function paragraphDialogError(t: Translate, problem: ParagraphFieldError): DialogError {
  const field = t(`paragraphDialog.fields.${problem.field}`);
  return { message: t(problem.key, { field, ...problem.params }), field: problem.field };
}

/** The Font dialog's message for a field problem, on the field to mark. @internal */
export function fontDialogError(t: Translate, problem: FontFieldError): DialogError {
  const field = t(`fontDialog.fields.${problem.field}`);
  return { message: t(problem.key, { field, ...problem.params }), field: problem.field };
}

/** The Columns dialog's message for a field problem, on the field to mark. @internal */
export function columnsDialogError(t: Translate, problem: ColumnsFieldError): DialogError {
  const field = t(`columnsDialog.${problem.field === 'count' ? 'number' : 'spacing'}`);
  return { message: t(problem.key, { field, ...problem.params }), field: problem.field };
}

/**
 * The Bullets and Numbering dialog's message for a field problem. The tab that is shown
 * (`kind`) holds the field: the Numbered tab's size field has its own name. @internal
 */
export function bulletsDialogError(
  t: Translate,
  problem: BulletsFieldError,
  kind: BulletsDialogTab
): DialogError {
  const label: TranslationKey =
    problem.field === 'sizePercent'
      ? 'bulletsAndNumberingDialog.size'
      : problem.field === 'char'
        ? 'bulletsAndNumberingDialog.character'
        : 'bulletsAndNumberingDialog.startAt';
  const name =
    problem.field === 'sizePercent' && kind === 'number' ? 'numberSizePercent' : problem.field;
  return { message: t(problem.key, { field: t(label), ...problem.params }), field: name };
}

// --- Options -------------------------------------------------------------------------------

/** The Font dialog's font list: the draft's font first, then the editor's families. @internal */
export function fontDialogFamilies(current: string, families: readonly string[]): string[] {
  return [...new Set([current, ...families])].filter(Boolean);
}

/**
 * The style list of one tab of the Bullets and Numbering dialog: the gallery, then a character
 * or font of the user's own, or the paragraphs' own bullet or scheme, when the gallery has no
 * such item. `opened` is the selection's list format when the dialog opened. @internal
 */
export function bulletStyleOptions(
  t: Translate,
  kind: BulletsDialogTab,
  seed: BulletsDialogFields | null,
  draft: BulletsDialogFields | null,
  opened: BulletFormatRead | null | undefined
): (readonly [string, string])[] {
  const gallery = kind === 'number' ? NUMBER_GALLERY : BULLET_GALLERY;
  const prefix = kind === 'number' ? 'numberStyles' : 'bulletStyles';
  const options: (readonly [string, string])[] = gallery.map(
    ({ key }) => [key, t(`toolbar.${prefix}.${key}` as 'toolbar.bulletStyles.none')] as const
  );
  if (kind === 'bullet' && draft && seed && customBullet(seed, draft))
    options.push([
      CURRENT_STYLE,
      t('bulletsAndNumberingDialog.customChar', { char: customBulletPreview(draft) }),
    ]);
  else if (opened && opened.kind === kind && (seed?.[kind] ?? '') === CURRENT_STYLE)
    options.push([
      CURRENT_STYLE,
      t('bulletsAndNumberingDialog.custom', { char: currentStylePreview(opened) }),
    ]);
  return options;
}

/**
 * The Font field of the Bullets and Numbering dialog: the text font, the bullet fonts, then the
 * deck's and the host's families, in an order that does not change while the dialog is open.
 * @internal
 */
export function bulletFontOptions(
  t: Translate,
  seed: BulletsDialogFields | null,
  families: readonly string[]
): (readonly [string, string])[] {
  return [
    [BULLET_TEXT_FONT, t('bulletsAndNumberingDialog.textFont')],
    ...bulletFontChoices([seed?.font ?? ''], families).map((f) => [f, f] as const),
  ];
}

// --- Background ----------------------------------------------------------------------------

/** The base color of the Background dialog's custom gradient when the theme has none. @internal */
export const BACKGROUND_GRADIENT_FALLBACK = '#4472C4';

/**
 * A Background dialog gradient preset. Background presets are softer than shape presets: light
 * at the top, slightly darker below. @internal
 */
export function backgroundPresetGradient(base: string): GradientFill {
  return presetGradient(base, 0.5, 0.2);
}

/**
 * The status line of the Background dialog, which a screen reader announces: what the draft
 * holds, or null when a pressed swatch shows it. A gradient the user chose that no preset shows
 * says so, as a picture does; the opening fill is the slide's own. @internal
 */
export function backgroundStatusKey(
  chosen: SlideBackground | null | undefined,
  opening: SlideBackground | null | undefined,
  presets: readonly GradientFill[]
): TranslationKey | null {
  if (chosen === null) return 'background.layoutChosen';
  if (chosen === undefined) return 'background.ownFill';
  if (chosen.kind === 'picture') return 'background.imageChosen';
  const customGradient =
    chosen.kind === 'gradient' &&
    chosen !== opening &&
    !presets.some((g) => sameBackground(chosen, { kind: 'gradient', gradient: g }));
  return customGradient ? 'background.gradientChosen' : null;
}

// --- Dialog sections -----------------------------------------------------------------------

// Which part of a dialog a menu row opens on: Character Spacing > More Spacing… opens the Font
// dialog on its Character Spacing tab. The row asks before it opens the registry dialog, and
// the dialog takes the request once, when it opens. The popups registry opens a dialog by key
// only, so a host dialog in `popups` can ignore the request.

const sectionRequests = new WeakMap<object, Map<string, string>>();

/** Ask `dialog` of `editor` to open on `section` the next time it opens. @internal */
export function requestDialogSection(editor: object, dialog: string, section: string): void {
  const map = sectionRequests.get(editor) ?? new Map<string, string>();
  map.set(dialog, section);
  sectionRequests.set(editor, map);
}

/** The section asked for `dialog` of `editor`, once; undefined when none was asked. @internal */
export function takeDialogSection(editor: object | null, dialog: string): string | undefined {
  if (!editor) return undefined;
  const map = sectionRequests.get(editor);
  const section = map?.get(dialog);
  map?.delete(dialog);
  return section;
}
