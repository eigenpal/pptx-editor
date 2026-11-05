// The Paragraph group's controls of the chrome registry: the Bullets and Numbering galleries,
// Text Direction, Columns with More Columns…, Line Spacing with its dialog row, and the
// Paragraph dialog slot. Their value commands are in
// `chrome-slot-commands.ts`. Split from `chrome-controls.ts`, which places them in `CHROME_GROUPS`;
// type-only imports keep the modules free of an import cycle.

import type { TranslationKey } from '@pptx-editor.dev/i18n';
import type { EditorSnapshot } from '../contracts/index.ts';
import type { ChromeControl } from './chrome-controls.ts';
import { symbolBulletText } from '../layout/index.ts';
import { BULLET_GALLERY, NUMBER_GALLERY, galleryKeyOf } from './bullet-format.ts';
import { ICON_PATHS } from './icon-paths.ts';
import { TEXT_DIRECTIONS } from './text-body-commands.ts';

const icon = (name: string): readonly string[] => ICON_PATHS[name] ?? [];

const TEXT_DIRECTION_ICONS: Readonly<Record<(typeof TEXT_DIRECTIONS)[number], string>> = {
  horizontal: 'text_rotation_none',
  vertical: 'text_rotation_down',
  vertical270: 'text_rotate_up',
  wordArtVertical: 'text_rotate_vertical',
  eastAsianVertical: 'format_textdirection_vertical',
};

/** The `textDirection` control. @internal */
export const TEXT_DIRECTION_CONTROL: ChromeControl = {
  // Home > Text Direction: horizontal, rotated 90° or 270°, stacked, and upright East
  // Asian text.
  id: 'textDirection',
  shape: 'dropdown',
  labelKey: 'toolbar.textDirection',
  toggle: 'choice',
  paths: icon('text_rotate_vertical'),
  // Each row shows its Material Symbol, and the button shows the current row's.
  options: TEXT_DIRECTIONS.map((d) => ({
    value: d,
    labelKey: `toolbar.textDirections.${d}` as const,
    paths: icon(TEXT_DIRECTION_ICONS[d]),
  })),
  state: { kind: 'value' },
};

/** The `autofit` control. @internal */
export const AUTOFIT_CONTROL: ChromeControl = {
  // How the selected text bodies fit their shapes (ECMA-376 Part 1, 21.1.2.1): `a:noAutofit`,
  // `a:normAutofit` (shrink text on overflow), or `a:spAutoFit` (resize shape to fit text).
  id: 'autofit',
  shape: 'dropdown',
  labelKey: 'panel.format.textFitting',
  toggle: 'choice',
  paths: icon('text_fields'),
  options: [
    { value: 'none', labelKey: 'panel.format.fitting.none', paths: icon('text_fields') },
    { value: 'shrink', labelKey: 'panel.format.fitting.shrink', paths: icon('unfold_less') },
    { value: 'shape', labelKey: 'panel.format.fitting.shape', paths: icon('unfold_more') },
  ],
  state: { kind: 'value' },
};

/** The `columns` control. @internal */
export const COLUMNS_CONTROL: ChromeControl = {
  // Home > Columns: one, two, or three columns, and More Columns… for the count and gap. Off
  // the default bar: Format > Columns… sets them.
  id: 'columns',
  defaultToolbar: false,
  shape: 'dropdown',
  labelKey: 'toolbar.columns',
  toggle: 'choice',
  paths: icon('view_column'),
  options: [
    ...(
      [
        ['1', 'one', 'check_box_outline_blank'],
        ['2', 'two', 'view_column_2'],
        ['3', 'three', 'view_column'],
      ] as const
    ).map(([value, name, symbol]) => ({
      value,
      labelKey: `toolbar.columnCounts.${name}` as const,
      paths: icon(symbol),
    })),
    { value: 'dialog', labelKey: 'toolbar.columnsMore', dialog: 'columns' },
  ],
  state: { kind: 'value' },
};

/** The `lineSpacing` control. @internal */
export const LINE_SPACING_CONTROL: ChromeControl = {
  id: 'lineSpacing',
  shape: 'dropdown',
  labelKey: 'toolbar.lineSpacing',
  paths: icon('format_line_spacing'),
  // One spacing of several: the menu marks the current one (`textSlotValue`).
  toggle: 'choice',
  options: [
    ...['1', '1.15', '1.5', '2'].map((v) => ({
      value: v,
      labelKey: 'toolbar.lineSpacingValue' as const,
      labelParams: { value: v },
    })),
    // The last row: Line Spacing Options… opens the Paragraph dialog.
    { value: 'dialog', labelKey: 'toolbar.lineSpacingOptions', dialog: 'paragraph' },
  ],
  state: { kind: 'value' },
};

/** The `dialog` control. @internal */
export const PARAGRAPH_DIALOG_CONTROL: ChromeControl = {
  // The Paragraph dialog. DOCX Editor's slot of the same name. Contextual, as in DOCX
  // Editor (no dialog launchers on the ribbon): the Line Spacing
  // menu row and the text context menu open it.
  id: 'dialog',
  labelKey: 'toolbar.paragraphDialog',
  paths: icon('format_paragraph'),
  defaultToolbar: false,
  state: { kind: 'dialog', dialog: 'paragraph' },
};

/** The `bulletStyle` control. @internal */
export const BULLET_STYLE_CONTROL: ChromeControl = {
  // Home > Bullets, the gallery: the arrow half of the Bullets split button, right
  // after the toggle. Each row previews its bullet (a symbol font's Unicode twin).
  id: 'bulletStyle',
  shape: 'dropdown',
  labelKey: 'toolbar.bulletStyle',
  toggle: 'choice',
  paths: null,
  options: [
    ...BULLET_GALLERY.map(({ key, format }) => ({
      value: key,
      labelKey: `toolbar.bulletStyles.${key}` as TranslationKey,
      ...(format?.kind === 'bullet' && format.char
        ? { glyph: symbolBulletText(format.char, format.font) }
        : {}),
    })),
    {
      value: 'dialog',
      labelKey: 'toolbar.bulletsAndNumbering',
      dialog: 'bulletsAndNumbering',
    },
  ],
  state: { kind: 'value' },
};

/** The `numberStyle` control. @internal */
export const NUMBER_STYLE_CONTROL: ChromeControl = {
  // Home > Numbering, the gallery: the arrow half of the Numbering split button.
  id: 'numberStyle',
  shape: 'dropdown',
  labelKey: 'toolbar.numberStyle',
  toggle: 'choice',
  paths: null,
  options: [
    ...NUMBER_GALLERY.map(({ key }) => ({
      value: key,
      labelKey: `toolbar.numberStyles.${key}` as TranslationKey,
    })),
    {
      value: 'dialog',
      labelKey: 'toolbar.bulletsAndNumbering',
      dialog: 'bulletsAndNumbering',
    },
  ],
  state: { kind: 'value' },
};

/**
 * The font size box's value: the size as drawn, so shrunk text shows its shrunk size
 * (`a:normAutofit/@fontScale`), and no size when the selected shapes draw different sizes.
 * @internal
 */
export function fontSizeSlotValue(f: EditorSnapshot['format']): number | undefined {
  return f?.renderedFontSize === null ? undefined : (f?.renderedFontSize ?? f?.fontSize);
}

/**
 * The line spacing the Line Spacing menu marks: a multiple (`1.15`), rounded to the menu's two
 * decimals. Undefined for spacing in points and for paragraphs that differ. @internal
 */
export function lineSpacingSlotValue(f: EditorSnapshot['format']): number | undefined {
  const spacing = f?.paragraph?.lineSpacing;
  return spacing?.rule === 'multiple' ? Math.round(spacing.value * 100) / 100 : undefined;
}

/**
 * The value a slot of this module or of `text.characterSpacing` shows: the spacing in points,
 * the gallery key, the text direction, or the column count. Undefined for any other slot, and
 * when the selection has no single value. @internal
 */
export function textSlotValue(id: string, f: EditorSnapshot['format']): string | undefined {
  if (id === 'text.characterSpacing')
    return typeof f?.font?.characterSpacingPt === 'number'
      ? String(f.font.characterSpacingPt)
      : undefined;
  if (id === 'list.bulletStyle' || id === 'list.numberStyle')
    return galleryKeyOf(
      id === 'list.bulletStyle' ? BULLET_GALLERY : NUMBER_GALLERY,
      f?.bulletFormat
    );
  if (id === 'paragraph.textDirection') return f?.textDirection;

  return id === 'paragraph.columns' && f?.columns ? String(f.columns.count) : undefined;
}

/** The probe values of the text format slots: `can` runs their command with these. @internal */
export const TEXT_FORMAT_PROBES: Readonly<Record<string, unknown>> = {
  'paragraph.dialog': { alignment: 'left' },
  'text.characterSpacing': 0,
  'paragraph.textDirection': 'horizontal',
  'paragraph.autofit': 'shrink',
  'list.bulletStyle': 'filledRound',
  'list.numberStyle': 'arabicPeriod',
  'paragraph.columns': 1,
  'text.changeCase': 'upper',
  'text.dialog': { bold: false },
};
