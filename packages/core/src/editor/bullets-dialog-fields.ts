// The Bullets and Numbering dialog's form model, shared by the adapters, as
// `font-dialog-fields.ts` is for the Font dialog: the fields as the controls hold them, their
// seed from `BulletFormatRead`, the range checks with a message per field, and the one
// `setBulletFormat` command that OK sends.
//
// Selected paragraphs that differ open with a mixed style and a blank size. OK sends only what
// the user changed: a size or color change alone keeps each paragraph's own bullet. The
// Bulleted tab also takes any character in any font, as Customize… does.

import type { BulletFormat, BulletFormatMixed, BulletFormatRead } from '../contracts/index.ts';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import {
  BULLET_GALLERY,
  NUMBER_GALLERY,
  galleryFormat,
  galleryKeyOf,
  isBulletChar,
  type ListGalleryItem,
} from './bullet-format.ts';
import { colorChange, type ColorPick } from './color-fields.ts';
import { parseAmount } from './paragraph-dialog-fields.ts';
import { autonumberText, symbolBulletText } from '../layout/index.ts';

/** The dialog's tabs, after the list kinds. @public */
export type BulletsDialogTab = 'bullet' | 'number';

/** The style row's value for the paragraphs' own bullet or scheme, when no gallery item matches. @internal */
export const CURRENT_STYLE = 'current';

/** The first gallery bullet, which the Bulleted tab shows for paragraphs with none. */
const FIRST_BULLET = { char: '•', font: 'Arial' };

/** Every field of the form. `''` is a blank: the selected paragraphs differ. @public */
export interface BulletsDialogFields {
  /** The open tab: Bulleted or Numbered. */
  readonly kind: BulletsDialogTab;
  /** A Bullets gallery key, `current` (the paragraphs' own or a custom bullet), or `''`. */
  readonly bullet: string;
  /** The bullet character, in its font's own code (`v` is the Wingdings star). */
  readonly char: string;
  /** The bullet font; `''` follows the text font. */
  readonly font: string;
  /** A Numbering gallery key, `current`, or `''`. */
  readonly number: string;
  /** Percent of the text size. */
  readonly sizePercent: string;
  /** The bullet or number color, as the color field picks it. */
  readonly color: ColorPick;
  /** The first number, 1 to 32767, as typed. */
  readonly startAt: string;
}

/**
 * Open the form on the selection: `undefined` when the paragraphs differ, with `mixed`, what
 * they share; null when they have no list. The form opens on the tab of the paragraphs' list.
 * `tab` is the tab a menu row asks for (the Numbering gallery opens Numbered): it applies when
 * the paragraphs have no list, or both kinds. @internal
 */
export function seedBulletFields(
  read: BulletFormatRead | null | undefined,
  tab?: BulletsDialogTab,
  mixed?: BulletFormatMixed
): BulletsDialogFields {
  if (read === undefined) {
    const lists = mixed?.kinds.filter((k) => k !== 'none') ?? [];
    const start = mixed && mixed.startAt === undefined ? 1 : mixed?.startAt;
    return {
      kind: lists.length === 1 ? lists[0]! : (tab ?? lists[0] ?? 'bullet'),
      bullet: '',
      char: '',
      font: '',
      number: '',
      sizePercent: mixed?.sizePercent == null ? '' : String(mixed.sizePercent),
      color: 'keep',
      startAt: start == null ? '' : String(start),
    };
  }
  if (read === null)
    return {
      kind: tab ?? 'bullet',
      bullet: 'none',
      ...FIRST_BULLET,
      number: 'none',
      sizePercent: '100',
      color: 'auto',
      startAt: '1',
    };
  const key = (gallery: readonly ListGalleryItem[]) => galleryKeyOf(gallery, read) ?? CURRENT_STYLE;
  return {
    kind: read.kind,
    bullet: read.kind === 'bullet' ? key(BULLET_GALLERY) : 'none',
    ...(read.kind === 'bullet' ? { char: read.char, font: read.font ?? '' } : FIRST_BULLET),
    number: read.kind === 'number' ? key(NUMBER_GALLERY) : 'none',
    sizePercent: String(read.sizePercent),
    color: read.color ? 'keep' : 'auto',
    startAt: String(read.kind === 'number' ? read.startAt : 1),
  };
}

/**
 * A preview of the paragraphs' own bullet for the "Current" style row: the character (a symbol
 * font's Unicode twin), or the first three numbers of the scheme ("I) II) III)"). @internal
 */
export function currentStylePreview(read: BulletFormatRead): string {
  if (read.kind === 'bullet') return symbolBulletText(read.char, read.font);
  return [1, 2, 3].map((n) => autonumberText(read.scheme, n)).join(' ');
}

/**
 * A Bullets style pick: the gallery item's character and font show in their fields, so an edit
 * starts from them. @internal
 */
export function pickBulletStyle(d: BulletsDialogFields, key: string): BulletsDialogFields {
  const f = galleryFormat(BULLET_GALLERY, key);
  return f?.kind === 'bullet' && f.char
    ? { ...d, bullet: key, char: f.char, font: f.font ?? '' }
    : { ...d, bullet: key };
}

/**
 * An edit of the character or the font: the bullet becomes the gallery item with that character
 * and font, or, when the edit brings back the character and font the dialog opened with, the
 * style it opened with, or else a custom one. A dialog that opened on paragraphs with no list
 * never goes back to None this way: the fields then show the first gallery bullet.
 * @internal
 */
export function editBulletGlyph(
  d: BulletsDialogFields,
  field: 'char' | 'font',
  value: string,
  seed?: BulletsDialogFields
): BulletsDialogFields {
  const next = { ...d, [field]: value };
  const item = BULLET_GALLERY.find(
    ({ format: f }) => f?.kind === 'bullet' && f.char === next.char && (f.font ?? '') === next.font
  );
  // No way back to None or to a mixed style row: those had no character or font of their own.
  const back =
    seed &&
    seed.bullet !== 'none' &&
    seed.bullet !== '' &&
    next.char === seed.char &&
    next.font === seed.font;
  return { ...next, bullet: item?.key ?? (back ? seed.bullet : CURRENT_STYLE) };
}

/**
 * The Font field's families: the fonts the field shows now and opened with, the gallery's
 * bullet fonts and Symbol (so a custom character keeps its symbol font), then `families`.
 * @internal
 */
export function bulletFontChoices(
  current: readonly string[],
  families: readonly string[]
): string[] {
  const gallery = BULLET_GALLERY.flatMap(({ format: f }) =>
    f?.kind === 'bullet' && f.font ? [f.font] : []
  );
  return [...new Set([...current, ...gallery, 'Symbol', ...families])].filter(Boolean);
}

/**
 * Whether the user typed a character or picked a font of their own on the Bulleted tab. Over
 * paragraphs whose bullets differ, any character or font edit makes a custom bullet, so a pick
 * of "Text font" asks for a character as a pick of Symbol does. @internal
 */
export function customBullet(seed: BulletsDialogFields, d: BulletsDialogFields): boolean {
  return (
    d.kind === 'bullet' &&
    d.bullet === CURRENT_STYLE &&
    (d.char !== seed.char || d.font !== seed.font || seed.bullet === '')
  );
}

/**
 * The Font field's option value for "Text font" (the form's empty font), so an adapter's select
 * can keep its blank for "Mixed". U+0001 cannot come from a file (XML 1.0). @internal
 */
export const BULLET_TEXT_FONT = '\u0001text';

/**
 * Whether the Font field shows "Mixed": the bullets of the selected paragraphs differ and no
 * style, character, or font is chosen yet. The read has no per-field state, so this is also
 * true when only the size or the color differs. @internal
 */
export function bulletFontMixed(seed: BulletsDialogFields, d: BulletsDialogFields): boolean {
  return seed.bullet === '' && d.bullet === '';
}

/** A preview of a custom bullet for the style row: the character, or its symbol font's twin. @internal */
export function customBulletPreview(d: BulletsDialogFields): string {
  return symbolBulletText(d.char, d.font);
}

/** A field's problem: the field, the message key, and its parameters. @internal */
export interface BulletsFieldError {
  readonly field: 'sizePercent' | 'startAt' | 'char';
  readonly key: TranslationKey;
  readonly params: Readonly<Record<string, number>>;
}

/**
 * The first changed field OK cannot apply, or null. The limits are the engine's; dimmed fields
 * (style None) are not checked. @internal
 */
export function bulletsFieldError(
  seed: BulletsDialogFields,
  d: BulletsDialogFields
): BulletsFieldError | null {
  if (!bulletFieldsEnabled(d)) return null;
  if (customBullet(seed, d) && !isBulletChar(d.char))
    return { field: 'char', key: 'bulletsAndNumberingDialog.errors.character', params: {} };
  const size = parseAmount(d.sizePercent);
  if (d.sizePercent !== seed.sizePercent && !(Number.isFinite(size) && size >= 25 && size <= 400))
    return {
      field: 'sizePercent',
      key: 'bulletsAndNumberingDialog.errors.percent',
      params: { min: 25, max: 400 },
    };
  const start = parseAmount(d.startAt);
  if (
    d.kind === 'number' &&
    d.startAt !== seed.startAt &&
    !(Number.isInteger(start) && start >= 1 && start <= 32_767)
  )
    return {
      field: 'startAt',
      key: 'bulletsAndNumberingDialog.errors.whole',
      params: { min: 1, max: 32_767 },
    };
  return null;
}

/**
 * Whether the size, color, and start fields apply: not while the tab's style is None, where
 * they are dimmed. @internal
 */
export function bulletFieldsEnabled(d: BulletsDialogFields): boolean {
  return (d.kind === 'bullet' ? d.bullet : d.number) !== 'none';
}

/**
 * What OK sends: a bullet format, null for None, or undefined when nothing moved. The style is
 * sent only when its row moved; a tab change alone moves nothing. A size or color change alone
 * sends no kind, so each paragraph keeps its own bullet or number, also in a selection that
 * mixes them; a start change numbers the paragraphs. @internal
 */
export function bulletCommandOf(
  seed: BulletsDialogFields,
  d: BulletsDialogFields
): BulletFormat | null | undefined {
  const key = d.kind === 'bullet' ? d.bullet : d.number;
  const moved = key !== (d.kind === 'bullet' ? seed.bullet : seed.number);
  if (key === 'none') return moved ? null : undefined;
  const font = d.font.trim();
  const picked: BulletFormat | null | undefined = customBullet(seed, d)
    ? { kind: 'bullet', char: d.char, ...(font ? { font } : {}) }
    : moved
      ? galleryFormat(d.kind === 'bullet' ? BULLET_GALLERY : NUMBER_GALLERY, key)
      : undefined;
  const sizePercent =
    d.sizePercent !== seed.sizePercent && d.sizePercent !== ''
      ? parseAmount(d.sizePercent)
      : undefined;
  const color = colorChange(seed.color, d.color);
  const startAt =
    d.kind === 'number' && d.startAt !== seed.startAt && d.startAt !== ''
      ? parseAmount(d.startAt)
      : undefined;
  const look = {
    ...(sizePercent !== undefined ? { sizePercent } : {}),
    ...(color !== undefined ? { color } : {}),
  };
  if (picked)
    return {
      ...picked,
      ...look,
      ...(startAt !== undefined && picked.kind === 'number' ? { startAt } : {}),
    };
  if (startAt !== undefined) return { kind: 'number', ...look, startAt };
  return sizePercent !== undefined || color !== undefined ? look : undefined;
}
