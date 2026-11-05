// The Font dialog's form model, shared by the adapters, as `paragraph-dialog-fields.ts` is for
// the Paragraph dialog: the fields as the controls hold them (typed text for numbers, a
// "mixed" state for checkboxes), their seed from `FontFormatRead`, the range checks with a
// message and a tab per field, and the one `setFontFormat` command that OK sends.
//
// A field the selected runs disagree about opens blank (a checkbox opens mixed). OK writes only
// what the user changed; a value set over a mixed field goes to every selected run.

import type {
  CapsStyle,
  FontFormatCommand,
  FontFormatRead,
  StrikeStyle,
  UnderlineStyle,
} from '../contracts/index.ts';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { colorChange, type ColorPick } from './color-fields.ts';
import { amountText, parseAmount } from './paragraph-dialog-fields.ts';

/** The Font style choice: bold and italic together. @public */
export type FontStyleChoice = 'regular' | 'italic' | 'bold' | 'boldItalic';
/** A checkbox: on, off, or mixed (the selected runs disagree). @public */
export type CheckState = boolean | 'mixed';
/** The Character Spacing choice. @public */
export type SpacingChoice = 'normal' | 'expanded' | 'condensed';
/** The dialog's tabs. @internal */
export type FontDialogTab = 'font' | 'spacing';

/** Every field of the form. `''` is a blank: the selected runs disagree. @public */
export interface FontDialogFields {
  readonly fontFamily: string;
  readonly style: FontStyleChoice | '';
  /** Points. */
  readonly size: string;
  /** The text always has a color, so the field opens at "no change". */
  readonly color: ColorPick;
  readonly underline: UnderlineStyle | '';
  /** `auto`: the underline takes the text color (`a:uFillTx`, or no fill). */
  readonly underlineColor: ColorPick;
  readonly strike: CheckState;
  readonly doubleStrike: CheckState;
  readonly superscript: CheckState;
  readonly subscript: CheckState;
  /** Percent, without its sign: Superscript or Subscript gives the direction. */
  readonly offset: string;
  readonly smallCaps: CheckState;
  readonly allCaps: CheckState;
  readonly spacing: SpacingChoice | '';
  /** Points, without its sign: Expanded or Condensed gives the direction. */
  readonly by: string;
  readonly kerning: CheckState;
  /** Points and above. */
  readonly kerningAt: string;
}

/** The offset the field starts from when Superscript is checked. @internal */
export const SUPER_OFFSET = '30';
/** The offset the field starts from when Subscript is checked. @internal */
export const SUB_OFFSET = '25';
/** The size "Points and above" starts from. @internal */
const KERNING_AT = '12';

const styleOf = (bold: boolean, italic: boolean): FontStyleChoice =>
  bold ? (italic ? 'boldItalic' : 'bold') : italic ? 'italic' : 'regular';

/** Open the form on the selection. A field the runs disagree about opens blank. @internal */
export function seedFontFields(f: FontFormatRead | null): FontDialogFields {
  const d = f?.disagrees;
  const mixed = (key: keyof FontFormatRead['disagrees']): boolean => !f || !!d?.[key];
  const strike = f?.strike ?? 'none';
  const baseline = f?.baseline ?? 0;
  const caps = f?.caps ?? 'none';
  const spc = f?.characterSpacingPt ?? 0;
  return {
    fontFamily: mixed('fontFamily') ? '' : (f?.fontFamily ?? ''),
    style: mixed('bold') || mixed('italic') ? '' : styleOf(f?.bold ?? false, f?.italic ?? false),
    size: mixed('fontSizePt') ? '' : amountText(f?.fontSizePt ?? 0),
    color: 'keep',
    underline: mixed('underline') ? '' : (f?.underline ?? 'none'),
    underlineColor: mixed('underlineColor') ? 'mixed' : f?.underlineColor ? 'keep' : 'auto',
    strike: mixed('strike') ? 'mixed' : strike === 'single',
    doubleStrike: mixed('strike') ? 'mixed' : strike === 'double',
    superscript: mixed('baseline') ? 'mixed' : baseline > 0,
    subscript: mixed('baseline') ? 'mixed' : baseline < 0,
    offset: mixed('baseline') ? '' : baseline === 0 ? '0' : amountText(Math.abs(baseline)),
    smallCaps: mixed('caps') ? 'mixed' : caps === 'small',
    allCaps: mixed('caps') ? 'mixed' : caps === 'all',
    spacing: mixed('characterSpacingPt')
      ? ''
      : spc > 0
        ? 'expanded'
        : spc < 0
          ? 'condensed'
          : 'normal',
    by: mixed('characterSpacingPt') ? '' : amountText(Math.abs(spc)),
    kerning: mixed('kerningPt') ? 'mixed' : f?.kerningPt !== null && f !== null,
    kerningAt: mixed('kerningPt') ? '' : amountText(f?.kerningPt ?? Number(KERNING_AT)),
  };
}

/** Superscript and Subscript exclude each other; checking one starts its offset. @internal */
export function withScript(
  d: FontDialogFields,
  which: 'superscript' | 'subscript',
  on: boolean
): FontDialogFields {
  const other = which === 'superscript' ? 'subscript' : 'superscript';
  if (!on) return { ...d, [which]: false };
  const offset = which === 'superscript' ? SUPER_OFFSET : SUB_OFFSET;
  return { ...d, [which]: true, [other]: false, offset };
}

/** Pairs that exclude each other: Strikethrough and Double, Small Caps and All Caps. @internal */
export function withExclusive(
  d: FontDialogFields,
  which: 'strike' | 'doubleStrike' | 'smallCaps' | 'allCaps',
  on: boolean
): FontDialogFields {
  const other = {
    strike: 'doubleStrike',
    doubleStrike: 'strike',
    smallCaps: 'allCaps',
    allCaps: 'smallCaps',
  }[which] as typeof which;
  return on ? { ...d, [which]: true, [other]: false } : { ...d, [which]: false };
}

/** Expanded and Condensed start "By" at 1 pt when it held nothing. @internal */
export function withSpacing(d: FontDialogFields, spacing: SpacingChoice): FontDialogFields {
  const by = parseAmount(d.by);
  return {
    ...d,
    spacing,
    by: spacing === 'normal' ? '0' : Number.isFinite(by) && by > 0 ? d.by : '1',
  };
}

/** A field's range problem: the field, its tab, the message key, and its parameters. @internal */
export interface FontFieldError {
  readonly field: 'size' | 'offset' | 'by' | 'kerningAt';
  readonly tab: FontDialogTab;
  readonly key: TranslationKey;
  readonly params: Readonly<Record<string, number>>;
}

/**
 * The first changed field OK cannot apply, or null. The limits are the engine's
 * (`checkFontFormat`). @internal
 */
export function fontFieldError(seed: FontDialogFields, d: FontDialogFields): FontFieldError | null {
  const bad = (
    field: FontFieldError['field'],
    tab: FontDialogTab,
    min: number,
    max: number,
    unit: 'pt' | 'percent',
    changed: boolean
  ): FontFieldError | null => {
    if (!changed) return null;
    const v = parseAmount(d[field]);
    if (Number.isFinite(v) && v >= min && v <= max) return null;
    return { field, tab, key: `fontDialog.errors.${unit}` as TranslationKey, params: { min, max } };
  };
  const scriptOn = d.superscript === true || d.subscript === true;
  return (
    bad('size', 'font', 1, 4000, 'pt', d.size !== seed.size) ??
    bad(
      'offset',
      'font',
      0,
      100,
      'percent',
      scriptOn &&
        (d.offset !== seed.offset ||
          d.superscript !== seed.superscript ||
          d.subscript !== seed.subscript)
    ) ??
    bad(
      'by',
      'spacing',
      0,
      1584,
      'pt',
      d.spacing !== '' && d.spacing !== 'normal' && (d.by !== seed.by || d.spacing !== seed.spacing)
    ) ??
    bad(
      'kerningAt',
      'spacing',
      // 0 is "Kerning for fonts" off in the file, so the size starts at the smallest step.
      0.01,
      1584,
      'pt',
      d.kerning === true && (d.kerningAt !== seed.kerningAt || d.kerning !== seed.kerning)
    )
  );
}

/**
 * The command OK sends: only what moved from the seed. Two controls over one value (Style for
 * bold and italic, the strike pair, the script pair and Offset, the caps pair, Spacing and By,
 * Kerning and its size) write that value when either moves and neither is still mixed. Null
 * when nothing moved. @internal
 */
export function fontCommandOf(
  seed: FontDialogFields,
  d: FontDialogFields
): FontFormatCommand | null {
  const out: { -readonly [K in keyof FontFormatCommand]: FontFormatCommand[K] } = {};
  const moved = (...keys: (keyof FontDialogFields)[]): boolean =>
    keys.some((k) => d[k] !== seed[k]);
  const settled = (...keys: (keyof FontDialogFields)[]): boolean =>
    keys.every((k) => d[k] !== 'mixed' && d[k] !== '');
  if (d.fontFamily !== '' && moved('fontFamily')) out.fontFamily = d.fontFamily;
  if (d.size !== '' && moved('size')) out.fontSizePt = parseAmount(d.size);
  // A Style choice decides both: Bold Italic over a mixed selection makes every run both.
  if (d.style !== '' && moved('style')) {
    out.bold = d.style === 'bold' || d.style === 'boldItalic';
    out.italic = d.style === 'italic' || d.style === 'boldItalic';
  }
  const color = colorChange(seed.color, d.color);
  if (color !== undefined) out.color = color;
  if (d.underline !== '' && moved('underline')) out.underline = d.underline;
  // A dimmed underline color (no underline) is not sent.
  const underlined = (d.underline !== '' ? d.underline : seed.underline) !== 'none';
  const underlineColor = colorChange(seed.underlineColor, d.underlineColor);
  if (underlineColor !== undefined && underlined) out.underlineColor = underlineColor;
  if (settled('strike', 'doubleStrike') && moved('strike', 'doubleStrike'))
    out.strike = (d.doubleStrike ? 'double' : d.strike ? 'single' : 'none') satisfies StrikeStyle;
  if (settled('superscript', 'subscript') && moved('superscript', 'subscript', 'offset')) {
    const offset = d.superscript || d.subscript ? parseAmount(d.offset) : 0;
    out.baseline = d.subscript ? -offset : offset;
  }
  if (settled('smallCaps', 'allCaps') && moved('smallCaps', 'allCaps'))
    out.caps = (d.allCaps ? 'all' : d.smallCaps ? 'small' : 'none') satisfies CapsStyle;
  if (d.spacing !== '' && moved('spacing', 'by')) {
    const by = d.spacing === 'normal' ? 0 : parseAmount(d.by);
    out.characterSpacingPt = d.spacing === 'condensed' ? -by : by;
  }
  if (d.kerning !== 'mixed' && moved('kerning', 'kerningAt'))
    out.kerningPt = d.kerning ? parseAmount(d.kerningAt) : null;
  return Object.keys(out).length ? out : null;
}
