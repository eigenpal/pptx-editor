// The Paragraph dialog's form model, shared by the adapters: the fields as the controls hold
// them (typed text for numbers), their seed from `ParagraphFormatRead`, the range checks with a
// message per field, and the one `setParagraphFormat` command that OK sends. After DOCX
// Editor's `packages/core/src/editor/paragraph-dialog-fields.ts`, in slide units: inches
// for indents, points or lines for spacing. DOCX Editor commit adf0e178.
//
// A field the selection disagrees about opens blank (or on a "Mixed" choice). A blank is not a
// value, so OK writes only what the user set; any value typed into a blank field differs from
// the seed and is written to every selected paragraph.

import type {
  ParagraphAlignment,
  ParagraphFormatCommand,
  ParagraphFormatRead,
} from '../contracts/index.ts';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import {
  MAX_LINES,
  MAX_MARGIN,
  MAX_MULTIPLE,
  MAX_POINTS,
} from '../automation/commands/paragraph-format.ts';

export const EMU_PER_INCH = 914_400;

/** The Special choice: the sign of the first-line offset. @public */
export type SpecialIndent = 'none' | 'firstLine' | 'hanging';
/** The Line Spacing choice, as the dialog lists it. @public */
export type LineRule = 'single' | 'onePointFive' | 'double' | 'exact' | 'multiple';
/** The unit of a Before or After field. @public */
export type SpacingUnit = 'pt' | 'lines';

/**
 * Every field of the form, as the controls hold it. `''` is a blank: the selection disagrees
 * and the user has not chosen. @public
 */
export interface ParagraphDialogFields {
  readonly alignment: ParagraphAlignment | '';
  /** Inches. */
  readonly indentLeft: string;
  readonly special: SpecialIndent | '';
  /** Inches. */
  readonly specialBy: string;
  readonly spaceBefore: string;
  readonly spaceBeforeUnit: SpacingUnit;
  readonly spaceAfter: string;
  readonly spaceAfterUnit: SpacingUnit;
  readonly lineRule: LineRule | '';
  readonly lineAt: string;
}

/** The fields a range check names. @internal */
export type ParagraphDialogField = Exclude<
  keyof ParagraphDialogFields,
  'spaceBeforeUnit' | 'spaceAfterUnit'
>;

/** A number as a field shows it: at most two decimals. @internal */
export const amountText = (n: number): string => String(Math.round(n * 100) / 100);

/** A field's number, or NaN when it is empty or not a number. @internal */
export const parseAmount = (text: string): number =>
  text.trim() === '' ? Number.NaN : Number(text);

const inchesText = (emu: number | null): string =>
  emu === null ? '' : amountText(emu / EMU_PER_INCH);

const specialOf = (signed: number): SpecialIndent =>
  signed === 0 ? 'none' : signed < 0 ? 'hanging' : 'firstLine';

function lineRuleOf(s: ParagraphFormatRead['lineSpacing']): LineRule | '' {
  if (!s) return '';
  if (s.rule === 'exact') return 'exact';
  if (s.value === 1) return 'single';
  if (s.value === 1.5) return 'onePointFive';
  return s.value === 2 ? 'double' : 'multiple';
}

function spaceOf(points: number | null, lines: number | null): [string, SpacingUnit] {
  if (points !== null) return [amountText(points), 'pt'];
  if (lines !== null) return [amountText(lines), 'lines'];
  return ['', 'pt'];
}

/**
 * Open the form on the selection. A field the paragraphs disagree about opens blank, and so
 * does every field when nothing could be read (the first selected shape has no text). @internal
 */
export function seedParagraphFields(format: ParagraphFormatRead | null): ParagraphDialogFields {
  if (!format)
    return {
      alignment: '',
      indentLeft: '',
      special: '',
      specialBy: '',
      spaceBefore: '',
      spaceBeforeUnit: 'pt',
      spaceAfter: '',
      spaceAfterUnit: 'pt',
      lineRule: '',
      lineAt: '',
    };
  const [spaceBefore, spaceBeforeUnit] = spaceOf(format.spaceBeforePt, format.spaceBeforeLines);
  const [spaceAfter, spaceAfterUnit] = spaceOf(format.spaceAfterPt, format.spaceAfterLines);
  const firstLine = format.indentFirstLineEmu;
  return {
    alignment: format.alignment ?? '',
    indentLeft: inchesText(format.indentLeftEmu),
    special: firstLine === null ? '' : specialOf(firstLine),
    specialBy: firstLine === null ? '' : inchesText(Math.abs(firstLine)),
    spaceBefore,
    spaceBeforeUnit,
    spaceAfter,
    spaceAfterUnit,
    lineRule: lineRuleOf(format.lineSpacing),
    lineAt: format.lineSpacing ? amountText(format.lineSpacing.value) : '',
  };
}

/** The "At" value a Line Spacing choice starts with. @internal */
export function lineAtFor(rule: LineRule, previous: ParagraphDialogFields): string {
  if (rule === 'single') return '1';
  if (rule === 'onePointFive') return '1.5';
  if (rule === 'double') return '2';
  if (rule === 'exact') return previous.lineRule === 'exact' ? previous.lineAt : '12';
  // Multiple keeps a multiple; it never takes a number of points as lines.
  return previous.lineRule === 'exact' || previous.lineRule === '' ? '1' : previous.lineAt;
}

/**
 * The fields after a Before or After unit change. A number the user typed stays (it is meant
 * in the new unit); the opened value starts again at 0, because a number in points is not the
 * same number in lines; a blank stays blank, so a look at the unit list writes nothing.
 * @internal
 */
export function withSpaceUnit(
  d: ParagraphDialogFields,
  seed: ParagraphDialogFields,
  side: 'spaceBefore' | 'spaceAfter',
  unit: SpacingUnit
): ParagraphDialogFields {
  const unitKey = side === 'spaceBefore' ? 'spaceBeforeUnit' : 'spaceAfterUnit';
  if (unit === d[unitKey]) return d;
  // Back to the opened unit with the reset 0 still there: the opened value comes back.
  if (unit === seed[unitKey] && d[side] === '0')
    return { ...d, [unitKey]: unit, [side]: seed[side] };
  const opened = d[side] === seed[side];
  return { ...d, [unitKey]: unit, [side]: d[side] === '' || !opened ? d[side] : '0' };
}

/** The "By" value a Special choice starts with: half an inch when there was none. @internal */
export function specialByFor(special: SpecialIndent, previous: ParagraphDialogFields): string {
  if (special === 'none') return '0';
  const by = parseAmount(previous.specialBy);
  return Number.isFinite(by) && by > 0 ? previous.specialBy : '0.5';
}

/** A field's range problem: the field, the message key, and its parameters. @internal */
export interface ParagraphFieldError {
  readonly field: ParagraphDialogField;
  readonly key: TranslationKey;
  readonly params: Readonly<Record<string, string | number>>;
}

/**
 * The first field OK cannot apply, or null. Only a field the user changed is checked (with its
 * unit or rule), so a value the file holds outside the dialog's limits never blocks OK: OK
 * does not write it. A changed field must hold a number in range. The limits are the engine's
 * (`checkParagraphFormat`), in the dialog's units. @internal
 */
export function paragraphFieldError(
  seed: ParagraphDialogFields,
  current: ParagraphDialogFields
): ParagraphFieldError | null {
  const maxInches = amountText(MAX_MARGIN / EMU_PER_INCH);
  const check = (
    field: ParagraphDialogField,
    min: number,
    max: number,
    unit: 'in' | 'pt' | 'lines' | 'multiple',
    unchanged: boolean
  ): ParagraphFieldError | null => {
    if (unchanged) return null;
    const v = parseAmount(current[field]);
    if (Number.isFinite(v) && v >= min && v <= max) return null;
    return { field, key: `paragraphDialog.errors.${unit}` as TranslationKey, params: { min, max } };
  };
  const same = (...keys: (keyof ParagraphDialogFields)[]): boolean =>
    keys.every((k) => current[k] === seed[k]);
  // A field left blank writes nothing, whatever its unit, so it needs no check.
  const blank = (k: ParagraphDialogField): boolean => current[k] === '' && seed[k] === '';
  const spaceMax = (unit: SpacingUnit): number => (unit === 'pt' ? MAX_POINTS : MAX_LINES);
  // The smallest At the dialog writes: one hundredth, the engine's smallest nonzero value.
  const lineAt =
    current.lineRule === 'exact'
      ? check('lineAt', 0.01, MAX_POINTS, 'pt', same('lineRule', 'lineAt'))
      : current.lineRule === 'multiple'
        ? check('lineAt', 0.01, MAX_MULTIPLE, 'multiple', same('lineRule', 'lineAt'))
        : null;
  return (
    check('indentLeft', 0, Number(maxInches), 'in', same('indentLeft')) ??
    (current.special === 'none' || current.special === ''
      ? null
      : check('specialBy', 0, Number(maxInches), 'in', same('special', 'specialBy'))) ??
    check(
      'spaceBefore',
      0,
      spaceMax(current.spaceBeforeUnit),
      current.spaceBeforeUnit,
      same('spaceBefore', 'spaceBeforeUnit') || blank('spaceBefore')
    ) ??
    check(
      'spaceAfter',
      0,
      spaceMax(current.spaceAfterUnit),
      current.spaceAfterUnit,
      same('spaceAfter', 'spaceAfterUnit') || blank('spaceAfter')
    ) ??
    lineAt
  );
}

const emu = (inches: string): number => Math.round(parseAmount(inches) * EMU_PER_INCH);

/**
 * The command OK sends: only the fields that moved from the seed. Null when nothing moved, so
 * the caller just closes (an empty write would still push an undo step). @internal
 */
export function paragraphCommandOf(
  seed: ParagraphDialogFields,
  current: ParagraphDialogFields
): ParagraphFormatCommand | null {
  const out: { -readonly [K in keyof ParagraphFormatCommand]: ParagraphFormatCommand[K] } = {};
  if (current.alignment !== '' && current.alignment !== seed.alignment)
    out.alignment = current.alignment;
  if (current.indentLeft !== '' && current.indentLeft !== seed.indentLeft)
    out.indentLeftEmu = emu(current.indentLeft);
  // Special and By are two controls over one signed value: either one moving rewrites it.
  if (
    current.special !== '' &&
    (current.special !== seed.special || current.specialBy !== seed.specialBy)
  ) {
    const by = current.special === 'none' ? 0 : Math.abs(emu(current.specialBy));
    out.indentFirstLineEmu = current.special === 'hanging' ? -by : by;
  }
  const space = (
    text: string,
    unit: SpacingUnit,
    seedText: string,
    seedUnit: SpacingUnit,
    pt: 'spaceBeforePt' | 'spaceAfterPt',
    lines: 'spaceBeforeLines' | 'spaceAfterLines'
  ): void => {
    if (text === '' || (text === seedText && unit === seedUnit)) return;
    if (unit === 'pt') out[pt] = parseAmount(text);
    else out[lines] = parseAmount(text);
  };
  space(
    current.spaceBefore,
    current.spaceBeforeUnit,
    seed.spaceBefore,
    seed.spaceBeforeUnit,
    'spaceBeforePt',
    'spaceBeforeLines'
  );
  space(
    current.spaceAfter,
    current.spaceAfterUnit,
    seed.spaceAfter,
    seed.spaceAfterUnit,
    'spaceAfterPt',
    'spaceAfterLines'
  );
  const rule = current.lineRule;
  if (rule !== '' && (rule !== seed.lineRule || current.lineAt !== seed.lineAt)) {
    if (rule === 'single') out.lineSpacing = { rule: 'multiple', value: 1 };
    else if (rule === 'onePointFive') out.lineSpacing = { rule: 'multiple', value: 1.5 };
    else if (rule === 'double') out.lineSpacing = { rule: 'multiple', value: 2 };
    else out.lineSpacing = { rule, value: parseAmount(current.lineAt) };
  }
  return Object.keys(out).length ? out : null;
}
