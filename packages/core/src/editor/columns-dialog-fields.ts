// The Columns dialog's form model (Home > Columns > More Columns…), shared by the adapters, as
// `paragraph-dialog-fields.ts` is for the Paragraph dialog: the fields as the controls hold them
// (typed text), their seed from the selection, the range checks with a message per field, and
// the one `setColumns` command that OK sends. The spacing shows in inches.

import type { TextColumns } from '../contracts/index.ts';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { EMU_PER_INCH, amountText, parseAmount } from './paragraph-dialog-fields.ts';
import { MAX_COLUMN_GAP } from './text-body-commands.ts';

/** Every field of the form, as typed text. `''` is a blank: the selected shapes differ. @public */
export interface ColumnsDialogFields {
  readonly count: string;
  /** Inches. */
  readonly spacing: string;
}

/** A field's range problem: the field, the message key, and its parameters. @internal */
export interface ColumnsFieldError {
  readonly field: keyof ColumnsDialogFields;
  readonly key: TranslationKey;
  readonly params: Readonly<Record<string, number>>;
}

/**
 * Open the form on the selection: blank fields when the selected shapes differ (`undefined`),
 * as the Font and Paragraph dialogs open a field the selection disagrees about. @internal
 */
export function seedColumnsFields(c: TextColumns | undefined): ColumnsDialogFields {
  if (!c) return { count: '', spacing: '' };
  return { count: String(c.count), spacing: amountText(c.gapEmu / EMU_PER_INCH) };
}

/** The first changed field OK cannot apply, or null. The limits are the engine's. @internal */
export function columnsFieldError(
  seed: ColumnsDialogFields,
  d: ColumnsDialogFields
): ColumnsFieldError | null {
  const count = parseAmount(d.count);
  // `setColumns` sets a count, so a spacing change on shapes that differ needs one first.
  const needsCount = d.count !== seed.count || (seed.count === '' && d.spacing !== seed.spacing);
  if (needsCount && !(Number.isInteger(count) && count >= 1 && count <= 16))
    return { field: 'count', key: 'columnsDialog.errors.count', params: { min: 1, max: 16 } };
  const max = Math.floor((MAX_COLUMN_GAP / EMU_PER_INCH) * 100) / 100;
  const spacing = parseAmount(d.spacing);
  if (d.spacing !== seed.spacing && !(Number.isFinite(spacing) && spacing >= 0 && spacing <= max))
    return { field: 'spacing', key: 'columnsDialog.errors.in', params: { min: 0, max } };
  return null;
}

/** The command OK sends: the count, and the spacing when it moved. Null when nothing moved. @internal */
export function columnsCommandOf(
  seed: ColumnsDialogFields,
  d: ColumnsDialogFields
): { readonly type: 'setColumns'; readonly count: number; readonly gapEmu?: number } | null {
  if (d.count === seed.count && d.spacing === seed.spacing) return null;
  return {
    type: 'setColumns',
    count: parseAmount(d.count),
    ...(d.spacing !== seed.spacing && d.spacing !== ''
      ? { gapEmu: Math.round(parseAmount(d.spacing) * EMU_PER_INCH) }
      : {}),
  };
}
