/** @jsxImportSource react */
// The Paragraph dialog (Format > Paragraph…, Line Spacing Options…, or the text context menu),
// with alignment, indentation before text with a first-line or hanging indent,
// spacing before and after, and line spacing. The fields seed from the selection when the dialog
// opens; a field the selected paragraphs disagree about opens blank. OK sends only the fields
// that changed as ONE `setParagraphFormat` command, so the dialog is one undo step (DOCX
// Editor's `DocxEditorParagraphDialog` does the same). The form model is the engine's
// (`paragraph-dialog-fields.ts`), so the Vue twin can share it. The host owns visibility.

import {
  formCommandRejection,
  lineAtFor,
  paragraphCommandOf,
  paragraphDialogError,
  paragraphFieldError,
  seedParagraphFields,
  specialByFor,
  withSpaceUnit,
  type LineRule,
  type ParagraphAlignment,
  type ParagraphDialogFields,
  type SpacingUnit,
  type SpecialIndent,
} from '@pptx-editor.dev/core/editor';
import type { ReactNode } from 'react';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import {
  FormDialog,
  FormSection,
  NumberField,
  SelectField,
  fieldPart,
  formDialogState,
  useDialogDraft,
  type FormDialogParts,
} from './dialog-form.tsx';
import {
  createDialogParts,
  type DialogCustomizationProps,
  type UseDialogReturn,
} from './dialog-parts.tsx';

/** Props for `PptxEditor.ParagraphDialog`. @public */
export interface ParagraphDialogProps extends DialogCustomizationProps {
  /** Whether the dialog is shown. The host owns this state. */
  readonly open: boolean;
  /** Called on Cancel, Escape, a click on the backdrop, and after OK applies. */
  readonly onClose: () => void;
}

/**
 * The field names of the Paragraph dialog, for `ParagraphDialog.Field`: a section (`general`,
 * `indentation`, `spacing`) or one field in it, named as its `ParagraphDialogFields` key. A
 * spacing field holds its unit list. @public
 */
export type ParagraphDialogFieldName =
  | 'general'
  | 'indentation'
  | 'spacing'
  | 'alignment'
  | 'indentLeft'
  | 'special'
  | 'specialBy'
  | 'spaceBefore'
  | 'spaceAfter'
  | 'lineRule'
  | 'lineAt';

/** The Paragraph draft and its actions, for a host's own field inside the dialog. @public */
export type UseParagraphDialogReturn = UseDialogReturn<ParagraphDialogFields>;

const parts: FormDialogParts<ParagraphDialogFieldName, ParagraphDialogFields> = createDialogParts<
  ParagraphDialogFieldName,
  UseParagraphDialogReturn
>('ParagraphDialog');

const ALIGNMENTS: readonly ParagraphAlignment[] = [
  'left',
  'center',
  'right',
  'justify',
  'distributed',
];
const SPECIALS: readonly SpecialIndent[] = ['none', 'firstLine', 'hanging'];
const LINE_RULES: readonly LineRule[] = ['single', 'onePointFive', 'double', 'exact', 'multiple'];

function ParagraphDialogRoot({
  open,
  onClose,
  className,
  style,
  preset,
  children,
}: ParagraphDialogProps): ReactNode {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const current = useEditorState((s) => s.format?.paragraph ?? null);
  const editable = useEditorState((s) => s.status === 'ready' && s.mode === 'edit');
  const { seed, draft, set, update, error, setError } = useDialogDraft(open, () =>
    seedParagraphFields(current)
  );
  const apply = (): void => {
    if (!draft || !seed) return;
    const problem = paragraphFieldError(seed, draft);
    if (problem) return setError(paragraphDialogError(t, problem));
    const command = paragraphCommandOf(seed, draft);
    if (!command) return onClose();
    const refused = formCommandRejection(
      t,
      editor?.exec({ type: 'setParagraphFormat', ...command })
    );
    if (refused === null) onClose();
    else setError({ message: refused });
  };
  const mixed = t('paragraphDialog.mixed');
  const inch = t('paragraphDialog.units.in');
  const choose = <K extends keyof ParagraphDialogFields>(k: K, v: ParagraphDialogFields[K]) =>
    set(k, v);
  const spacing = (side: 'Before' | 'After') => {
    if (!draft) return null;
    const key = `space${side}` as const;
    const unitKey = `space${side}Unit` as const;
    return (
      <div {...fieldPart(key)} className="pptx-field__pair">
        <NumberField
          name={key}
          label={t(`paragraphDialog.fields.${key}`)}
          unit=""
          step={draft[unitKey] === 'pt' ? '6' : '0.1'}
          min="0"
          placeholder={seed?.[key] === '' ? mixed : undefined}
          value={draft[key]}
          onChange={(v) => set(key, v)}
        />
        <SelectField<SpacingUnit>
          name={unitKey}
          label={t('paragraphDialog.unitOf', { field: t(`paragraphDialog.fields.${key}`) })}
          hideLabel
          value={draft[unitKey]}
          options={(['pt', 'lines'] as const).map((u) => [u, t(`paragraphDialog.units.${u}`)])}
          onChange={(u) => seed && update((d) => withSpaceUnit(d, seed, key, u))}
        />
      </div>
    );
  };
  const state = formDialogState({
    draft,
    set,
    error,
    isEnabled: editable,
    onSubmit: apply,
    onClose,
  });
  return (
    <FormDialog
      name="paragraph"
      className={className}
      style={style}
      preset={preset}
      parts={parts}
      state={state}
      hostChildren={children}
      title={t('paragraphDialog.title')}
      open={open}
      okLabel={t('paragraphDialog.ok')}
      cancelLabel={t('paragraphDialog.cancel')}
      closeLabel={t('paragraphDialog.close')}
      error={error}
      onClose={onClose}
      onSubmit={apply}
    >
      {draft ? (
        <>
          <FormSection {...fieldPart('general')} title={t('paragraphDialog.general')}>
            <SelectField
              {...fieldPart('alignment')}
              name="alignment"
              label={t('paragraphDialog.fields.alignment')}
              mixed={mixed}
              value={draft.alignment}
              options={ALIGNMENTS.map((a) => [a, t(`paragraphDialog.align.${a}`)])}
              onChange={(v) => choose('alignment', v)}
            />
          </FormSection>
          <FormSection {...fieldPart('indentation')} title={t('paragraphDialog.indentation')}>
            <div className="pptx-dialog__grid">
              <NumberField
                {...fieldPart('indentLeft')}
                name="indentLeft"
                label={t('paragraphDialog.fields.indentLeft')}
                unit={inch}
                step="0.1"
                min="0"
                placeholder={seed?.indentLeft === '' ? mixed : undefined}
                value={draft.indentLeft}
                onChange={(v) => set('indentLeft', v)}
              />
              <SelectField
                {...fieldPart('special')}
                name="special"
                label={t('paragraphDialog.fields.special')}
                mixed={mixed}
                value={draft.special}
                options={SPECIALS.map((s) => [s, t(`paragraphDialog.specials.${s}`)])}
                onChange={(v) => {
                  set('special', v);
                  set('specialBy', specialByFor(v, draft));
                }}
              />
              <NumberField
                {...fieldPart('specialBy')}
                name="specialBy"
                label={t('paragraphDialog.fields.specialBy')}
                unit={inch}
                step="0.1"
                min="0"
                placeholder={seed?.specialBy === '' ? mixed : undefined}
                disabled={draft.special === 'none' || draft.special === ''}
                value={draft.specialBy}
                onChange={(v) => set('specialBy', v)}
              />
            </div>
          </FormSection>
          <FormSection {...fieldPart('spacing')} title={t('paragraphDialog.spacing')}>
            <div className="pptx-dialog__rows">
              {spacing('Before')}
              {spacing('After')}
              <div className="pptx-field__pair">
                <SelectField
                  {...fieldPart('lineRule')}
                  name="lineRule"
                  label={t('paragraphDialog.fields.lineRule')}
                  mixed={mixed}
                  value={draft.lineRule}
                  options={LINE_RULES.map((r) => [r, t(`paragraphDialog.lineRules.${r}`)])}
                  onChange={(rule) => {
                    set('lineRule', rule);
                    set('lineAt', lineAtFor(rule, draft));
                  }}
                />
                <NumberField
                  {...fieldPart('lineAt')}
                  name="lineAt"
                  label={t('paragraphDialog.fields.lineAt')}
                  unit={draft.lineRule === 'exact' ? t('paragraphDialog.units.pt') : ''}
                  step={draft.lineRule === 'exact' ? '1' : '0.1'}
                  min="0"
                  placeholder={seed?.lineAt === '' ? mixed : undefined}
                  disabled={draft.lineRule !== 'exact' && draft.lineRule !== 'multiple'}
                  value={draft.lineAt}
                  onChange={(v) => set('lineAt', v)}
                />
              </div>
            </div>
          </FormSection>
        </>
      ) : null}
    </FormDialog>
  );
}

/**
 * Read the draft of the enclosing Paragraph dialog, for a host's own field inside it. It does
 * not open the dialog: `useEditorDialog('paragraph').open()` does. @public
 */
export function useParagraphDialog(): UseParagraphDialogReturn {
  return parts.useState();
}

/**
 * Format > Paragraph: alignment, indents, spacing, and line spacing of the selected paragraphs,
 * applied as one undo step, with replaceable parts (`Header`, `Title`, `Body`, `Footer`,
 * `Field`, `Error`, `Apply`, `Cancel`).
 *
 * @example
 * ```tsx
 * <PptxEditor.ParagraphDialog {...props}>
 *   <PptxEditor.ParagraphDialog.Field name="indentation" hidden />
 * </PptxEditor.ParagraphDialog>
 * ```
 *
 * @public
 */
export const ParagraphDialog = Object.assign(ParagraphDialogRoot, {
  Header: parts.Header,
  Title: parts.Title,
  Body: parts.Body,
  Footer: parts.Footer,
  Field: parts.Field,
  Error: parts.Error,
  Apply: parts.Apply,
  Cancel: parts.Cancel,
});
