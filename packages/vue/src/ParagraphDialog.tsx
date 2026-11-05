/** @jsxImportSource vue */
// The Paragraph dialog (Format > Paragraph…, Line Spacing Options…, or the text context menu):
// the Vue twin of the React adapter's `ParagraphDialog`. Alignment, indentation before text
// with a first-line or hanging indent, spacing before and after, and line spacing. The fields
// seed from the selection when the dialog opens; a field the selected paragraphs disagree about
// opens blank. OK sends only the fields that changed as ONE `setParagraphFormat` command, so the
// dialog is one undo step. The form model is core's (`paragraph-dialog-fields.ts`). The host
// owns visibility (`open`, `onClose`).

import {
  defineComponent,
  h,
  type CSSProperties,
  type ExtractPublicPropTypes,
  type PropType,
} from 'vue';
import {
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
import { useEditorState, useTranslation } from './composables.ts';
import { createDialogParts, type UseDialogReturn } from './dialog-parts.ts';
import {
  FormDialog,
  FormSection,
  NumberField,
  SelectField,
  fieldPart,
  formDialogState,
  type FormDialogParts,
} from './dialog-form.tsx';
import { useFormDraft } from './form-draft.ts';

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
  UseParagraphDialogReturn,
  ParagraphDialogFieldName
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

const paragraphDialogProps = {
  /** Whether the dialog is shown. The host owns this state. */
  open: { type: Boolean, required: true as const },
  /** Called on Cancel, Escape, a click on the backdrop, and after OK applies. */
  onClose: { type: Function as PropType<() => void>, required: true as const },
  /** `false`: only the default slot, with no default arrangement. */
  preset: { type: Boolean, default: true },
  /** A class name added to the dialog surface, as React's `className` (`class` works too). */
  className: { type: String, default: undefined },
  /** Styles added to the dialog surface. */
  style: { type: Object as PropType<CSSProperties>, default: undefined },
} as const;

/** Props of `ParagraphDialog`, the twin of React's `ParagraphDialogProps`. @public */
export type ParagraphDialogProps = ExtractPublicPropTypes<typeof paragraphDialogProps>;

const ParagraphDialogOpen = defineComponent({
  name: 'ParagraphDialogOpen',
  inheritAttrs: false,
  props: paragraphDialogProps,
  setup(props, { attrs, slots }) {
    const { t } = useTranslation();
    const current = useEditorState((s) => s.format?.paragraph ?? null);
    const editable = useEditorState((s) => s.status === 'ready' && s.mode === 'edit');
    const seed = seedParagraphFields(current.value);
    const draft = useFormDraft<ParagraphDialogFields>({
      seed,
      problem: (values) => {
        const problem = paragraphFieldError(seed, values);
        return problem ? paragraphDialogError(t, problem) : null;
      },
      command: (values) => {
        const command = paragraphCommandOf(seed, values);
        return command ? { type: 'setParagraphFormat', ...command } : null;
      },
    });
    const state = formDialogState({ draft, isEnabled: editable, onClose: () => props.onClose() });
    const set = draft.setValue;
    const spacing = (side: 'Before' | 'After', d: ParagraphDialogFields, mixed: string) => {
      const key = `space${side}` as const;
      const unitKey = `space${side}Unit` as const;
      return (
        <div {...fieldPart(key)} class="pptx-field__pair">
          <NumberField
            name={key}
            label={t(`paragraphDialog.fields.${key}`)}
            unit=""
            step={d[unitKey] === 'pt' ? '6' : '0.1'}
            min="0"
            placeholder={seed[key] === '' ? mixed : undefined}
            value={d[key]}
            onChange={(v: string) => set(key, v)}
          />
          <SelectField
            name={unitKey}
            label={t('paragraphDialog.unitOf', { field: t(`paragraphDialog.fields.${key}`) })}
            hideLabel
            value={d[unitKey]}
            options={(['pt', 'lines'] as const).map(
              (u) => [u, t(`paragraphDialog.units.${u}`)] as const
            )}
            onChange={(u: SpacingUnit) => draft.update((cur) => withSpaceUnit(cur, seed, key, u))}
          />
        </div>
      );
    };
    return () => {
      const d = draft.values.value;
      const mixed = t('paragraphDialog.mixed');
      const inch = t('paragraphDialog.units.in');
      return (
        <FormDialog
          {...attrs}
          name="paragraph"
          className={props.className}
          style={props.style}
          preset={props.preset}
          parts={parts}
          state={state as UseDialogReturn<object>}
          title={t('paragraphDialog.title')}
          okLabel={t('paragraphDialog.ok')}
          cancelLabel={t('paragraphDialog.cancel')}
          closeLabel={t('paragraphDialog.close')}
        >
          {{
            default: () => [
              <FormSection {...fieldPart('general')} title={t('paragraphDialog.general')}>
                {{
                  default: () => (
                    <>
                      <SelectField
                        {...fieldPart('alignment')}
                        name="alignment"
                        label={t('paragraphDialog.fields.alignment')}
                        mixed={mixed}
                        value={d.alignment}
                        options={ALIGNMENTS.map(
                          (a) => [a, t(`paragraphDialog.align.${a}`)] as const
                        )}
                        onChange={(v: ParagraphDialogFields['alignment']) => set('alignment', v)}
                      />
                    </>
                  ),
                }}
              </FormSection>,
              <FormSection {...fieldPart('indentation')} title={t('paragraphDialog.indentation')}>
                {{
                  default: () => (
                    <>
                      <div class="pptx-dialog__grid">
                        <NumberField
                          {...fieldPart('indentLeft')}
                          name="indentLeft"
                          label={t('paragraphDialog.fields.indentLeft')}
                          unit={inch}
                          step="0.1"
                          min="0"
                          placeholder={seed.indentLeft === '' ? mixed : undefined}
                          value={d.indentLeft}
                          onChange={(v: string) => set('indentLeft', v)}
                        />
                        <SelectField
                          {...fieldPart('special')}
                          name="special"
                          label={t('paragraphDialog.fields.special')}
                          mixed={mixed}
                          value={d.special}
                          options={SPECIALS.map(
                            (s) => [s, t(`paragraphDialog.specials.${s}`)] as const
                          )}
                          onChange={(v: SpecialIndent) =>
                            draft.update((cur) => ({
                              ...cur,
                              special: v,
                              specialBy: specialByFor(v, cur),
                            }))
                          }
                        />
                        <NumberField
                          {...fieldPart('specialBy')}
                          name="specialBy"
                          label={t('paragraphDialog.fields.specialBy')}
                          unit={inch}
                          step="0.1"
                          min="0"
                          placeholder={seed.specialBy === '' ? mixed : undefined}
                          disabled={d.special === 'none' || d.special === ''}
                          value={d.specialBy}
                          onChange={(v: string) => set('specialBy', v)}
                        />
                      </div>
                    </>
                  ),
                }}
              </FormSection>,
              <FormSection {...fieldPart('spacing')} title={t('paragraphDialog.spacing')}>
                {{
                  default: () => (
                    <>
                      <div class="pptx-dialog__rows">
                        {spacing('Before', d, mixed)}
                        {spacing('After', d, mixed)}
                        <div class="pptx-field__pair">
                          <SelectField
                            {...fieldPart('lineRule')}
                            name="lineRule"
                            label={t('paragraphDialog.fields.lineRule')}
                            mixed={mixed}
                            value={d.lineRule}
                            options={LINE_RULES.map(
                              (r) => [r, t(`paragraphDialog.lineRules.${r}`)] as const
                            )}
                            onChange={(rule: LineRule) =>
                              draft.update((cur) => ({
                                ...cur,
                                lineRule: rule,
                                lineAt: lineAtFor(rule, cur),
                              }))
                            }
                          />
                          <NumberField
                            {...fieldPart('lineAt')}
                            name="lineAt"
                            label={t('paragraphDialog.fields.lineAt')}
                            unit={d.lineRule === 'exact' ? t('paragraphDialog.units.pt') : ''}
                            step={d.lineRule === 'exact' ? '1' : '0.1'}
                            min="0"
                            placeholder={seed.lineAt === '' ? mixed : undefined}
                            disabled={d.lineRule !== 'exact' && d.lineRule !== 'multiple'}
                            value={d.lineAt}
                            onChange={(v: string) => set('lineAt', v)}
                          />
                        </div>
                      </div>
                    </>
                  ),
                }}
              </FormSection>,
            ],
            host: () => slots.default?.(),
          }}
        </FormDialog>
      );
    };
  },
});

/**
 * Format > Paragraph: alignment, indents, spacing, and line spacing of the selected paragraphs,
 * applied as one undo step, with replaceable parts (`Header`, `Title`, `Body`, `Footer`,
 * `Field`, `Error`, `Apply`, `Cancel`), as React's `ParagraphDialog`. Renders nothing while
 * `open` is false. @public
 */
export const ParagraphDialog = Object.assign(
  defineComponent({
    name: 'ParagraphDialog',
    inheritAttrs: false,
    props: paragraphDialogProps,
    setup(props, { attrs, slots }) {
      return () => (props.open ? h(ParagraphDialogOpen, { ...attrs, ...props }, slots) : null);
    },
  }),
  {
    Header: parts.Header,
    Title: parts.Title,
    Body: parts.Body,
    Footer: parts.Footer,
    Field: parts.Field,
    Error: parts.Error,
    Apply: parts.Apply,
    Cancel: parts.Cancel,
  }
);

/**
 * Read the draft of the enclosing Paragraph dialog, for a host's own field inside it. It does
 * not open the dialog: `useEditorDialog('paragraph').open()` does. @public
 */
export function useParagraphDialog(): UseParagraphDialogReturn {
  return parts.useState();
}
