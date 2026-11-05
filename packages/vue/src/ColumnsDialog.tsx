/** @jsxImportSource vue */
// The Columns dialog (Home > Columns > More Columns…): the Vue twin of the React adapter's
// `ColumnsDialog`. The number of columns and the spacing between them, in inches. OK sends one
// `setColumns` command, so the dialog is one undo step. The form model is core's
// (`columns-dialog-fields.ts`), so both adapters show the same fields and send the same command.
// The host owns visibility (`open`, `onClose`).

import {
  defineComponent,
  h,
  type CSSProperties,
  type ExtractPublicPropTypes,
  type PropType,
} from 'vue';
import {
  columnsCommandOf,
  columnsDialogError,
  columnsFieldError,
  seedColumnsFields,
  type ColumnsDialogFields,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, useTranslation } from './composables.ts';
import { createDialogParts, type UseDialogReturn } from './dialog-parts.ts';
import {
  FormDialog,
  NumberField,
  fieldPart,
  formDialogState,
  type FormDialogParts,
} from './dialog-form.tsx';
import { useFormDraft } from './form-draft.ts';

/**
 * The field names of the Columns dialog, for `ColumnsDialog.Field`: both fields (`columns`), or
 * one of them, named as its `ColumnsDialogFields` key. @public
 */
export type ColumnsDialogFieldName = 'columns' | 'count' | 'spacing';

/** The Columns draft and its actions, for a host's own field inside the dialog. @public */
export type UseColumnsDialogReturn = UseDialogReturn<ColumnsDialogFields>;

const parts: FormDialogParts<ColumnsDialogFieldName, ColumnsDialogFields> = createDialogParts<
  UseColumnsDialogReturn,
  ColumnsDialogFieldName
>('ColumnsDialog');

const columnsDialogProps = {
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

/** Props of `ColumnsDialog`, the twin of React's `ColumnsDialogProps`. @public */
export type ColumnsDialogProps = ExtractPublicPropTypes<typeof columnsDialogProps>;

const ColumnsDialogOpen = defineComponent({
  name: 'ColumnsDialogOpen',
  inheritAttrs: false,
  props: columnsDialogProps,
  setup(props, { attrs, slots }) {
    const { t } = useTranslation();
    // Undefined when the selected shapes differ: the fields open blank.
    const current = useEditorState((s) => s.format?.columns);
    const editable = useEditorState((s) => s.status === 'ready' && s.mode === 'edit');
    const seed = seedColumnsFields(current.value);
    const draft = useFormDraft<ColumnsDialogFields>({
      seed,
      problem: (values) => {
        const problem = columnsFieldError(seed, values);
        return problem ? columnsDialogError(t, problem) : null;
      },
      command: (values) => columnsCommandOf(seed, values),
    });
    const state = formDialogState({ draft, isEnabled: editable, onClose: () => props.onClose() });
    return () => {
      const values = draft.values.value;
      const mixed = t('columnsDialog.mixed');
      return (
        <FormDialog
          {...attrs}
          name="columns"
          className={props.className}
          style={props.style}
          preset={props.preset}
          parts={parts}
          state={state as UseDialogReturn<object>}
          title={t('columnsDialog.title')}
          okLabel={t('columnsDialog.ok')}
          cancelLabel={t('columnsDialog.cancel')}
          closeLabel={t('columnsDialog.close')}
        >
          {{
            default: () => (
              <div class="pptx-dialog__grid" {...fieldPart('columns')}>
                <NumberField
                  {...fieldPart('count')}
                  name="count"
                  label={t('columnsDialog.number')}
                  unit=""
                  step="1"
                  min="1"
                  max="16"
                  placeholder={seed.count === '' ? mixed : undefined}
                  autoFocus
                  value={values.count}
                  onChange={(v: string) => draft.setValue('count', v)}
                />
                <NumberField
                  {...fieldPart('spacing')}
                  name="spacing"
                  label={t('columnsDialog.spacing')}
                  unit={t('columnsDialog.inches')}
                  step="0.1"
                  min="0"
                  placeholder={seed.spacing === '' ? mixed : undefined}
                  value={values.spacing}
                  onChange={(v: string) => draft.setValue('spacing', v)}
                />
              </div>
            ),
            host: () => slots.default?.(),
          }}
        </FormDialog>
      );
    };
  },
});

/**
 * Home > Columns > More Columns: the number of text columns of the selected shapes and the
 * spacing between them, applied as one undo step, with replaceable parts (`Header`, `Title`,
 * `Body`, `Footer`, `Field`, `Error`, `Apply`, `Cancel`), as React's `ColumnsDialog`. Renders
 * nothing while `open` is false. @public
 */
export const ColumnsDialog = Object.assign(
  defineComponent({
    name: 'ColumnsDialog',
    inheritAttrs: false,
    props: columnsDialogProps,
    setup(props, { attrs, slots }) {
      return () => (props.open ? h(ColumnsDialogOpen, { ...attrs, ...props }, slots) : null);
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
 * Read the draft of the enclosing Columns dialog, for a host's own field inside it. It does not
 * open the dialog: `useEditorDialog('columns').open()` does. @public
 */
export function useColumnsDialog(): UseColumnsDialogReturn {
  return parts.useState();
}
