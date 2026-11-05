/** @jsxImportSource react */
// The Columns dialog (Home > Columns > More Columns…): the number of columns
// and the spacing between them, in inches. OK sends one `setColumns` command, so the dialog is
// one undo step. The form model is the engine's (`columns-dialog-fields.ts`), so the Vue twin
// can share it. The host owns visibility (`open` / `onClose`).

import type { ReactNode } from 'react';
import {
  columnsCommandOf,
  columnsDialogError,
  formCommandRejection,
  columnsFieldError,
  seedColumnsFields,
  type ColumnsDialogFields,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import {
  FormDialog,
  NumberField,
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

/** Props for `PptxEditor.ColumnsDialog`. @public */
export interface ColumnsDialogProps extends DialogCustomizationProps {
  /** Whether the dialog is shown. The host owns this state. */
  readonly open: boolean;
  /** Called on Cancel, Escape, a click on the backdrop, and after OK applies. */
  readonly onClose: () => void;
}

/**
 * The field names of the Columns dialog, for `ColumnsDialog.Field`: both fields (`columns`), or
 * one of them, named as its `ColumnsDialogFields` key. @public
 */
export type ColumnsDialogFieldName = 'columns' | 'count' | 'spacing';

/** The Columns draft and its actions, for a host's own field inside the dialog. @public */
export type UseColumnsDialogReturn = UseDialogReturn<ColumnsDialogFields>;

const parts: FormDialogParts<ColumnsDialogFieldName, ColumnsDialogFields> = createDialogParts<
  ColumnsDialogFieldName,
  UseColumnsDialogReturn
>('ColumnsDialog');

function ColumnsDialogRoot({
  open,
  onClose,
  className,
  style,
  preset,
  children,
}: ColumnsDialogProps): ReactNode {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  // Undefined when the selected shapes differ: the fields open blank.
  const current = useEditorState((s) => s.format?.columns);
  const editable = useEditorState((s) => s.status === 'ready' && s.mode === 'edit');
  const { seed, draft, set, error, setError } = useDialogDraft(open, () =>
    seedColumnsFields(current)
  );
  const apply = (): void => {
    if (!draft || !seed) return;
    const problem = columnsFieldError(seed, draft);
    if (problem) return setError(columnsDialogError(t, problem));
    const command = columnsCommandOf(seed, draft);
    if (!command) return onClose();
    const refused = formCommandRejection(t, editor?.exec(command));
    if (refused === null) onClose();
    else setError({ message: refused });
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
      name="columns"
      className={className}
      style={style}
      preset={preset}
      parts={parts}
      state={state}
      hostChildren={children}
      title={t('columnsDialog.title')}
      open={open}
      okLabel={t('columnsDialog.ok')}
      cancelLabel={t('columnsDialog.cancel')}
      closeLabel={t('columnsDialog.close')}
      error={error}
      onClose={onClose}
      onSubmit={apply}
    >
      {draft ? (
        <div className="pptx-dialog__grid" {...fieldPart('columns')}>
          <NumberField
            {...fieldPart('count')}
            name="count"
            label={t('columnsDialog.number')}
            unit=""
            step="1"
            min="1"
            max="16"
            placeholder={seed?.count === '' ? t('columnsDialog.mixed') : undefined}
            autoFocus
            value={draft.count}
            onChange={(v) => set('count', v)}
          />
          <NumberField
            {...fieldPart('spacing')}
            name="spacing"
            label={t('columnsDialog.spacing')}
            unit={t('columnsDialog.inches')}
            step="0.1"
            min="0"
            placeholder={seed?.spacing === '' ? t('columnsDialog.mixed') : undefined}
            value={draft.spacing}
            onChange={(v) => set('spacing', v)}
          />
        </div>
      ) : null}
    </FormDialog>
  );
}

/**
 * Read the draft of the enclosing Columns dialog, for a host's own field inside it. It does not
 * open the dialog: `useEditorDialog('columns').open()` does. @public
 */
export function useColumnsDialog(): UseColumnsDialogReturn {
  return parts.useState();
}

/**
 * Home > Columns > More Columns: the number of text columns of the selected shapes and the
 * spacing between them, applied as one undo step, with replaceable parts (`Header`, `Title`,
 * `Body`, `Footer`, `Field`, `Error`, `Apply`, `Cancel`).
 *
 * @public
 */
export const ColumnsDialog = Object.assign(ColumnsDialogRoot, {
  Header: parts.Header,
  Title: parts.Title,
  Body: parts.Body,
  Footer: parts.Footer,
  Field: parts.Field,
  Error: parts.Error,
  Apply: parts.Apply,
  Cancel: parts.Cancel,
});
