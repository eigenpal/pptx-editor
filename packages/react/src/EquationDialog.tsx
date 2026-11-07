/** @jsxImportSource react */
// The Equation dialog (Insert > Equation): a gallery of common equations and the linear form of
// the equation. With an equation selected whose math the linear form can write, it edits that
// equation; otherwise OK inserts a new one. Built on the dialog parts (`dialog-parts.tsx`): the
// field writes a draft, OK applies it as one undo step, and Cancel discards it. The draft rules
// live in core (`equation-fields.ts`), so the Vue twin applies the same command.

import { useId, useMemo, useState, type ReactNode } from 'react';
import {
  EQUATION_GALLERY,
  MAX_LINEAR_LENGTH,
  commitDialogCommand,
  equationDialogEdits,
  equationDraftCommand,
  equationDraftFrom,
  translateDialogErrors,
  translateKeyOrText,
  validateEquationDraft,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { ToolButton } from './chrome.tsx';
import {
  DialogFrame,
  applyOnEnter,
  createDialogParts,
  useCloseOnNewDocument,
  useDialogDraft,
  type DialogCustomizationProps,
  type UseDialogReturn,
} from './dialog-parts.tsx';

/** The draft of the Equation dialog. @public */
export interface EquationDialogFields {
  /** The linear form of the equation, such as `x=(-b\pm\sqrt(b^2-4ac))/(2a)`. */
  readonly linear: string;
}

/** The Equation draft and its actions. @public */
export type UseEquationDialogReturn = UseDialogReturn<EquationDialogFields>;

/** The field names of the Equation dialog, for `EquationDialog.Field`. @public */
export type EquationDialogFieldName = 'gallery' | 'linear';

/** Props for `PptxEditor.EquationDialog`. @public */
export interface EquationDialogProps extends DialogCustomizationProps {
  /** Whether the dialog is shown. The host owns this state. */
  readonly open: boolean;
  /** Called on Cancel, Escape, a click on the backdrop, and after OK. */
  readonly onClose: () => void;
}

const parts = createDialogParts<EquationDialogFieldName, UseEquationDialogReturn>('EquationDialog');

function EquationDialogRoot(props: EquationDialogProps): ReactNode {
  return props.open ? <EquationDialogOpen {...props} /> : null;
}

function EquationDialogOpen({
  onClose,
  className,
  style,
  preset = true,
  children,
}: EquationDialogProps): ReactNode {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const id = useId();
  // The equation the dialog opened on: OK edits it, even when the selection moves meanwhile.
  const current = useEditorState((s) => s.equation);
  const [selected] = useState(() => current);
  const enabled = useEditorState((s) => s.status === 'ready' && s.mode === 'edit');
  const edits = equationDialogEdits(selected);
  useCloseOnNewDocument(onClose);
  const draft = useDialogDraft<EquationDialogFields>(() => ({
    initial: equationDraftFrom(selected),
    validate: validateEquationDraft,
    commit: (values) => commitDialogCommand(editor, equationDraftCommand(values, selected)),
  }));
  const translatedErrors = useMemo(() => translateDialogErrors(t, draft.errors), [t, draft.errors]);
  const state: UseEquationDialogReturn = {
    values: draft.values,
    setValue: draft.setValue,
    errors: translatedErrors,
    isEnabled: enabled,
    apply: () => {
      if (draft.apply()) return onClose();
      // Focus the field with the error (WCAG 3.3.1).
      if (validateEquationDraft(draft.values).linear)
        document.getElementById(`${id}-linear`)?.focus();
    },
    cancel: onClose,
  };
  const error = draft.errors.linear
    ? translateKeyOrText(t, draft.errors.linear)
    : draft.errors.form
      ? translateKeyOrText(t, draft.errors.form)
      : null;
  const title = t(edits ? 'equation.editTitle' : 'equation.insertTitle');
  return (
    <DialogFrame
      kind="equation"
      label={title}
      className={className}
      style={style}
      onClose={onClose}
      onKeyDown={(event) => applyOnEnter(event, state.isEnabled, state.apply)}
    >
      <parts.Composition
        state={state}
        preset={preset}
        defaults={
          <>
            <div data-pptx-part="header" className="pptx-dialog__header">
              <h2 data-pptx-part="title" className="pptx-dialog__title">
                {title}
              </h2>
              <ToolButton label={t('equation.close')} icon="close" onClick={onClose} />
            </div>
            <div data-pptx-part="body" className="pptx-dialog__body">
              {selected && !edits ? (
                <p className="pptx-dialog__muted" role="note">
                  {t('equation.notEditable')}
                </p>
              ) : null}
              <section
                data-pptx-part="field"
                data-pptx-field="gallery"
                className="pptx-dialog__section"
                aria-labelledby={`${id}-gallery`}
              >
                <h3 className="pptx-dialog__label" id={`${id}-gallery`}>
                  {t('equation.gallery')}
                </h3>
                <div className="pptx-equation-gallery">
                  {EQUATION_GALLERY.map((item) => (
                    <button
                      key={item.id}
                      type="button"
                      className="pptx-button pptx-equation-gallery__item"
                      aria-pressed={draft.values.linear === item.linear}
                      onClick={() => draft.setValue('linear', item.linear)}
                    >
                      <span>{t(item.labelKey)}</span>
                      <code className="pptx-dialog__muted">{item.linear}</code>
                    </button>
                  ))}
                </div>
              </section>
              <section
                data-pptx-part="field"
                data-pptx-field="linear"
                className="pptx-dialog__section pptx-field"
              >
                <label className="pptx-dialog__label" htmlFor={`${id}-linear`}>
                  {t('equation.linear')}
                </label>
                <p className="pptx-dialog__muted" id={`${id}-help`}>
                  {t('equation.linearHelp')}
                </p>
                <input
                  id={`${id}-linear`}
                  type="text"
                  spellCheck={false}
                  autoComplete="off"
                  maxLength={MAX_LINEAR_LENGTH}
                  value={draft.values.linear}
                  aria-invalid={draft.errors.linear ? true : false}
                  aria-describedby={draft.errors.linear ? `${id}-help ${id}-error` : `${id}-help`}
                  onChange={(e) => draft.setValue('linear', e.target.value)}
                />
              </section>
            </div>
            <div data-pptx-part="footer" className="pptx-dialog__footer">
              <span
                id={`${id}-error`}
                data-pptx-part="error"
                role="alert"
                className="pptx-dialog__error"
              >
                {error}
              </span>
              <button
                type="button"
                data-pptx-part="cancel"
                className="pptx-button"
                onClick={state.cancel}
              >
                {t('dialog.cancel')}
              </button>
              <button
                type="button"
                data-pptx-part="apply"
                className="pptx-button pptx-button--primary"
                disabled={!state.isEnabled}
                onClick={state.apply}
              >
                {edits ? t('dialog.ok') : t('equation.insert')}
              </button>
            </div>
          </>
        }
      >
        {children}
      </parts.Composition>
    </DialogFrame>
  );
}

/**
 * Read the draft of the enclosing Equation dialog, for a host's own field inside it. It does
 * not open the dialog: `useEditorDialog('equation').open()` does. @public
 */
export function useEquationDialog(): UseEquationDialogReturn {
  return parts.useState();
}

/**
 * Insert Equation and Edit Equation: a gallery of common equations and the linear form, with
 * replaceable parts (`Header`, `Title`, `Body`, `Footer`, `Field`, `Error`, `Apply`, `Cancel`).
 *
 * @example
 * ```tsx
 * <PptxEditor.EquationDialog {...props}>
 *   <PptxEditor.EquationDialog.Field name="gallery" hidden />
 * </PptxEditor.EquationDialog>
 * ```
 *
 * @public
 */
export const EquationDialog = Object.assign(EquationDialogRoot, {
  Header: parts.Header,
  Title: parts.Title,
  Body: parts.Body,
  Footer: parts.Footer,
  Field: parts.Field,
  Error: parts.Error,
  Apply: parts.Apply,
  Cancel: parts.Cancel,
});
