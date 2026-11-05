/** @jsxImportSource react */
// The Alt Text dialog (Picture Format → Alt Text, and Edit Alt Text in the context menu): the
// description and title of the selected element, and "Mark as decorative", as in the
// Alt Text pane. The DOCX Editor equivalent is the `imageAltText` popup. Built on the dialog
// parts (`dialog-parts.tsx`): the fields write a draft, OK applies it as one undo step, and
// Cancel discards it. The draft rules live in core (`picture-fields.ts`), so the Vue twin
// applies the same command.

import { useEffect, useId, useMemo, useRef, useState, type ReactNode } from 'react';
import {
  MAX_ALT_TEXT,
  altTextDraftCommand,
  altTextDraftFrom,
  commitDialogCommand,
  translateDialogErrors,
  translateKeyOrText,
  validateAltTextDraft,
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

/** The draft of the Alt Text dialog. @public */
export interface ImageAltTextDialogFields {
  /** The description a screen reader reads (`p:cNvPr/@descr`). */
  readonly description: string;
  /** The title (`p:cNvPr/@title`). */
  readonly title: string;
  /** Mark as decorative: a screen reader skips the element, and the description is dropped. */
  readonly decorative: boolean;
}

/** The Alt Text draft and its actions. @public */
export type UseImageAltTextDialogReturn = UseDialogReturn<ImageAltTextDialogFields>;

/** The field names of the Alt Text dialog, for `ImageAltTextDialog.Field`. @public */
export type ImageAltTextDialogFieldName = 'description' | 'title' | 'decorative';

/** Props for `PptxEditor.ImageAltTextDialog`. @public */
export interface ImageAltTextDialogProps extends DialogCustomizationProps {
  /** Whether the dialog is shown. The host owns this state. */
  readonly open: boolean;
  /** Called on Cancel, Escape, a click on the backdrop, and after OK. */
  readonly onClose: () => void;
}

const parts = createDialogParts<ImageAltTextDialogFieldName, UseImageAltTextDialogReturn>(
  'ImageAltTextDialog'
);

function ImageAltTextDialogRoot(props: ImageAltTextDialogProps): ReactNode {
  return props.open ? <ImageAltTextDialogOpen {...props} /> : null;
}

function ImageAltTextDialogOpen({
  onClose,
  className,
  style,
  preset = true,
  children,
}: ImageAltTextDialogProps): ReactNode {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const id = useId();
  const current = useEditorState((s) => s.geometry?.altText);
  const enabled = useEditorState(
    (s) => s.status === 'ready' && s.mode === 'edit' && s.geometry !== null
  );
  const [initial] = useState(() => altTextDraftFrom(current));
  // The draft belongs to the object the dialog opened on: another selection, or the object's
  // removal, closes the dialog, so OK never applies one object's draft to another.
  const nodeId = useEditorState((s) => s.geometry?.nodeId ?? null);
  const [openedOn] = useState(nodeId);
  const closeRef = useRef(onClose);
  closeRef.current = onClose;
  useEffect(() => {
    if (nodeId !== openedOn) closeRef.current();
  }, [nodeId, openedOn]);
  useCloseOnNewDocument(onClose);
  const draft = useDialogDraft<ImageAltTextDialogFields>(() => ({
    initial,
    validate: validateAltTextDraft,
    commit: (values) => commitDialogCommand(editor, altTextDraftCommand(values, initial)),
  }));
  // The state gives messages; the raw keys stay inside, for focus.
  const translatedErrors = useMemo(() => translateDialogErrors(t, draft.errors), [t, draft.errors]);
  const state: UseImageAltTextDialogReturn = {
    values: draft.values,
    setValue: draft.setValue,
    errors: translatedErrors,
    isEnabled: enabled,
    apply: () => {
      if (draft.apply()) return onClose();
      // Focus the field with the error (WCAG 3.3.1), as Format Picture does.
      const bad = (['description', 'title'] as const).find(
        (f) => validateAltTextDraft(draft.values)[f]
      );
      if (bad) document.getElementById(`${id}-${bad}`)?.focus();
    },
    cancel: onClose,
  };
  const { description, title, decorative } = draft.values;
  // The first error, with the name of its field.
  const bad = draft.errors.description ? 'description' : draft.errors.title ? 'title' : null;
  const error = bad
    ? t('imageAltText.fieldError', {
        field: t(bad === 'description' ? 'imageAltText.description' : 'imageAltText.titleLabel'),
        message: translateKeyOrText(t, draft.errors[bad]!),
      })
    : draft.errors.form
      ? translateKeyOrText(t, draft.errors.form)
      : null;
  const invalid = (f: 'description' | 'title') =>
    draft.errors[f] ? { 'aria-invalid': true as const } : { 'aria-invalid': false as const };
  return (
    <DialogFrame
      kind="imageAltText"
      label={t('imageAltText.title')}
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
                {t('imageAltText.title')}
              </h2>
              <ToolButton label={t('imageAltText.close')} icon="close" onClick={onClose} />
            </div>
            <div data-pptx-part="body" className="pptx-dialog__body">
              <section
                data-pptx-part="field"
                data-pptx-field="description"
                className="pptx-dialog__section"
              >
                <label className="pptx-dialog__label" htmlFor={`${id}-description`}>
                  {t('imageAltText.description')}
                </label>
                <p className="pptx-dialog__muted" id={`${id}-help`}>
                  {t('imageAltText.descriptionHelp')}
                </p>
                <textarea
                  id={`${id}-description`}
                  className="pptx-dialog__textarea"
                  rows={5}
                  value={description}
                  disabled={decorative}
                  {...invalid('description')}
                  aria-describedby={[
                    `${id}-help`,
                    `${id}-count`,
                    decorative ? `${id}-decorative-reason` : null,
                    bad === 'description' ? `${id}-error` : null,
                  ]
                    .filter(Boolean)
                    .join(' ')}
                  onChange={(e) => draft.setValue('description', e.target.value)}
                />
                {decorative ? (
                  <span className="pptx-sr-only" id={`${id}-decorative-reason`}>
                    {t('imageAltText.decorativeReason')}
                  </span>
                ) : null}
                <span className="pptx-dialog__muted" id={`${id}-count`}>
                  {t('imageAltText.count', { count: description.length, max: MAX_ALT_TEXT })}
                </span>
              </section>
              <section
                data-pptx-part="field"
                data-pptx-field="title"
                className="pptx-dialog__section pptx-field"
              >
                <label className="pptx-dialog__label" htmlFor={`${id}-title`}>
                  {t('imageAltText.titleLabel')}
                </label>
                <input
                  id={`${id}-title`}
                  type="text"
                  value={title}
                  {...invalid('title')}
                  aria-describedby={bad === 'title' ? `${id}-error` : undefined}
                  onChange={(e) => draft.setValue('title', e.target.value)}
                />
              </section>
              <section
                data-pptx-part="field"
                data-pptx-field="decorative"
                className="pptx-dialog__section"
              >
                <label className="pptx-radio">
                  <input
                    type="checkbox"
                    checked={decorative}
                    aria-describedby={`${id}-decorative-help`}
                    onChange={(e) => draft.setValue('decorative', e.target.checked)}
                  />
                  {t('imageAltText.decorative')}
                </label>
                <p className="pptx-dialog__muted" id={`${id}-decorative-help`}>
                  {t('imageAltText.decorativeHelp')}
                </p>
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
                {t('dialog.ok')}
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
 * Read the draft of the enclosing Alt Text dialog, for a host's own field inside it. It does
 * not open the dialog: `useEditorDialog('imageAltText').open()` does. @public
 */
export function useImageAltTextDialog(): UseImageAltTextDialogReturn {
  return parts.useState();
}

/**
 * Alt text of the selected element: description, title, and "Mark as decorative", with
 * replaceable parts (`Header`, `Title`, `Body`, `Footer`, `Field`, `Error`, `Apply`, `Cancel`).
 *
 * @example
 * ```tsx
 * <PptxEditor.ImageAltTextDialog {...props}>
 *   <PptxEditor.ImageAltTextDialog.Field name="title" hidden />
 * </PptxEditor.ImageAltTextDialog>
 * ```
 *
 * @public
 */
export const ImageAltTextDialog = Object.assign(ImageAltTextDialogRoot, {
  Header: parts.Header,
  Title: parts.Title,
  Body: parts.Body,
  Footer: parts.Footer,
  Field: parts.Field,
  Error: parts.Error,
  Apply: parts.Apply,
  Cancel: parts.Cancel,
});
