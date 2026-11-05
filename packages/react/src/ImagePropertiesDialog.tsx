/** @jsxImportSource react */
// The Format Picture dialog (`image.properties`, Format > Format picture, and the canvas context
// menu): the size, rotation, and crop position of the selected picture, as DOCX Editor's
// `imageProperties` dialog. Built on the dialog parts (`dialog-parts.tsx`): the fields write a
// draft, OK applies it as one `setImageProperties` (one undo step), and Cancel discards it.
// The draft rules live in core (`picture-fields.ts`), so the Vue twin applies the same command.

import { useEffect, useId, useMemo, useRef, useState, type ReactNode } from 'react';
import {
  commitDialogCommand,
  picturePropertiesDraftCommand,
  picturePropertiesDraftFrom,
  picturePropertiesWithValue,
  translateDialogErrors,
  translateKeyOrText,
  validatePicturePropertiesDraft,
} from '@pptx-editor.dev/core/editor';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
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

/** The draft of the Format Picture dialog, in inches, degrees, and percent. @public */
export interface ImagePropertiesDialogFields {
  /** Width in inches, on the slide (`SelectedImageState.widthEmu`). */
  readonly width: number;
  /** Height in inches, on the slide (`SelectedImageState.heightEmu`). */
  readonly height: number;
  /** Lock aspect ratio: a width change scales the height, and the other way round. */
  readonly lockAspect: boolean;
  /** Rotation in degrees, clockwise. */
  readonly rotation: number;
  /** Crop of the left edge, in percent of the image; a negative value adds space. */
  readonly cropLeft: number;
  /** Crop of the top edge, in percent of the image. */
  readonly cropTop: number;
  /** Crop of the right edge, in percent of the image. */
  readonly cropRight: number;
  /** Crop of the bottom edge, in percent of the image. */
  readonly cropBottom: number;
}

/** The Format Picture draft and its actions. @public */
export type UseImagePropertiesDialogReturn = UseDialogReturn<ImagePropertiesDialogFields>;

/** The field names of the Format Picture dialog, for `ImagePropertiesDialog.Field`. @public */
export type ImagePropertiesDialogFieldName = 'size' | 'rotation' | 'crop';

/** Props for `PptxEditor.ImagePropertiesDialog`. @public */
export interface ImagePropertiesDialogProps extends DialogCustomizationProps {
  /** Whether the dialog is shown. The host owns this state. */
  readonly open: boolean;
  /** Called on Cancel, Escape, a click on the backdrop, and after OK. */
  readonly onClose: () => void;
}

const parts = createDialogParts<ImagePropertiesDialogFieldName, UseImagePropertiesDialogReturn>(
  'ImagePropertiesDialog'
);

type NumberField = Exclude<keyof ImagePropertiesDialogFields, 'lockAspect'>;

/** The number fields in screen order (Height before Width), for errors and focus. */
const NUMBER_FIELDS: readonly NumberField[] = [
  'height',
  'width',
  'rotation',
  'cropLeft',
  'cropTop',
  'cropRight',
  'cropBottom',
];

/** The label of each number field, to name the field an error is about. */
const FIELD_LABELS: Record<NumberField, TranslationKey> = {
  width: 'imageProperties.width',
  height: 'imageProperties.height',
  rotation: 'imageProperties.rotation',
  cropLeft: 'imageProperties.cropLeft',
  cropTop: 'imageProperties.cropTop',
  cropRight: 'imageProperties.cropRight',
  cropBottom: 'imageProperties.cropBottom',
};

const textsOf = (fields: ImagePropertiesDialogFields): Record<NumberField, string> =>
  Object.fromEntries(NUMBER_FIELDS.map((f) => [f, String(fields[f])])) as Record<
    NumberField,
    string
  >;

function ImagePropertiesDialogRoot(props: ImagePropertiesDialogProps): ReactNode {
  return props.open ? <ImagePropertiesDialogOpen {...props} /> : null;
}

function ImagePropertiesDialogOpen({
  onClose,
  className,
  style,
  preset = true,
  children,
}: ImagePropertiesDialogProps): ReactNode {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const id = useId();
  const picture = useEditorState((s) => s.image);
  // Format Picture edits one picture.
  const enabled = useEditorState(
    (s) => s.status === 'ready' && s.mode === 'edit' && s.image?.count === 1
  );
  const [initial] = useState(() => picturePropertiesDraftFrom(picture));
  const [exactCrop] = useState(() => picture?.crop);
  // The draft belongs to the picture the dialog opened on: another selection, or the picture's
  // removal, closes the dialog, so OK never applies one picture's draft to another.
  const [openedOn] = useState(() => picture?.nodeId ?? null);
  const nodeId = picture?.nodeId ?? null;
  const closeRef = useRef(onClose);
  closeRef.current = onClose;
  useEffect(() => {
    if (nodeId !== openedOn) closeRef.current();
  }, [nodeId, openedOn]);
  useCloseOnNewDocument(onClose);
  const draft = useDialogDraft<ImagePropertiesDialogFields>(() => ({
    initial,
    validate: (values) => validatePicturePropertiesDraft(values, initial),
    commit: (values) =>
      commitDialogCommand(editor, picturePropertiesDraftCommand(values, initial, exactCrop)),
  }));
  // The state gives messages; the raw keys stay inside, for focus.
  const translatedErrors = useMemo(() => translateDialogErrors(t, draft.errors), [t, draft.errors]);
  const state: UseImagePropertiesDialogReturn = {
    values: draft.values,
    // A host part's setValue takes the packaged fields' rules: Lock aspect ratio and the shown text.
    setValue: (name, value) => {
      if (name !== 'lockAspect' && typeof value === 'number') set(name, String(value));
      else draft.setValue(name, value);
    },
    errors: translatedErrors,
    isEnabled: enabled,
    apply: () => {
      if (draft.apply()) return onClose();
      // Focus the first field with an error (WCAG 3.3.1).
      const bad = NUMBER_FIELDS.find(
        (f) => validatePicturePropertiesDraft(draft.values, initial)[f]
      );
      if (bad) document.getElementById(`${id}-${bad}`)?.focus();
    },
    cancel: onClose,
  };
  const values = draft.values;
  // The text of each field as typed; the draft holds the numbers.
  const [texts, setTexts] = useState<Record<NumberField, string>>(() => textsOf(initial));
  const set = (field: NumberField, text: string): void => {
    const value = text.trim() === '' ? NaN : Number(text);
    if (field === 'width' || field === 'height') {
      // The live draft, as in Vue: two calls in one event each see the other's value.
      const current = draft.values;
      const next = picturePropertiesWithValue(current, field, value, initial);
      const other = field === 'width' ? 'height' : 'width';
      draft.setValue('width', next.width);
      draft.setValue('height', next.height);
      setTexts((all) => ({
        ...all,
        [field]: text,
        ...(next[other] !== current[other] ? { [other]: String(next[other]) } : {}),
      }));
    } else {
      draft.setValue(field, value);
      setTexts((all) => ({ ...all, [field]: text }));
    }
  };
  // A picture locked against cropping (`noCrop`), resizing (`noResize`), or rotation (`noRot`)
  // keeps its crop, size, or rotation, as in DOCX Editor: the fields are off, and one note says why.
  const cropLocked = picture?.canCrop === false;
  const sizeLocked = picture?.canResize === false;
  const rotationLocked = picture?.canRotate === false;
  const lockNote = (field: NumberField): string | null =>
    field.startsWith('crop')
      ? cropLocked
        ? `${id}-crop-locked`
        : null
      : field === 'rotation'
        ? rotationLocked
          ? `${id}-rotation-locked`
          : null
        : sizeLocked
          ? `${id}-size-locked`
          : null;
  const number = (field: NumberField, label: string, unit: string): ReactNode => (
    <label className="pptx-field" htmlFor={`${id}-${field}`}>
      {label}
      <span className="pptx-field__control">
        <input
          id={`${id}-${field}`}
          type="number"
          step={field === 'width' || field === 'height' ? 0.1 : 1}
          disabled={lockNote(field) !== null}
          aria-invalid={draft.errors[field] ? true : false}
          aria-describedby={
            [lockNote(field), draft.errors[field] ? `${id}-error` : null]
              .filter(Boolean)
              .join(' ') || undefined
          }
          value={texts[field]}
          onChange={(e) => set(field, e.target.value)}
        />
        <span className="pptx-field__unit">{unit}</span>
      </span>
    </label>
  );
  // The first error, with the name of its field.
  const bad = NUMBER_FIELDS.find((f) => draft.errors[f]);
  const error = bad
    ? t('imageProperties.fieldError', {
        field: t(FIELD_LABELS[bad]),
        message: translateKeyOrText(t, draft.errors[bad]!),
      })
    : draft.errors.form
      ? translateKeyOrText(t, draft.errors.form)
      : null;
  return (
    <DialogFrame
      kind="imageProperties"
      label={t('imageProperties.title')}
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
                {t('imageProperties.title')}
              </h2>
              <ToolButton label={t('imageProperties.close')} icon="close" onClick={onClose} />
            </div>
            <div data-pptx-part="body" className="pptx-dialog__body">
              <section
                data-pptx-part="field"
                data-pptx-field="size"
                className="pptx-dialog__section"
                role="group"
                aria-labelledby={`${id}-size`}
              >
                <h3 id={`${id}-size`} className="pptx-dialog__label">
                  {t('imageProperties.size')}
                </h3>
                {sizeLocked ? (
                  <p id={`${id}-size-locked`} className="pptx-dialog__muted">
                    {t('disabled.sizeLocked')}
                  </p>
                ) : null}
                <div className="pptx-panel__grid">
                  {number('height', t('imageProperties.height'), t('imageProperties.inches'))}
                  {number('width', t('imageProperties.width'), t('imageProperties.inches'))}
                </div>
                <label className="pptx-radio">
                  <input
                    type="checkbox"
                    checked={values.lockAspect}
                    disabled={sizeLocked}
                    aria-describedby={sizeLocked ? `${id}-size-locked` : undefined}
                    onChange={(e) => draft.setValue('lockAspect', e.target.checked)}
                  />
                  {t('imageProperties.lockAspect')}
                </label>
              </section>
              <section
                data-pptx-part="field"
                data-pptx-field="rotation"
                className="pptx-dialog__section"
              >
                {rotationLocked ? (
                  <p id={`${id}-rotation-locked`} className="pptx-dialog__muted">
                    {t('disabled.rotationLocked')}
                  </p>
                ) : null}
                <div className="pptx-panel__grid">
                  {number('rotation', t('imageProperties.rotation'), t('imageProperties.degrees'))}
                </div>
              </section>
              <section
                data-pptx-part="field"
                data-pptx-field="crop"
                className="pptx-dialog__section"
                role="group"
                aria-labelledby={`${id}-crop`}
              >
                <h3 id={`${id}-crop`} className="pptx-dialog__label">
                  {t('imageProperties.crop')}
                </h3>
                {cropLocked ? (
                  <p id={`${id}-crop-locked`} className="pptx-dialog__muted">
                    {t('disabled.cropLocked')}
                  </p>
                ) : null}
                <div className="pptx-panel__grid">
                  {number('cropLeft', t('imageProperties.cropLeft'), t('imageProperties.percent'))}
                  {number('cropTop', t('imageProperties.cropTop'), t('imageProperties.percent'))}
                  {number(
                    'cropRight',
                    t('imageProperties.cropRight'),
                    t('imageProperties.percent')
                  )}
                  {number(
                    'cropBottom',
                    t('imageProperties.cropBottom'),
                    t('imageProperties.percent')
                  )}
                </div>
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
 * Read the draft of the enclosing Format Picture dialog, for a host's own field inside it. It
 * does not open the dialog: `useEditorDialog('imageProperties').open()` does. @public
 */
export function useImagePropertiesDialog(): UseImagePropertiesDialogReturn {
  return parts.useState();
}

/**
 * Format Picture: size, rotation, and crop position of the selected picture, with replaceable
 * parts (`Header`, `Title`, `Body`, `Footer`, `Field`, `Error`, `Apply`, `Cancel`).
 *
 * @example
 * ```tsx
 * <PptxEditor.ImagePropertiesDialog {...props}>
 *   <PptxEditor.ImagePropertiesDialog.Field name="rotation" hidden />
 * </PptxEditor.ImagePropertiesDialog>
 * ```
 *
 * @public
 */
export const ImagePropertiesDialog = Object.assign(ImagePropertiesDialogRoot, {
  Header: parts.Header,
  Title: parts.Title,
  Body: parts.Body,
  Footer: parts.Footer,
  Field: parts.Field,
  Error: parts.Error,
  Apply: parts.Apply,
  Cancel: parts.Cancel,
});
