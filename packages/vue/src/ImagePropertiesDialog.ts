// The Format Picture dialog: the Vue twin of the React adapter's `ImagePropertiesDialog`, on
// the same dialog parts and the same core rules (`picture-fields.ts`), so both show the same
// fields and apply the same `setImageProperties` command. The host owns visibility (`open`,
// `onClose`).

import {
  Fragment,
  computed,
  defineComponent,
  h,
  reactive,
  useId,
  watch,
  type CSSProperties,
  type ExtractPublicPropTypes,
  type PropType,
  type VNode,
} from 'vue';
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
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { ToolButton } from './chrome.tsx';
import {
  DialogFrame,
  applyOnEnter,
  createDialogParts,
  useCloseOnNewDocument,
  useDialogDraft,
  type UseDialogReturn,
} from './dialog-parts.ts';

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

const parts = createDialogParts<UseImagePropertiesDialogReturn, ImagePropertiesDialogFieldName>(
  'ImagePropertiesDialog'
);

const imagePropertiesDialogProps = {
  /** Whether the dialog is shown. The host owns this state. */
  open: { type: Boolean, required: true as const },
  /** Called on Cancel, Escape, a click on the backdrop, and after OK. */
  onClose: { type: Function as PropType<() => void>, required: true as const },
  /** `false`: only the default slot, with no default arrangement. */
  preset: { type: Boolean, default: true },
  /** A class name added to the dialog surface, as React's `className` (`class` works too). */
  className: { type: String, default: undefined },
  /** Styles added to the dialog surface. */
  style: { type: Object as PropType<CSSProperties>, default: undefined },
} as const;

/** Props of `ImagePropertiesDialog`, the twin of React's `ImagePropertiesDialogProps`. @public */
export type ImagePropertiesDialogProps = ExtractPublicPropTypes<typeof imagePropertiesDialogProps>;

const ImagePropertiesDialogOpen = defineComponent({
  name: 'ImagePropertiesDialogOpen',
  inheritAttrs: false,
  props: imagePropertiesDialogProps,
  setup(props, { attrs, slots }) {
    const { t } = useTranslation();
    const editorRef = usePptxEditor();
    const id = useId();
    const picture = useEditorState((s) => s.image ?? null);
    // Format Picture edits one picture.
    const enabled = useEditorState(
      (s) => s.status === 'ready' && s.mode === 'edit' && s.image?.count === 1
    );
    const initial = picturePropertiesDraftFrom(picture.value);
    const exactCrop = picture.value?.crop;
    // As the React twin: the draft belongs to the picture the dialog opened on, so another
    // selection, the picture's removal, or a new document closes the dialog.
    const openedOn = picture.value?.nodeId ?? null;
    watch(
      () => picture.value?.nodeId ?? null,
      (now) => {
        if (now !== openedOn) props.onClose();
      }
    );
    useCloseOnNewDocument(() => props.onClose());
    const draft = useDialogDraft<ImagePropertiesDialogFields>({
      initial,
      validate: (values) => validatePicturePropertiesDraft(values, initial),
      commit: (values) =>
        commitDialogCommand(
          editorRef.value,
          picturePropertiesDraftCommand(values, initial, exactCrop)
        ),
    });
    const state: UseImagePropertiesDialogReturn = {
      values: draft.values,
      // The state gives messages; the raw keys stay inside, for focus.
      errors: computed(() => translateDialogErrors(t, draft.errors.value)),
      isEnabled: computed(() => enabled.value),
      // A host part's setValue takes the packaged fields' rules: Lock aspect ratio and the shown text.
      setValue: (name, value) => {
        if (name !== 'lockAspect' && typeof value === 'number') set(name, String(value));
        else draft.setValue(name, value);
      },
      apply: () => {
        if (draft.apply()) return props.onClose();
        // Focus the first field with an error (WCAG 3.3.1).
        const errors = validatePicturePropertiesDraft(draft.values.value, initial);
        const bad = NUMBER_FIELDS.find((f) => errors[f]);
        if (bad) document.getElementById(`${id}-${bad}`)?.focus();
      },
      cancel: () => props.onClose(),
    };
    // The text of each field as typed; the draft holds the numbers.
    const texts = reactive(
      Object.fromEntries(NUMBER_FIELDS.map((f) => [f, String(initial[f])])) as Record<
        NumberField,
        string
      >
    );
    const set = (field: NumberField, text: string): void => {
      const value = text.trim() === '' ? NaN : Number(text);
      texts[field] = text;
      if (field !== 'width' && field !== 'height') {
        draft.setValue(field, value);
        return;
      }
      const values = draft.values.value;
      const next = picturePropertiesWithValue(values, field, value, initial);
      const other = field === 'width' ? 'height' : 'width';
      draft.setValue('width', next.width);
      draft.setValue('height', next.height);
      if (next[other] !== values[other]) texts[other] = String(next[other]);
    };
    const render = parts.provideParts(state);
    return () => {
      const errors = draft.errors.value;
      // The first error, with the name of its field.
      const bad = NUMBER_FIELDS.find((f) => errors[f]);
      const error = bad
        ? t('imageProperties.fieldError', {
            field: t(FIELD_LABELS[bad]),
            message: translateKeyOrText(t, errors[bad]!),
          })
        : errors.form
          ? translateKeyOrText(t, errors.form)
          : '';
      // A picture locked against cropping (`noCrop`), resizing (`noResize`), or rotation
      // (`noRot`) keeps it, as in DOCX Editor: the fields are off, and one note says why.
      const cropLocked = picture.value?.canCrop === false;
      const sizeLocked = picture.value?.canResize === false;
      const rotationLocked = picture.value?.canRotate === false;
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
      const note = (
        locked: boolean,
        field: 'size' | 'crop' | 'rotation',
        key: 'disabled.sizeLocked' | 'disabled.cropLocked' | 'disabled.rotationLocked'
      ) =>
        locked
          ? h('p', { id: `${id}-${field}-locked`, class: 'pptx-dialog__muted' }, t(key))
          : null;
      const number = (field: NumberField, label: string, unit: string): VNode =>
        h('label', { class: 'pptx-field', for: `${id}-${field}` }, [
          label,
          h('span', { class: 'pptx-field__control' }, [
            h('input', {
              id: `${id}-${field}`,
              type: 'number',
              step: field === 'width' || field === 'height' ? 0.1 : 1,
              disabled: lockNote(field) !== null,
              'aria-invalid': errors[field] ? 'true' : 'false',
              'aria-describedby':
                [lockNote(field), errors[field] ? `${id}-error` : null].filter(Boolean).join(' ') ||
                undefined,
              value: texts[field],
              onInput: (e: Event) => set(field, (e.target as HTMLInputElement).value),
            }),
            h('span', { class: 'pptx-field__unit' }, unit),
          ]),
        ]);
      // A section with a heading is a labelled group, so a screen reader names its fields.
      const section = (
        field: ImagePropertiesDialogFieldName,
        children: (VNode | null)[],
        labelled = false
      ) =>
        h(
          'section',
          {
            'data-pptx-part': 'field',
            'data-pptx-field': field,
            class: 'pptx-dialog__section',
            ...(labelled ? { role: 'group', 'aria-labelledby': `${id}-${field}` } : {}),
          },
          children
        );
      const inches = t('imageProperties.inches');
      const percent = t('imageProperties.percent');
      const defaults = h(Fragment, [
        h('div', { 'data-pptx-part': 'header', class: 'pptx-dialog__header' }, [
          h(
            'h2',
            { 'data-pptx-part': 'title', class: 'pptx-dialog__title' },
            t('imageProperties.title')
          ),
          h(ToolButton, {
            label: t('imageProperties.close'),
            icon: 'close',
            onClick: props.onClose,
          }),
        ]),
        h('div', { 'data-pptx-part': 'body', class: 'pptx-dialog__body' }, [
          section(
            'size',
            [
              h('h3', { id: `${id}-size`, class: 'pptx-dialog__label' }, t('imageProperties.size')),
              note(sizeLocked, 'size', 'disabled.sizeLocked'),
              h('div', { class: 'pptx-panel__grid' }, [
                number('height', t('imageProperties.height'), inches),
                number('width', t('imageProperties.width'), inches),
              ]),
              h('label', { class: 'pptx-radio' }, [
                h('input', {
                  type: 'checkbox',
                  checked: draft.values.value.lockAspect,
                  disabled: sizeLocked,
                  'aria-describedby': sizeLocked ? `${id}-size-locked` : undefined,
                  onChange: (e: Event) =>
                    draft.setValue('lockAspect', (e.target as HTMLInputElement).checked),
                }),
                t('imageProperties.lockAspect'),
              ]),
            ],
            true
          ),
          section('rotation', [
            note(rotationLocked, 'rotation', 'disabled.rotationLocked'),
            h('div', { class: 'pptx-panel__grid' }, [
              number('rotation', t('imageProperties.rotation'), t('imageProperties.degrees')),
            ]),
          ]),
          section(
            'crop',
            [
              h('h3', { id: `${id}-crop`, class: 'pptx-dialog__label' }, t('imageProperties.crop')),
              note(cropLocked, 'crop', 'disabled.cropLocked'),
              h('div', { class: 'pptx-panel__grid' }, [
                number('cropLeft', t('imageProperties.cropLeft'), percent),
                number('cropTop', t('imageProperties.cropTop'), percent),
                number('cropRight', t('imageProperties.cropRight'), percent),
                number('cropBottom', t('imageProperties.cropBottom'), percent),
              ]),
            ],
            true
          ),
        ]),
        h('div', { 'data-pptx-part': 'footer', class: 'pptx-dialog__footer' }, [
          h(
            'span',
            {
              id: `${id}-error`,
              'data-pptx-part': 'error',
              role: 'alert',
              class: 'pptx-dialog__error',
            },
            error
          ),
          h(
            'button',
            {
              type: 'button',
              'data-pptx-part': 'cancel',
              class: 'pptx-button',
              onClick: state.cancel,
            },
            t('dialog.cancel')
          ),
          h(
            'button',
            {
              type: 'button',
              'data-pptx-part': 'apply',
              class: 'pptx-button pptx-button--primary',
              disabled: !enabled.value,
              onClick: state.apply,
            },
            t('dialog.ok')
          ),
        ]),
      ]);
      return h(
        DialogFrame,
        {
          ...attrs,
          class: [attrs.class, props.className],
          style: [attrs.style, props.style],
          kind: 'imageProperties',
          label: t('imageProperties.title'),
          onClose: props.onClose,
          onKeydown: (event: KeyboardEvent) => applyOnEnter(event, enabled.value, state.apply),
        },
        () => render(defaults, slots.default?.(), props.preset)
      );
    };
  },
});

/**
 * Format Picture: size, rotation, and crop position of the selected picture, with replaceable
 * parts (`Header`, `Title`, `Body`, `Footer`, `Field`, `Error`, `Apply`, `Cancel`), as React's
 * `ImagePropertiesDialog`. Renders nothing while `open` is false. @public
 */
export const ImagePropertiesDialog = Object.assign(
  defineComponent({
    name: 'ImagePropertiesDialog',
    inheritAttrs: false,
    props: imagePropertiesDialogProps,
    setup(props, { attrs, slots }) {
      return () =>
        props.open ? h(ImagePropertiesDialogOpen, { ...attrs, ...props }, slots) : null;
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
 * Read the draft of the enclosing Format Picture dialog, for a host's own field inside it. It
 * does not open the dialog: `useEditorDialog('imageProperties').open()` does. @public
 */
export function useImagePropertiesDialog(): UseImagePropertiesDialogReturn {
  return parts.useState();
}
