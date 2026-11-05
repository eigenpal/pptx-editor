// The Alt Text dialog: the Vue twin of the React adapter's `ImageAltTextDialog`, on the same
// dialog parts and the same core rules (`picture-fields.ts`), so both show the same fields and
// apply the same `setAltText` command. The host owns visibility (`open`, `onClose`).

import {
  Fragment,
  computed,
  defineComponent,
  h,
  useId,
  watch,
  type CSSProperties,
  type ExtractPublicPropTypes,
  type PropType,
} from 'vue';
import {
  MAX_ALT_TEXT,
  altTextDraftCommand,
  altTextDraftFrom,
  commitDialogCommand,
  translateDialogErrors,
  translateKeyOrText,
  validateAltTextDraft,
} from '@pptx-editor.dev/core/editor';
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

const parts = createDialogParts<UseImageAltTextDialogReturn, ImageAltTextDialogFieldName>(
  'ImageAltTextDialog'
);

const imageAltTextDialogProps = {
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

/** Props of `ImageAltTextDialog`, the twin of React's `ImageAltTextDialogProps`. @public */
export type ImageAltTextDialogProps = ExtractPublicPropTypes<typeof imageAltTextDialogProps>;

const ImageAltTextDialogOpen = defineComponent({
  name: 'ImageAltTextDialogOpen',
  inheritAttrs: false,
  props: imageAltTextDialogProps,
  setup(props, { attrs, slots }) {
    const { t } = useTranslation();
    const editorRef = usePptxEditor();
    const id = useId();
    const current = useEditorState((s) => s.geometry?.altText);
    const enabled = useEditorState(
      (s) => s.status === 'ready' && s.mode === 'edit' && s.geometry !== null
    );
    const initial = altTextDraftFrom(current.value);
    // As the React twin: the draft belongs to the object the dialog opened on, so another
    // selection, the object's removal, or a new document closes the dialog.
    const nodeId = useEditorState((s) => s.geometry?.nodeId ?? null);
    const openedOn = nodeId.value;
    watch(nodeId, (now) => {
      if (now !== openedOn) props.onClose();
    });
    useCloseOnNewDocument(() => props.onClose());
    const draft = useDialogDraft<ImageAltTextDialogFields>({
      initial,
      validate: validateAltTextDraft,
      commit: (values) =>
        commitDialogCommand(editorRef.value, altTextDraftCommand(values, initial)),
    });
    const state: UseImageAltTextDialogReturn = {
      values: draft.values,
      // The state gives messages; the raw keys stay inside, for focus.
      errors: computed(() => translateDialogErrors(t, draft.errors.value)),
      isEnabled: computed(() => enabled.value),
      setValue: draft.setValue,
      apply: () => {
        if (draft.apply()) return props.onClose();
        // Focus the field with the error (WCAG 3.3.1), as Format Picture does.
        const errors = validateAltTextDraft(draft.values.value);
        const bad = (['description', 'title'] as const).find((f) => errors[f]);
        if (bad) document.getElementById(`${id}-${bad}`)?.focus();
      },
      cancel: () => props.onClose(),
    };
    const render = parts.provideParts(state);
    return () => {
      const { description, title, decorative } = draft.values.value;
      const errors = draft.errors.value;
      // The first error, with the name of its field.
      const bad = errors.description ? 'description' : errors.title ? 'title' : null;
      const error = bad
        ? t('imageAltText.fieldError', {
            field: t(
              bad === 'description' ? 'imageAltText.description' : 'imageAltText.titleLabel'
            ),
            message: translateKeyOrText(t, errors[bad]!),
          })
        : errors.form
          ? translateKeyOrText(t, errors.form)
          : '';
      const field = (name: ImageAltTextDialogFieldName, extra: string, children: unknown[]) =>
        h(
          'section',
          {
            'data-pptx-part': 'field',
            'data-pptx-field': name,
            class: `pptx-dialog__section${extra}`,
          },
          children as never
        );
      const defaults = h(Fragment, [
        h('div', { 'data-pptx-part': 'header', class: 'pptx-dialog__header' }, [
          h(
            'h2',
            { 'data-pptx-part': 'title', class: 'pptx-dialog__title' },
            t('imageAltText.title')
          ),
          h(ToolButton, { label: t('imageAltText.close'), icon: 'close', onClick: props.onClose }),
        ]),
        h('div', { 'data-pptx-part': 'body', class: 'pptx-dialog__body' }, [
          field('description', '', [
            h(
              'label',
              { class: 'pptx-dialog__label', for: `${id}-description` },
              t('imageAltText.description')
            ),
            h(
              'p',
              { class: 'pptx-dialog__muted', id: `${id}-help` },
              t('imageAltText.descriptionHelp')
            ),
            h('textarea', {
              id: `${id}-description`,
              class: 'pptx-dialog__textarea',
              rows: 5,
              value: description,
              disabled: decorative,
              'aria-invalid': errors.description ? 'true' : 'false',
              'aria-describedby': [
                `${id}-help`,
                `${id}-count`,
                decorative ? `${id}-decorative-reason` : null,
                bad === 'description' ? `${id}-error` : null,
              ]
                .filter(Boolean)
                .join(' '),
              onInput: (e: Event) =>
                draft.setValue('description', (e.target as HTMLTextAreaElement).value),
            }),
            decorative
              ? h(
                  'span',
                  { class: 'pptx-sr-only', id: `${id}-decorative-reason` },
                  t('imageAltText.decorativeReason')
                )
              : null,
            h(
              'span',
              { class: 'pptx-dialog__muted', id: `${id}-count` },
              t('imageAltText.count', { count: description.length, max: MAX_ALT_TEXT })
            ),
          ]),
          field('title', ' pptx-field', [
            h(
              'label',
              { class: 'pptx-dialog__label', for: `${id}-title` },
              t('imageAltText.titleLabel')
            ),
            h('input', {
              id: `${id}-title`,
              type: 'text',
              value: title,
              'aria-invalid': errors.title ? 'true' : 'false',
              'aria-describedby': bad === 'title' ? `${id}-error` : undefined,
              onInput: (e: Event) => draft.setValue('title', (e.target as HTMLInputElement).value),
            }),
          ]),
          field('decorative', '', [
            h('label', { class: 'pptx-radio' }, [
              h('input', {
                type: 'checkbox',
                checked: decorative,
                'aria-describedby': `${id}-decorative-help`,
                onChange: (e: Event) =>
                  draft.setValue('decorative', (e.target as HTMLInputElement).checked),
              }),
              t('imageAltText.decorative'),
            ]),
            h(
              'p',
              { class: 'pptx-dialog__muted', id: `${id}-decorative-help` },
              t('imageAltText.decorativeHelp')
            ),
          ]),
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
          kind: 'imageAltText',
          label: t('imageAltText.title'),
          onClose: props.onClose,
          // Enter in the title applies, as in the other dialogs; the description takes newlines.
          onKeydown: (event: KeyboardEvent) => applyOnEnter(event, enabled.value, state.apply),
        },
        () => render(defaults, slots.default?.(), props.preset)
      );
    };
  },
});

/**
 * Alt text of the selected element: description, title, and "Mark as decorative", with
 * replaceable parts (`Header`, `Title`, `Body`, `Footer`, `Field`, `Error`, `Apply`, `Cancel`),
 * as React's `ImageAltTextDialog`. Renders nothing while `open` is false. @public
 */
export const ImageAltTextDialog = Object.assign(
  defineComponent({
    name: 'ImageAltTextDialog',
    inheritAttrs: false,
    props: imageAltTextDialogProps,
    setup(props, { attrs, slots }) {
      return () => (props.open ? h(ImageAltTextDialogOpen, { ...attrs, ...props }, slots) : null);
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
 * Read the draft of the enclosing Alt Text dialog, for a host's own field inside it. It does
 * not open the dialog: `useEditorDialog('imageAltText').open()` does. @public
 */
export function useImageAltTextDialog(): UseImageAltTextDialogReturn {
  return parts.useState();
}
