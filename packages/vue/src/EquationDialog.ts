// The Equation dialog: the Vue twin of the React adapter's `EquationDialog`, on the same dialog
// parts and the same core rules (`equation-fields.ts`), so both show the same gallery and apply
// the same `insertEquation` or `setEquation` command. The host owns visibility (`open`,
// `onClose`).

import {
  Fragment,
  computed,
  defineComponent,
  h,
  useId,
  type CSSProperties,
  type ExtractPublicPropTypes,
  type PropType,
} from 'vue';
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

/** The draft of the Equation dialog. @public */
export interface EquationDialogFields {
  /** The linear form of the equation, such as `x=(-b\pm\sqrt(b^2-4ac))/(2a)`. */
  readonly linear: string;
}

/** The Equation draft and its actions. @public */
export type UseEquationDialogReturn = UseDialogReturn<EquationDialogFields>;

/** The field names of the Equation dialog, for `EquationDialog.Field`. @public */
export type EquationDialogFieldName = 'gallery' | 'linear';

const parts = createDialogParts<UseEquationDialogReturn, EquationDialogFieldName>('EquationDialog');

const equationDialogProps = {
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

/** Props of `EquationDialog`, the twin of React's `EquationDialogProps`. @public */
export type EquationDialogProps = ExtractPublicPropTypes<typeof equationDialogProps>;

const EquationDialogOpen = defineComponent({
  name: 'EquationDialogOpen',
  inheritAttrs: false,
  props: equationDialogProps,
  setup(props, { attrs, slots }) {
    const { t } = useTranslation();
    const editorRef = usePptxEditor();
    const id = useId();
    // As the React twin: the equation the dialog opened on, which OK edits.
    const selected = useEditorState((s) => s.equation).value;
    const enabled = useEditorState((s) => s.status === 'ready' && s.mode === 'edit');
    const edits = equationDialogEdits(selected);
    useCloseOnNewDocument(() => props.onClose());
    const draft = useDialogDraft<EquationDialogFields>({
      initial: equationDraftFrom(selected),
      validate: validateEquationDraft,
      commit: (values) =>
        commitDialogCommand(editorRef.value, equationDraftCommand(values, selected)),
    });
    const state: UseEquationDialogReturn = {
      values: draft.values,
      errors: computed(() => translateDialogErrors(t, draft.errors.value)),
      isEnabled: computed(() => enabled.value),
      setValue: draft.setValue,
      apply: () => {
        if (draft.apply()) return props.onClose();
        // Focus the field with the error (WCAG 3.3.1).
        if (validateEquationDraft(draft.values.value).linear)
          document.getElementById(`${id}-linear`)?.focus();
      },
      cancel: () => props.onClose(),
    };
    const render = parts.provideParts(state);
    return () => {
      const { linear } = draft.values.value;
      const errors = draft.errors.value;
      const error = errors.linear
        ? translateKeyOrText(t, errors.linear)
        : errors.form
          ? translateKeyOrText(t, errors.form)
          : '';
      const title = t(edits ? 'equation.editTitle' : 'equation.insertTitle');
      const defaults = h(Fragment, [
        h('div', { 'data-pptx-part': 'header', class: 'pptx-dialog__header' }, [
          h('h2', { 'data-pptx-part': 'title', class: 'pptx-dialog__title' }, title),
          h(ToolButton, { label: t('equation.close'), icon: 'close', onClick: props.onClose }),
        ]),
        h('div', { 'data-pptx-part': 'body', class: 'pptx-dialog__body' }, [
          selected && !edits
            ? h('p', { class: 'pptx-dialog__muted', role: 'note' }, t('equation.notEditable'))
            : null,
          h(
            'section',
            {
              'data-pptx-part': 'field',
              'data-pptx-field': 'gallery',
              class: 'pptx-dialog__section',
              'aria-labelledby': `${id}-gallery`,
            },
            [
              h('h3', { class: 'pptx-dialog__label', id: `${id}-gallery` }, t('equation.gallery')),
              h(
                'div',
                { class: 'pptx-equation-gallery' },
                EQUATION_GALLERY.map((item) =>
                  h(
                    'button',
                    {
                      key: item.id,
                      type: 'button',
                      class: 'pptx-button pptx-equation-gallery__item',
                      'aria-pressed': linear === item.linear ? 'true' : 'false',
                      onClick: () => draft.setValue('linear', item.linear),
                    },
                    [
                      h('span', t(item.labelKey)),
                      h('code', { class: 'pptx-dialog__muted' }, item.linear),
                    ]
                  )
                )
              ),
            ]
          ),
          h(
            'section',
            {
              'data-pptx-part': 'field',
              'data-pptx-field': 'linear',
              class: 'pptx-dialog__section pptx-field',
            },
            [
              h(
                'label',
                { class: 'pptx-dialog__label', for: `${id}-linear` },
                t('equation.linear')
              ),
              h('p', { class: 'pptx-dialog__muted', id: `${id}-help` }, t('equation.linearHelp')),
              h('input', {
                id: `${id}-linear`,
                type: 'text',
                spellcheck: false,
                autocomplete: 'off',
                maxlength: MAX_LINEAR_LENGTH,
                value: linear,
                'aria-invalid': errors.linear ? 'true' : 'false',
                'aria-describedby': errors.linear ? `${id}-help ${id}-error` : `${id}-help`,
                onInput: (e: Event) =>
                  draft.setValue('linear', (e.target as HTMLInputElement).value),
              }),
            ]
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
            edits ? t('dialog.ok') : t('equation.insert')
          ),
        ]),
      ]);
      return h(
        DialogFrame,
        {
          ...attrs,
          class: [attrs.class, props.className],
          style: [attrs.style, props.style],
          kind: 'equation',
          label: title,
          onClose: props.onClose,
          onKeydown: (event: KeyboardEvent) => applyOnEnter(event, enabled.value, state.apply),
        },
        () => render(defaults, slots.default?.(), props.preset)
      );
    };
  },
});

/**
 * Insert Equation and Edit Equation: a gallery of common equations and the linear form, with
 * replaceable parts (`Header`, `Title`, `Body`, `Footer`, `Field`, `Error`, `Apply`, `Cancel`),
 * as React's `EquationDialog`. Renders nothing while `open` is false. @public
 */
export const EquationDialog = Object.assign(
  defineComponent({
    name: 'EquationDialog',
    inheritAttrs: false,
    props: equationDialogProps,
    setup(props, { attrs, slots }) {
      return () => (props.open ? h(EquationDialogOpen, { ...attrs, ...props }, slots) : null);
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
 * Read the draft of the enclosing Equation dialog, for a host's own field inside it. It does
 * not open the dialog: `useEditorDialog('equation').open()` does. @public
 */
export function useEquationDialog(): UseEquationDialogReturn {
  return parts.useState();
}
