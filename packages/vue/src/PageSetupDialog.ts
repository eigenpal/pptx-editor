// The Page setup dialog: the Vue twin of the React adapter's `PageSetupDialog`, on the same
// dialog parts and the same core helpers, so both show the same fields and apply the same
// draft. The host owns visibility (`open`, `onClose`).

import {
  Fragment,
  computed,
  defineComponent,
  h,
  type CSSProperties,
  type ExtractPublicPropTypes,
  type PropType,
  type VNodeArrayChildren,
} from 'vue';
import {
  SLIDE_SIZE_CHOICES,
  commitDialogCommand,
  radioGroupStep,
  seedSlideSizeFields,
  slideSizeCommandOf,
  slideSizeFieldError,
  translateDialogErrors,
  withSlideDimension,
  withSlideOrientation,
  withSlideSizePreset,
  type SlideOrientation,
  type SlideSizeFields,
} from '@pptx-editor.dev/core/editor';
import { ToolButton } from './chrome.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { CheckField, NumberField, SelectField } from './dialog-form.tsx';
import {
  DialogFrame,
  applyOnEnter,
  createDialogParts,
  useCloseOnNewDocument,
  useDialogDraft,
  type UseDialogReturn,
} from './dialog-parts.ts';

/** The draft of the Page setup dialog. @public */
export interface PageSetupDialogFields {
  /** The chrome theme the appearance choice shows. */
  readonly theme: 'light' | 'dark';
  /** The slide size: a preset, the width and height in inches, and whether content scales. */
  readonly slideSize: SlideSizeFields;
}

/** The Page setup draft and its actions. @public */
export type UsePageSetupDialogReturn = UseDialogReturn<PageSetupDialogFields>;

/** The field names of the Page setup dialog, for `PageSetupDialogField`. @public */
export type PageSetupDialogFieldName = 'slideSize' | 'theme' | 'fonts' | 'substitutes';

const parts = createDialogParts<UsePageSetupDialogReturn, PageSetupDialogFieldName>(
  'PageSetupDialog'
);
const THEMES = ['light', 'dark'] as const;
const ORIENTATIONS: readonly SlideOrientation[] = ['landscape', 'portrait'];

/** A segmented radio group with the arrow keys of the WAI-ARIA radio group pattern. */
function segmented<T extends string>(o: {
  readonly label: string;
  readonly choices: readonly T[];
  readonly value: T;
  readonly disabled?: boolean;
  readonly labelOf: (choice: T) => string;
  readonly onChange: (choice: T) => void;
}) {
  return h(
    'div',
    {
      class: 'pptx-segmented',
      role: 'radiogroup',
      'aria-label': o.label,
      'aria-disabled': o.disabled || undefined,
      onKeydown: (event: KeyboardEvent) => {
        if (o.disabled) return;
        const next = radioGroupStep(o.choices, o.value, event.key);
        if (!next) return;
        event.preventDefault();
        o.onChange(next);
        (event.currentTarget as HTMLElement)
          .querySelector<HTMLElement>(`[data-value="${next}"]`)
          ?.focus();
      },
    },
    o.choices.map((m) =>
      h(
        'button',
        {
          key: m,
          type: 'button',
          role: 'radio',
          'data-value': m,
          disabled: o.disabled,
          'aria-checked': o.value === m,
          tabindex: o.value === m && !o.disabled ? 0 : -1,
          onClick: () => o.onChange(m),
        },
        o.labelOf(m)
      )
    )
  );
}

const pageSetupDialogProps = {
  /** Whether the dialog is shown. The host owns this state. */
  open: { type: Boolean, required: true as const },
  /** Called on Cancel, Escape, a click on the backdrop, and after OK. */
  onClose: { type: Function as PropType<() => void>, required: true as const },
  /** The chrome theme; the appearance choice shows only with `onThemeChange`. */
  theme: { type: String as PropType<'light' | 'dark'>, default: 'light' },
  /** Called by OK when the appearance choice changed. */
  onThemeChange: {
    type: Function as PropType<(theme: 'light' | 'dark') => void>,
    default: undefined,
  },
  /** `false`: only the default slot, with no default arrangement. */
  preset: { type: Boolean, default: true },
  /** A class name added to the dialog surface, as React's `className` (`class` works too). */
  className: { type: String, default: undefined },
  /** Styles added to the dialog surface. */
  style: { type: Object as PropType<CSSProperties>, default: undefined },
} as const;

/** Props of `PageSetupDialog`, the twin of React's `PageSetupDialogProps`. @public */
export type PageSetupDialogProps = ExtractPublicPropTypes<typeof pageSetupDialogProps>;

const PageSetupDialogOpen = defineComponent({
  name: 'PageSetupDialogOpen',
  inheritAttrs: false,
  props: pageSetupDialogProps,
  setup(props, { attrs, slots }) {
    const { t } = useTranslation();
    const editor = usePptxEditor();
    const width = useEditorState((s) => s.slideWidth);
    const height = useEditorState((s) => s.slideHeight);
    const editable = useEditorState((s) => s.status === 'ready' && s.mode === 'edit');
    const embedded = useEditorState((s) => s.embeddedFontFaces);
    const substituted = useEditorState((s) => s.substitutedFonts);
    const opening = props.theme;
    // The slide size the dialog opened on: OK sends a size only when the fields moved from it.
    const opened = { width: width.value, height: height.value };
    const seed = seedSlideSizeFields(opened.width, opened.height);
    const draft = useDialogDraft<PageSetupDialogFields>({
      initial: { theme: opening, slideSize: seed },
      validate: (values) => {
        const problem = slideSizeFieldError(seed, values.slideSize, opened);
        return problem ? { slideSize: problem.key } : {};
      },
      // As the React twin: the slide size first, then a change from the opening choice,
      // reported to the latest callback.
      commit: (values) => {
        const resized = commitDialogCommand(
          editor.value,
          slideSizeCommandOf(seed, values.slideSize, opened)
        );
        if (!resized.ok) return resized;
        if (values.theme !== opening) props.onThemeChange?.(values.theme);
        return { ok: true };
      },
    });
    const setSize = (next: SlideSizeFields): void => draft.setValue('slideSize', next);
    useCloseOnNewDocument(() => props.onClose());
    const state: UsePageSetupDialogReturn = {
      values: draft.values,
      errors: computed(() => translateDialogErrors(t, draft.errors.value)),
      isEnabled: computed(() => true),
      setValue: draft.setValue,
      apply: () => {
        if (draft.apply()) props.onClose();
      },
      cancel: () => props.onClose(),
    };
    const render = parts.provideParts(state);
    const section = (field: PageSetupDialogFieldName, label: string, body: VNodeArrayChildren) =>
      h(
        'section',
        { 'data-pptx-part': 'field', 'data-pptx-field': field, class: 'pptx-dialog__section' },
        [h('h3', { class: 'pptx-dialog__label' }, label), ...body]
      );
    return () => {
      const size = draft.values.value.slideSize;
      const sizeError = state.errors.value.slideSize;
      const defaults = h(Fragment, [
        h('div', { 'data-pptx-part': 'header', class: 'pptx-dialog__header' }, [
          h('h2', { 'data-pptx-part': 'title', class: 'pptx-dialog__title' }, t('pageSetup.title')),
          h(ToolButton, { label: t('pageSetup.close'), icon: 'close', onClick: props.onClose }),
        ]),
        h('div', { 'data-pptx-part': 'body', class: 'pptx-dialog__body' }, [
          section('slideSize', t('pageSetup.slideSize'), [
            // Viewing mode: the size as text, as the fields cannot change it.
            editable.value
              ? null
              : h('p', { class: 'pptx-dialog__value' }, [
                  t(`pageSetup.slideSizeName.${seed.size}`),
                  h(
                    'span',
                    { class: 'pptx-dialog__muted' },
                    t('pageSetup.slideSizeValue', { width: seed.width, height: seed.height })
                  ),
                ]),
            h('div', { class: 'pptx-dialog__rows', hidden: !editable.value }, [
              h(SelectField, {
                name: 'size',
                label: t('pageSetup.size'),
                disabled: !editable.value,
                value: size.size,
                options: SLIDE_SIZE_CHOICES.map(
                  (c) => [c, t(`pageSetup.slideSizeName.${c}`)] as const
                ),
                onChange: (c: SlideSizeFields['size']) => setSize(withSlideSizePreset(size, c)),
              }),
              h(
                'div',
                { class: 'pptx-dialog__grid' },
                (['width', 'height'] as const).map((field) =>
                  h(NumberField, {
                    key: field,
                    name: field,
                    label: t(`pageSetup.${field}`),
                    unit: t('pageSetup.inches'),
                    step: '0.01',
                    min: '1',
                    max: '56',
                    disabled: !editable.value,
                    value: size[field],
                    onChange: (v: string) => setSize(withSlideDimension(size, field, v)),
                  })
                )
              ),
              segmented({
                label: t('pageSetup.orientation'),
                choices: ORIENTATIONS,
                value: size.orientation,
                disabled: !editable.value,
                labelOf: (o) => t(`pageSetup.${o}`),
                onChange: (o) => setSize(withSlideOrientation(size, o)),
              }),
              h(CheckField, {
                name: 'scaleContent',
                label: t('pageSetup.scaleContent'),
                value: size.scaleContent,
                disabled: !editable.value,
                onChange: (on: boolean) => setSize({ ...size, scaleContent: on }),
              }),
              sizeError ? h('p', { role: 'alert', class: 'pptx-dialog__error' }, sizeError) : null,
            ]),
          ]),
          props.onThemeChange
            ? section('theme', t('pageSetup.appearance'), [
                segmented({
                  label: t('pageSetup.appearance'),
                  choices: THEMES,
                  value: draft.values.value.theme,
                  labelOf: (m) => t(`pageSetup.${m}`),
                  onChange: (m) => draft.setValue('theme', m),
                }),
              ])
            : null,
          section('fonts', t('pageSetup.fonts'), [
            h(
              'p',
              { class: 'pptx-dialog__value' },
              embedded.value
                ? t('pageSetup.fontsEmbedded', { count: embedded.value })
                : t('pageSetup.fontsNone')
            ),
            // No substitutes: one line here, and no empty Substituted fonts section.
            substituted.value.length === 0
              ? h('p', { class: 'pptx-dialog__value' }, t('pageSetup.substitutesNone'))
              : null,
          ]),
          substituted.value.length === 0
            ? null
            : section('substitutes', t('pageSetup.substitutes'), [
                h(
                  'ul',
                  { class: 'pptx-dialog__list' },
                  substituted.value.map((f) =>
                    h('li', { key: f.family }, [
                      t('pageSetup.substitution', { family: f.family, substitute: f.substitute }),
                      f.metricCompatible
                        ? h(
                            'span',
                            { class: 'pptx-chip pptx-chip--quiet' },
                            t('pageSetup.metricCompatible')
                          )
                        : null,
                    ])
                  )
                ),
              ]),
        ]),
        h('div', { 'data-pptx-part': 'footer', class: 'pptx-dialog__footer' }, [
          h(
            'span',
            { 'data-pptx-part': 'error', role: 'alert', class: 'pptx-dialog__error' },
            state.errors.value.form ?? ''
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
          kind: 'pageSetup',
          label: t('pageSetup.title'),
          onClose: props.onClose,
          onKeydown: (event: KeyboardEvent) => applyOnEnter(event, true, state.apply),
        },
        () => render(defaults, slots.default?.(), props.preset)
      );
    };
  },
});

/**
 * Page setup: slide size, appearance, and fonts, with replaceable parts (`Header`, `Title`,
 * `Body`, `Footer`, `Field`, `Error`, `Apply`, `Cancel`), as React's `PageSetupDialog`.
 * Renders nothing while `open` is false. @public
 */
export const PageSetupDialog = Object.assign(
  defineComponent({
    name: 'PageSetupDialog',
    inheritAttrs: false,
    props: pageSetupDialogProps,
    setup(props, { attrs, slots }) {
      return () => (props.open ? h(PageSetupDialogOpen, { ...attrs, ...props }, slots) : null);
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
 * Read the draft of the enclosing Page setup dialog, for a host's own field inside it. It does
 * not open the dialog: `useEditorDialog('pageSetup').open()` does. @public
 */
export function usePageSetupDialog(): UsePageSetupDialogReturn {
  return parts.useState();
}
