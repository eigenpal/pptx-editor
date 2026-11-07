/** @jsxImportSource vue */
// Shared parts of the form dialogs (Paragraph, Font, Bullets, and Columns): the Vue twin of the
// React adapter's `dialog-form.tsx`. The modal frame with OK and Cancel sits on the dialog named
// parts (`dialog-parts.ts`), the draft is core's (`useDialogDraft` over `createDialogDraft`),
// and the number, text, select, and check fields carry a unit, a "Mixed" blank, and an error
// state, with the same markup and classes as React. Internal: each dialog exports its own parts.
//
// A form dialog renders its open component only while open (as `PageSetupDialog` does), so the
// draft seeds once when the dialog opens and never while it is open.

import {
  computed,
  defineComponent,
  onMounted,
  ref,
  toValue,
  useId,
  watch,
  watchEffect,
  type CSSProperties,
  type MaybeRefOrGetter,
  type PropType,
  type VNode,
} from 'vue';
import {
  cssFontFamily,
  fieldPart as coreFieldPart,
  formPanelId,
  formTabId,
  formTabStep,
  markDialogError,
  translateDialogErrors,
  type DialogError,
} from '@pptx-editor.dev/core/editor';
import { ToolButton } from './chrome.tsx';
import { useEditorState, useTranslation } from './composables.ts';
import {
  DialogFrame,
  applyOnEnter,
  useCloseOnNewDocument,
  useDialogDraft,
  type DialogParts,
  type UseDialogReturn,
} from './dialog-parts.ts';

export type { DialogError };

/** The named parts a form dialog exposes: `createDialogParts` over its draft. @internal */
export type FormDialogParts<Name extends string, D extends object> = DialogParts<
  UseDialogReturn<D>,
  Name
>;

/**
 * The draft and actions a form dialog hands to its parts, in the shape of DOCX Editor's
 * `UseDialogReturn`: core's draft as refs, its errors translated, and Apply that closes the
 * dialog when the document took the draft. Call it in the open dialog's setup. @internal
 */
export function formDialogState<D extends object>(args: {
  readonly draft: ReturnType<typeof useDialogDraft<D>>;
  readonly isEnabled: MaybeRefOrGetter<boolean>;
  readonly onClose: () => void;
}): UseDialogReturn<D> {
  const { t } = useTranslation();
  const { draft } = args;
  // A dialog closes when its editor opens another document, as the React twin's do.
  useCloseOnNewDocument(() => args.onClose());
  return {
    values: draft.values,
    errors: computed(() => translateDialogErrors(t, draft.errors.value)),
    isEnabled: computed(() => toValue(args.isEnabled)),
    setValue: draft.setValue,
    apply: () => {
      if (draft.apply()) args.onClose();
    },
    cancel: () => args.onClose(),
  };
}

/** The first error a form dialog shows: a field's, else the form's. @internal */
function firstError(errors: Readonly<Record<string, string | undefined>>): DialogError | null {
  for (const [field, message] of Object.entries(errors))
    if (typeof message === 'string' && field !== 'form') return { message, field };
  return typeof errors.form === 'string' ? { message: errors.form } : null;
}

/**
 * The modal frame of a form dialog on the dialog named parts (`dialog-parts.ts`): `Header`,
 * `Title`, `Body`, `Footer`, `Field` (each section with a `name`), `Error`, `Apply`, and
 * `Cancel`, each replaceable by a host child, as in DOCX Editor. The default slot holds the
 * dialog's fields; the `host` slot holds the host's named parts and extra content. @internal
 */
export const FormDialog = defineComponent({
  name: 'PptxFormDialog',
  props: {
    /** `data-pptx-dialog`: the registry key. */
    name: { type: String, required: true as const },
    /** A host class, after `pptx-dialog`. */
    className: { type: String, default: undefined },
    style: { type: Object as PropType<CSSProperties>, default: undefined },
    title: { type: String, required: true as const },
    okLabel: { type: String, required: true as const },
    cancelLabel: { type: String, required: true as const },
    /** The accessible name of the header's close button. */
    closeLabel: { type: String, required: true as const },
    /** The dialog's parts, and the draft they read. */
    parts: {
      // Any dialog's parts: each names its own fields and draft.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      type: Object as PropType<Pick<FormDialogParts<string, any>, 'provideParts'>>,
      required: true as const,
    },
    state: { type: Object as PropType<UseDialogReturn<object>>, required: true as const },
    /** `false`: only the host's content (`preset` of `DialogCustomizationProps`). */
    preset: { type: Boolean, default: true },
  },
  setup(props, { attrs, slots }) {
    const { t } = useTranslation();
    const viewing = useEditorState((s) => s.mode === 'view');
    const errorId = useId();
    const anchor = ref<HTMLElement | null>(null);
    const render = props.parts.provideParts(props.state);
    const error = computed(() => firstError(props.state.errors.value));
    // Mark the field the error names, tie it to the message, and move focus there. The dialog
    // carries `data-invalid` while an error shows, for host styles.
    watch(error, (now) => markDialogError(anchor.value?.closest('dialog'), now, errorId), {
      flush: 'post',
    });
    return () => {
      const { state } = props;
      const enabled = state.isEnabled.value;
      const defaults: VNode = (
        <>
          <div data-pptx-part="header" class="pptx-dialog__header">
            <h2 data-pptx-part="title" class="pptx-dialog__title">
              {props.title}
            </h2>
            <ToolButton label={props.closeLabel} icon="close" onClick={state.cancel} />
          </div>
          <div data-pptx-part="body" class="pptx-dialog__body">
            {slots.default?.()}
          </div>
          <div data-pptx-part="footer" class="pptx-dialog__footer">
            <span id={errorId} data-pptx-part="error" role="alert" class="pptx-dialog__error">
              {error.value?.message ?? null}
            </span>
            <button
              type="button"
              data-pptx-part="cancel"
              class="pptx-button"
              onClick={state.cancel}
            >
              {props.cancelLabel}
            </button>
            <button
              type="button"
              data-pptx-part="apply"
              class="pptx-button pptx-button--primary"
              disabled={!enabled}
              // A disabled OK says why, as every disabled control does.
              title={
                enabled ? undefined : t(viewing.value ? 'disabled.viewing' : 'disabled.noEditor')
              }
              onClick={state.apply}
            >
              {props.okLabel}
            </button>
          </div>
        </>
      ) as VNode;
      return (
        <DialogFrame
          {...attrs}
          class={[attrs.class, props.className]}
          style={[attrs.style as CSSProperties, props.style]}
          kind={props.name}
          label={props.title}
          onClose={state.cancel}
          onKeydown={(event: KeyboardEvent) => applyOnEnter(event, enabled, state.apply)}
        >
          {{
            default: () => [
              <span ref={anchor} hidden />,
              ...render(defaults, slots.host?.(), props.preset),
            ],
          }}
        </DialogFrame>
      );
    };
  },
});

/**
 * The markers that make a `FormSection` or one field the dialog's `Field` part of `name`. They
 * go on the element itself, where the dialog parts find them. @internal
 */
export const fieldPart = coreFieldPart;

/**
 * A titled group of fields: a fieldset with its legend. With `fieldPart(name)`, it is the
 * dialog's `Field` part of that name. Other attributes (`data-*`, `style`) reach the fieldset.
 * @internal
 */
export const FormSection = defineComponent({
  name: 'PptxFormSection',
  props: {
    title: { type: String, required: true as const },
    className: { type: String, default: undefined },
  },
  setup(props, { slots }) {
    return () => (
      <fieldset class={['pptx-dialog__section', props.className].filter(Boolean).join(' ')}>
        <legend class="pptx-dialog__label">{props.title}</legend>
        {slots.default?.()}
      </fieldset>
    );
  },
});

/**
 * Mark a field component as a dialog control: its `Field` part has no single element to give
 * `asChild` to, so a host's own content replaces the whole field. @internal
 */
export function dialogControl<T extends object>(component: T): T {
  return Object.assign(component, { dialogControl: true });
}

const fieldClass = (base: string, className: string | undefined): string =>
  [base, className].filter(Boolean).join(' ');

/** Focus `input` on mount when `autoFocus` is set, as React's `autoFocus`. */
function useAutoFocus(input: { value: HTMLElement | null }, autoFocus: () => boolean): void {
  onMounted(() => {
    if (autoFocus()) input.value?.focus();
  });
}

/**
 * A number field with its unit inside the box. An empty value shows `placeholder`. Other
 * attributes (`data-*`, `aria-*`, `style`, `id`, `title`) reach the outer label. @internal
 */
export const NumberField = dialogControl(
  defineComponent({
    name: 'PptxNumberField',
    props: {
      name: { type: String, required: true as const },
      label: { type: String, required: true as const },
      unit: { type: String, required: true as const },
      step: { type: String, required: true as const },
      min: { type: String, default: undefined },
      max: { type: String, default: undefined },
      value: { type: String, required: true as const },
      placeholder: { type: String, default: undefined },
      disabled: { type: Boolean, default: false },
      autoFocus: { type: Boolean, default: false },
      className: { type: String, default: undefined },
      onChange: { type: Function as PropType<(value: string) => void>, required: true as const },
    },
    setup(props) {
      const input = ref<HTMLInputElement | null>(null);
      useAutoFocus(input, () => props.autoFocus);
      return () => (
        <label class={fieldClass('pptx-field', props.className)}>
          {props.label}
          {props.unit ? <span class="pptx-visually-hidden"> ({props.unit})</span> : null}
          <span class="pptx-field__control">
            <input
              ref={input}
              name={props.name}
              type="number"
              inputmode="decimal"
              step={props.step}
              min={props.min}
              max={props.max}
              // A disabled field shows no number and no placeholder.
              placeholder={props.disabled ? undefined : props.placeholder}
              disabled={props.disabled}
              autofocus={props.autoFocus || undefined}
              value={props.disabled ? '' : props.value}
              onInput={(e: Event) => props.onChange((e.target as HTMLInputElement).value)}
            />
            {props.unit ? (
              <span class="pptx-field__unit" aria-hidden="true">
                {props.unit}
              </span>
            ) : null}
          </span>
        </label>
      );
    },
  })
);

/** A one-line text field with its label. @internal */
export const TextField = defineComponent({
  name: 'PptxTextField',
  props: {
    name: { type: String, required: true as const },
    label: { type: String, required: true as const },
    value: { type: String, required: true as const },
    disabled: { type: Boolean, default: false },
    /** Show the text in this font (a bullet character in its bullet font). */
    font: { type: String, default: undefined },
    /**
     * What the text shows as, inside the box at its end: a symbol font's character as its
     * Unicode twin, which reads the same where the symbol font is not installed.
     */
    preview: { type: String, default: undefined },
    className: { type: String, default: undefined },
    onChange: { type: Function as PropType<(value: string) => void>, required: true as const },
  },
  setup(props) {
    return () => (
      <label class={fieldClass('pptx-field', props.className)}>
        {props.label}
        <span class="pptx-field__control">
          <input
            name={props.name}
            type="text"
            spellcheck={false}
            disabled={props.disabled}
            value={props.value}
            style={props.font ? { fontFamily: cssFontFamily(props.font) } : undefined}
            onInput={(e: Event) => props.onChange((e.target as HTMLInputElement).value)}
          />
          {props.preview && !props.disabled ? (
            <span class="pptx-field__unit" aria-hidden="true" data-preview="">
              {props.preview}
            </span>
          ) : null}
        </span>
      </label>
    );
  },
});

/**
 * A select with its label. `options` are value and label pairs. With `mixed`, an empty value
 * shows as that label (the selection disagrees and nothing is chosen yet). @internal
 */
export const SelectField = dialogControl(
  defineComponent({
    name: 'PptxSelectField',
    props: {
      name: { type: String, required: true as const },
      label: { type: String, required: true as const },
      /** Keep the label for screen readers only (a unit list beside its number). */
      hideLabel: { type: Boolean, default: false },
      value: { type: String, required: true as const },
      options: {
        type: Array as PropType<readonly (readonly [string, string])[]>,
        required: true as const,
      },
      mixed: { type: String, default: undefined },
      disabled: { type: Boolean, default: false },
      autoFocus: { type: Boolean, default: false },
      className: { type: String, default: undefined },
      // The value type is the caller's: Vue has no generic components over runtime props.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      onChange: { type: Function as PropType<(value: any) => void>, required: true as const },
    },
    setup(props) {
      const select = ref<HTMLSelectElement | null>(null);
      useAutoFocus(select, () => props.autoFocus);
      return () => (
        <label class={fieldClass('pptx-field', props.className)}>
          {props.hideLabel ? <span class="pptx-visually-hidden">{props.label}</span> : props.label}
          <select
            ref={select}
            name={props.name}
            value={props.value}
            disabled={props.disabled}
            autofocus={props.autoFocus || undefined}
            onChange={(e: Event) => props.onChange((e.target as HTMLSelectElement).value)}
          >
            {props.value === '' && props.mixed !== undefined ? (
              <option value="" disabled>
                {props.mixed}
              </option>
            ) : null}
            {props.options.map(([value, label]) => (
              <option key={value} value={value}>
                {label}
              </option>
            ))}
          </select>
        </label>
      );
    },
  })
);

/**
 * The tabs of a form dialog, with the ARIA tabs pattern: a labelled tab list, one tab in the Tab
 * order, the arrow keys, Home, and End to move, and each tab tied to its panel. @internal
 */
export const FormTabs = defineComponent({
  name: 'PptxFormTabs',
  props: {
    label: { type: String, required: true as const },
    /** Prefix for the tab and panel ids (`useId()`). */
    idBase: { type: String, required: true as const },
    tabs: {
      type: Array as PropType<readonly (readonly [string, string])[]>,
      required: true as const,
    },
    value: { type: String, required: true as const },
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    onChange: { type: Function as PropType<(tab: any) => void>, required: true as const },
  },
  setup(props) {
    const move = (e: KeyboardEvent): void => {
      const ids = props.tabs.map(([id]) => id);
      const next = formTabStep(e.key, ids.indexOf(props.value), ids.length);
      if (next < 0) return;
      e.preventDefault();
      const tab = ids[next]!;
      props.onChange(tab);
      (e.currentTarget as HTMLElement).querySelector<HTMLElement>(`[data-tab="${tab}"]`)?.focus();
    };
    return () => (
      <div class="pptx-segmented" role="tablist" aria-label={props.label} onKeydown={move}>
        {props.tabs.map(([id, label]) => (
          <button
            key={id}
            id={formTabId(props.idBase, id)}
            type="button"
            role="tab"
            aria-selected={props.value === id}
            aria-controls={formPanelId(props.idBase, id)}
            tabindex={props.value === id ? 0 : -1}
            data-tab={id}
            onClick={() => props.onChange(id)}
          >
            {label}
          </button>
        ))}
      </div>
    );
  },
});

/** The panel of the selected tab. @internal */
export const FormTabPanel = defineComponent({
  name: 'PptxFormTabPanel',
  props: {
    idBase: { type: String, required: true as const },
    tab: { type: String, required: true as const },
    /** A tab that is not shown: it keeps its space (`.pptx-dialog__panels`) but is hidden. */
    inactive: { type: Boolean, default: false },
  },
  setup(props, { slots }) {
    return () => (
      <div
        role="tabpanel"
        id={formPanelId(props.idBase, props.tab)}
        aria-labelledby={formTabId(props.idBase, props.tab)}
        class="pptx-dialog__rows"
        {...(props.inactive ? { 'data-inactive': '', 'aria-hidden': true } : {})}
      >
        {slots.default?.()}
      </div>
    );
  },
});

/** A checkbox with a mixed state (the selected text disagrees), shown as indeterminate. @internal */
export const CheckField = dialogControl(
  defineComponent({
    name: 'PptxCheckField',
    props: {
      name: { type: String, required: true as const },
      label: { type: String, required: true as const },
      value: { type: [Boolean, String] as PropType<boolean | 'mixed'>, required: true as const },
      disabled: { type: Boolean, default: false },
      className: { type: String, default: undefined },
      onChange: { type: Function as PropType<(on: boolean) => void>, required: true as const },
    },
    setup(props) {
      const input = ref<HTMLInputElement | null>(null);
      watchEffect(
        () => {
          if (input.value) input.value.indeterminate = props.value === 'mixed';
        },
        { flush: 'post' }
      );
      return () => (
        <label class={fieldClass('pptx-radio', props.className)}>
          <input
            ref={input}
            type="checkbox"
            name={props.name}
            checked={props.value === true}
            disabled={props.disabled}
            onChange={(e: Event) => props.onChange((e.target as HTMLInputElement).checked)}
          />
          {props.label}
        </label>
      );
    },
  })
);
