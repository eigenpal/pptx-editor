/** @jsxImportSource react */
// Shared parts of the form dialogs (Paragraph, Font, Bullets, and Columns): the modal frame with
// OK and Cancel on the dialog named parts (`dialog-parts.tsx`), a draft that seeds once when the
// dialog opens, and number, text, and select fields with a unit, a "Mixed" blank, and an error
// state. Internal: each dialog exports its own parts.

import {
  useEffect,
  useId,
  useRef,
  useState,
  type CSSProperties,
  type KeyboardEvent as ReactKeyboardEvent,
  type ReactNode,
} from 'react';
import {
  cssFontFamily,
  dialogErrorsOf,
  focusDialogInitialField,
  fieldPart as coreFieldPart,
  formPanelId,
  formTabId,
  formTabStep,
  markDialogError,
  type DialogError as CoreDialogError,
} from '@pptx-editor.dev/core/editor';
import { ToolButton } from './chrome.tsx';
import { useEditorState, useTranslation } from './context.ts';
import {
  DialogFrame,
  applyOnEnter,
  type DialogParts,
  type UseDialogReturn,
} from './dialog-parts.tsx';

/** A problem the dialog shows: the message, and the field to mark and focus (by `name`). */
export type DialogError = CoreDialogError;

/**
 * A draft of a dialog's fields. It seeds from `seedOf` when the dialog opens and never while it
 * is open, so a remote edit does not overwrite what the user types.
 */
export function useDialogDraft<D>(open: boolean, seedOf: () => D) {
  const seed = useRef<D | null>(null);
  const [draft, setDraft] = useState<D | null>(null);
  const [error, setError] = useState<DialogError | null>(null);
  const seedRef = useRef(seedOf);
  seedRef.current = seedOf;
  useEffect(() => {
    if (!open) {
      seed.current = null;
      setDraft(null);
      setError(null);
      return;
    }
    if (seed.current) return;
    seed.current = seedRef.current();
    setDraft(seed.current);
  }, [open]);
  const set = <K extends keyof D>(key: K, value: D[K]): void => {
    setDraft((d) => (d ? { ...d, [key]: value } : d));
    setError(null);
  };
  /** Change several fields at once (a pair of controls over one value). */
  const update = (change: (d: D) => D): void => {
    setDraft((d) => (d ? change(d) : d));
    setError(null);
  };
  return { seed: seed.current, draft, set, update, error, setError };
}

/** The named parts a form dialog exposes: `createDialogParts` over its draft. */
export type FormDialogParts<Name extends string, D extends object> = DialogParts<
  Name,
  UseDialogReturn<D>
>;

/**
 * The draft and actions a form dialog hands to its parts (`useParagraphDialog()` and the
 * like), in the shape of DOCX Editor's `UseDialogReturn`.
 */
export function formDialogState<D extends object>(args: {
  readonly draft: D | null;
  readonly set: <K extends keyof D>(key: K, value: D[K]) => void;
  readonly error: DialogError | null;
  readonly isEnabled: boolean;
  readonly onSubmit: () => void;
  readonly onClose: () => void;
}): UseDialogReturn<D> | null {
  const { draft, error } = args;
  if (!draft) return null;
  const errors = dialogErrorsOf(error);
  return {
    values: draft,
    setValue: args.set,
    errors: errors as UseDialogReturn<D>['errors'],
    isEnabled: args.isEnabled,
    apply: args.onSubmit,
    cancel: args.onClose,
  };
}

/**
 * The modal frame of a form dialog on the dialog named parts (`dialog-parts.tsx`): `Header`,
 * `Title`, `Body`, `Footer`, `Field` (each section with a `name`), `Error`, `Apply`, and
 * `Cancel`, each replaceable by a host child, as in DOCX Editor. The frame is main's
 * `DialogFrame`: modal, Escape and backdrop close, Tab trap, focus return.
 */
export function FormDialog<Name extends string, D extends object>(props: {
  /** `data-pptx-dialog`: the registry key. */
  readonly name: string;
  /** A host class, after `pptx-dialog`. */
  readonly className?: string | undefined;
  readonly style?: CSSProperties | undefined;
  readonly title: string;
  readonly open: boolean;
  readonly okLabel: string;
  readonly cancelLabel: string;
  /** The accessible name of the header's close button. */
  readonly closeLabel: string;
  readonly error: DialogError | null;
  readonly onClose: () => void;
  readonly onSubmit: () => void;
  /** The dialog's parts, and the draft they read; null until the draft has seeded. */
  readonly parts: FormDialogParts<Name, D>;
  readonly state: UseDialogReturn<D> | null;
  /** `false`: only the host's children (`preset` of `DialogCustomizationProps`). */
  readonly preset?: boolean | undefined;
  /** The host's named parts and extra content. */
  readonly hostChildren?: ReactNode;
  /** The dialog's fields: `FormSection`s, each a `Field` part when it has a `name`. */
  readonly children: ReactNode;
}) {
  const anchorRef = useRef<HTMLSpanElement>(null);
  const errorId = useId();
  const { t } = useTranslation();
  const viewing = useEditorState((s) => s.mode === 'view');
  const { open, onClose, error, state } = props;
  const ready = state !== null;
  // Mark the field the error names, tie it to the message, and move focus there. The dialog
  // carries `data-invalid` while an error shows, for host styles.
  useEffect(() => {
    markDialogError(anchorRef.current?.closest('dialog'), error, errorId);
  }, [error, errorId, ready]);
  // The fields mount after the draft seeds, after the frame focused its (empty) panel: focus the
  // first field then, unless a field took focus itself.
  useEffect(() => {
    const dialog = anchorRef.current?.closest('dialog');
    if (ready && dialog) focusDialogInitialField(dialog);
  }, [ready]);
  if (!open) return null;
  const { parts } = props;
  return (
    <DialogFrame
      kind={props.name}
      label={props.title}
      className={props.className}
      style={props.style}
      onClose={onClose}
      onKeyDown={(event) => applyOnEnter(event, state?.isEnabled ?? false, props.onSubmit)}
    >
      <span ref={anchorRef} hidden />
      {state ? (
        <parts.Composition
          state={state}
          preset={props.preset ?? true}
          defaults={
            <>
              <div data-pptx-part="header" className="pptx-dialog__header">
                <h2 data-pptx-part="title" className="pptx-dialog__title">
                  {props.title}
                </h2>
                <ToolButton label={props.closeLabel} icon="close" onClick={onClose} />
              </div>
              <div data-pptx-part="body" className="pptx-dialog__body">
                {props.children}
              </div>
              <div data-pptx-part="footer" className="pptx-dialog__footer">
                <span
                  id={errorId}
                  data-pptx-part="error"
                  role="alert"
                  className="pptx-dialog__error"
                >
                  {error?.message ?? null}
                </span>
                <button
                  type="button"
                  data-pptx-part="cancel"
                  className="pptx-button"
                  onClick={onClose}
                >
                  {props.cancelLabel}
                </button>
                <button
                  type="button"
                  data-pptx-part="apply"
                  className="pptx-button pptx-button--primary"
                  disabled={!state.isEnabled}
                  // A disabled OK says why, as every disabled control does.
                  title={
                    state.isEnabled
                      ? undefined
                      : t(viewing ? 'disabled.viewing' : 'disabled.noEditor')
                  }
                  onClick={props.onSubmit}
                >
                  {props.okLabel}
                </button>
              </div>
            </>
          }
        >
          {props.hostChildren}
        </parts.Composition>
      ) : null}
    </DialogFrame>
  );
}

/**
 * The markers that make a `FormSection` or one field (`NumberField`, `SelectField`) the
 * dialog's `Field` part of `name`. They go on the element itself, where the dialog parts find
 * them. A section and each field in it are parts, as DOCX Editor names a part per field.
 */
export const fieldPart = coreFieldPart;

/**
 * A titled group of fields: a fieldset with its legend. With `fieldPart(name)`, it is the
 * dialog's `Field` part of that name: a host part adds its class, style, and attributes here.
 */
export function FormSection({
  title,
  children,
  className,
  style,
  ...rest
}: {
  readonly title: string;
  readonly children: ReactNode;
  readonly className?: string;
  readonly style?: CSSProperties;
  readonly [attribute: `data-${string}`]: unknown;
}) {
  return (
    <fieldset
      {...rest}
      className={['pptx-dialog__section', className].filter(Boolean).join(' ')}
      style={style}
    >
      <legend className="pptx-dialog__label">{title}</legend>
      {children}
    </fieldset>
  );
}

/**
 * What a field's `Field` part adds to its outer element: a class, a style, an id, a title, and
 * `data-*` and `aria-*` attributes.
 */
interface FieldPartAttributes {
  readonly className?: string;
  readonly style?: CSSProperties;
  readonly id?: string;
  readonly title?: string;
  readonly [attribute: `data-${string}`]: unknown;
  readonly [attribute: `aria-${string}`]: unknown;
}

/** The part attributes of a field's props, for its outer element. */
function partAttributes(props: object): FieldPartAttributes {
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(props))
    if (
      key.startsWith('data-') ||
      key.startsWith('aria-') ||
      key === 'style' ||
      key === 'id' ||
      key === 'title'
    )
      out[key] = value;
  return out;
}

/**
 * Mark a field component as a dialog control: its `Field` part has no single element to give
 * `asChild` to, so a host's own content replaces the whole field (see `dialog-parts.tsx`).
 */
export function dialogControl<T extends object>(component: T): T {
  return Object.assign(component, { dialogControl: true });
}

const fieldClass = (base: string, props: FieldPartAttributes): string =>
  [base, props.className].filter(Boolean).join(' ');

/** A number field with its unit inside the box. An empty value shows `placeholder`. */
export function NumberField(
  props: FieldPartAttributes & {
    name: string;
    label: string;
    unit: string;
    step: string;
    min?: string;
    max?: string;
    value: string;
    placeholder?: string;
    disabled?: boolean;
    autoFocus?: boolean;
    onChange: (value: string) => void;
  }
) {
  return (
    <label {...partAttributes(props)} className={fieldClass('pptx-field', props)}>
      {props.label}
      {props.unit ? <span className="pptx-visually-hidden"> ({props.unit})</span> : null}
      <span className="pptx-field__control">
        <input
          name={props.name}
          type="number"
          inputMode="decimal"
          step={props.step}
          min={props.min}
          max={props.max}
          // A disabled field shows no number and no placeholder.
          placeholder={props.disabled ? undefined : props.placeholder}
          disabled={props.disabled}
          autoFocus={props.autoFocus}
          value={props.disabled ? '' : props.value}
          onChange={(e) => props.onChange(e.target.value)}
        />
        {props.unit ? (
          <span className="pptx-field__unit" aria-hidden="true">
            {props.unit}
          </span>
        ) : null}
      </span>
    </label>
  );
}

/** A one-line text field with its label. */
export function TextField(
  props: FieldPartAttributes & {
    name: string;
    label: string;
    value: string;
    disabled?: boolean;
    /** Show the text in this font (a bullet character in its bullet font). */
    font?: string;
    /**
     * What the text shows as, inside the box at its end: a symbol font's character as its
     * Unicode twin, which reads the same where the symbol font is not installed.
     */
    preview?: string;
    onChange: (value: string) => void;
  }
) {
  return (
    <label {...partAttributes(props)} className={fieldClass('pptx-field', props)}>
      {props.label}
      <span className="pptx-field__control">
        <input
          name={props.name}
          type="text"
          spellCheck={false}
          disabled={props.disabled}
          value={props.value}
          style={props.font ? { fontFamily: cssFontFamily(props.font) } : undefined}
          onChange={(e) => props.onChange(e.target.value)}
        />
        {props.preview && !props.disabled ? (
          <span className="pptx-field__unit" aria-hidden="true" data-preview="">
            {props.preview}
          </span>
        ) : null}
      </span>
    </label>
  );
}

/**
 * A select with its label. `options` are value and label pairs. With `mixed`, an empty value
 * shows as that label (the selection disagrees and nothing is chosen yet).
 */
export function SelectField<V extends string>(
  props: FieldPartAttributes & {
    name: string;
    label: string;
    /** Keep the label for screen readers only (a unit list beside its number). */
    hideLabel?: boolean;
    value: V | '';
    options: readonly (readonly [V, string])[];
    mixed?: string;
    disabled?: boolean;
    autoFocus?: boolean;
    onChange: (value: V) => void;
  }
) {
  return (
    <label {...partAttributes(props)} className={fieldClass('pptx-field', props)}>
      {props.hideLabel ? <span className="pptx-visually-hidden">{props.label}</span> : props.label}
      <select
        name={props.name}
        value={props.value}
        disabled={props.disabled}
        autoFocus={props.autoFocus}
        onChange={(e) => props.onChange(e.target.value as V)}
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
}

/**
 * The tabs of a form dialog, with the ARIA tabs pattern: a labelled tab list, one tab in the Tab
 * order, the arrow keys, Home, and End to move, and each tab tied to its panel.
 */
export function FormTabs<T extends string>(props: {
  readonly label: string;
  /** Prefix for the tab and panel ids (`useId()`). */
  readonly idBase: string;
  readonly tabs: readonly (readonly [T, string])[];
  readonly value: T;
  readonly onChange: (tab: T) => void;
}) {
  const ids = props.tabs.map(([id]) => id);
  const move = (e: ReactKeyboardEvent<HTMLDivElement>): void => {
    const next = formTabStep(e.key, ids.indexOf(props.value), ids.length);
    if (next < 0) return;
    e.preventDefault();
    const tab = ids[next]!;
    props.onChange(tab);
    e.currentTarget.querySelector<HTMLElement>(`[data-tab="${tab}"]`)?.focus();
  };
  return (
    <div className="pptx-segmented" role="tablist" aria-label={props.label} onKeyDown={move}>
      {props.tabs.map(([id, label]) => (
        <button
          key={id}
          id={tabId(props.idBase, id)}
          type="button"
          role="tab"
          aria-selected={props.value === id}
          aria-controls={panelId(props.idBase, id)}
          tabIndex={props.value === id ? 0 : -1}
          data-tab={id}
          onClick={() => props.onChange(id)}
        >
          {label}
        </button>
      ))}
    </div>
  );
}

/** The panel of the selected tab. */
export function FormTabPanel(props: {
  readonly idBase: string;
  readonly tab: string;
  /** A tab that is not shown: it keeps its space (`.pptx-dialog__panels`) but is hidden. */
  readonly inactive?: boolean;
  readonly children: ReactNode;
}) {
  return (
    <div
      role="tabpanel"
      id={panelId(props.idBase, props.tab)}
      aria-labelledby={tabId(props.idBase, props.tab)}
      className="pptx-dialog__rows"
      {...(props.inactive ? { 'data-inactive': '', 'aria-hidden': true } : {})}
    >
      {props.children}
    </div>
  );
}

const tabId = formTabId;
const panelId = formPanelId;

/** A checkbox with a mixed state (the selected text disagrees), shown as indeterminate. */
export function CheckField(
  props: FieldPartAttributes & {
    readonly name: string;
    readonly label: string;
    readonly value: boolean | 'mixed';
    readonly onChange: (on: boolean) => void;
  }
) {
  const ref = useRef<HTMLInputElement>(null);
  useEffect(() => {
    if (ref.current) ref.current.indeterminate = props.value === 'mixed';
  }, [props.value]);
  return (
    <label {...partAttributes(props)} className={fieldClass('pptx-radio', props)}>
      <input
        ref={ref}
        type="checkbox"
        name={props.name}
        checked={props.value === true}
        onChange={(e) => props.onChange(e.target.checked)}
      />
      {props.label}
    </label>
  );
}

dialogControl(NumberField);
dialogControl(SelectField);
dialogControl(CheckField);
