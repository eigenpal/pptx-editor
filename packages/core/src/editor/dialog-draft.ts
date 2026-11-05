// Ported from DOCX Editor packages/react/src/editor/dialog-parts.tsx at 27720ad5 (the draft
// state of `UseDialogReturn`, the `data-docx-part` markers) and
// packages/core/src/editor/paragraph-dialog-fields.ts at e608e2df (`trapTabWithin`).
// The draft controller is new: DOCX Editor keeps a draft in
// each dialog. Framework-free, so the React and Vue
// adapters keep only their glue: a dialog's draft, validation, and commit behave the same in
// both.
//
// A dialog edits a DRAFT. Its fields write the draft; Apply validates it and commits it as one
// command (one undo step); Cancel discards it. A refused commit keeps the dialog open with the
// engine's reason as the `form` error, so a dialog never claims a change the engine refused.

/** A translated message per field, and `form` for the draft as a whole. @internal */
export type DialogErrors<Fields extends object> = Readonly<
  Partial<Record<Extract<keyof Fields, string> | 'form', string>>
>;

/** What committing a draft did. The reason is a translation key or the engine's text. @internal */
export type DialogCommitResult =
  { readonly ok: true } | { readonly ok: false; readonly reason: string };

/** How a draft starts, checks, and commits. @internal */
export interface DialogDraftOptions<Fields extends object> {
  /** The values the dialog opens with, read from the document. */
  readonly initial: Fields;
  /** Field errors for `values`, or an empty object when the draft can be committed. */
  readonly validate?: (values: Fields) => DialogErrors<Fields>;
  /** Write the draft to the document, as one command. */
  readonly commit: (values: Fields) => DialogCommitResult;
  /** The values Reset restores. Without it, Reset restores `initial`. */
  readonly resetValues?: () => Fields;
}

/** A dialog's draft: values, errors, and the actions over them. @internal */
export interface DialogDraft<Fields extends object> {
  /** The draft values now. A new object after every change. */
  readonly values: Fields;
  /** The errors of the last Apply, cleared field by field as the user edits. */
  readonly errors: DialogErrors<Fields>;
  /** Change one field of the draft. Nothing reaches the document before Apply. */
  setValue<K extends keyof Fields>(name: K, value: Fields[K]): void;
  /** Validate and commit the draft. True when the document took it. */
  apply(): boolean;
  /** Put the draft back to its reset values. */
  reset(): void;
  /** Be told after every change of values or errors. Returns the unsubscribe function. */
  subscribe(listener: () => void): () => void;
}

/** True when any field has a message: a key whose value is undefined is no error. */
const hasErrors = (errors: object): boolean =>
  Object.values(errors).some((message) => typeof message === 'string');

/**
 * A draft over `options.initial`. Apply runs `validate`, then `commit`; a refusal from either
 * keeps the draft and reports the errors. The adapters build their dialogs on it, as DOCX
 * Editor keeps its controller inside its adapters. @internal
 */
export function createDialogDraft<Fields extends object>(
  options: DialogDraftOptions<Fields>
): DialogDraft<Fields> {
  const none = {} as DialogErrors<Fields>;
  let values: Fields = { ...options.initial };
  let errors: DialogErrors<Fields> = none;
  const listeners = new Set<() => void>();
  const notify = (): void => {
    for (const listener of [...listeners]) listener();
  };
  return {
    get values() {
      return values;
    },
    get errors() {
      return errors;
    },
    setValue(name, value) {
      if (Object.is(values[name], value)) return;
      values = { ...values, [name]: value };
      // An edit drops the edited field's error and the form error, and any error the new values
      // no longer have (a pick that answers "choose a fill first"): the old reason no longer holds.
      if (hasErrors(errors)) {
        const still: Record<string, unknown> = options.validate?.(values) ?? none;
        const next: Record<string, string | undefined> = {};
        for (const [key, message] of Object.entries(errors))
          if (key !== name && key !== 'form' && typeof still[key] === 'string') next[key] = message;
        errors = next as DialogErrors<Fields>;
      }
      notify();
    },
    apply() {
      const invalid = options.validate?.(values) ?? none;
      if (hasErrors(invalid)) {
        errors = invalid;
        notify();
        return false;
      }
      const result = options.commit(values);
      errors = result.ok ? none : ({ form: result.reason } as DialogErrors<Fields>);
      notify();
      return result.ok;
    },
    reset() {
      values = { ...(options.resetValues?.() ?? options.initial) };
      errors = none;
      notify();
    },
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
  };
}

/**
 * The named parts of a dialog, as DOCX Editor names them. A packaged dialog marks each default
 * element with `data-pptx-part`; a host part with the same name replaces it in place.
 * @internal
 */
export type DialogPartName =
  'header' | 'title' | 'body' | 'footer' | 'field' | 'error' | 'apply' | 'cancel' | 'reset';

/** The attribute that marks a dialog's default part. @internal */
export const DIALOG_PART_ATTRIBUTE = 'data-pptx-part';

/** The attribute that names a `field` part. @internal */
export const DIALOG_FIELD_ATTRIBUTE = 'data-pptx-field';

/**
 * The key that matches a host part to a default part: the part name, or `field:<name>` for a
 * field. Null for an element that is not a part. @internal
 */
export function dialogPartKey(part: unknown, field?: unknown): string | null {
  if (typeof part !== 'string' || part === '') return null;
  return part === 'field' ? `field:${String(field)}` : part;
}

const FOCUSABLE =
  'input:not([disabled]):not([type="hidden"]), select:not([disabled]), textarea:not([disabled]), button:not([disabled]), a[href], [tabindex]:not([tabindex="-1"])';

/** The elements Tab visits inside `container`, in order. @internal */
export function dialogFocusables(container: HTMLElement): HTMLElement[] {
  // As DOCX Editor's trap: a control with tabindex -1 (an unchecked roving radio) is not a
  // tab stop.
  return [...container.querySelectorAll<HTMLElement>(FOCUSABLE)].filter(
    (el) => el.tabIndex >= 0 && !el.hidden && el.closest('[hidden]') === null
  );
}

/**
 * Keep Tab and Shift+Tab inside a dialog: from the last control to the first and back, as a
 * modal must. True when it moved focus, so the caller prevents the default. @internal
 */
export function trapTabWithin(
  container: HTMLElement,
  event: Pick<KeyboardEvent, 'key' | 'shiftKey'>
): boolean {
  if (event.key !== 'Tab') return false;
  const items = dialogFocusables(container);
  const first = items[0];
  const last = items[items.length - 1];
  if (!first || !last) return false;
  const active = container.ownerDocument.activeElement;
  const inside = active !== null && container.contains(active);
  if (event.shiftKey && (active === first || !inside)) {
    last.focus();
    return true;
  }
  if (!event.shiftKey && (active === last || !inside)) {
    first.focus();
    return true;
  }
  return false;
}

/**
 * The control a dialog focuses when it opens: the first field the user types or picks in, else
 * the first control of a `field` part (the chosen swatch of Format Background), else the
 * enabled `apply` part, else the first control outside the header. The header's Close button
 * never takes the first focus, so every dialog starts on its first field. @internal
 */
export function dialogInitialFocus(container: HTMLElement): HTMLElement | null {
  const items = dialogFocusables(container);
  return (
    // A field beside a swatch grid (the custom color's hex field, `data-menu-field`) is not the
    // first field: the grid before it is.
    items.find(
      (el) =>
        !el.hasAttribute('data-menu-field') &&
        el.matches(
          'input:not([type="checkbox"]):not([type="radio"]):not([type="file"]):not([type="button"]):not([type="submit"]):not([type="reset"]):not([type="image"]):not([type="range"]):not([type="color"]), select, textarea'
        )
    ) ??
    items.find((el) => el.closest(`[${DIALOG_PART_ATTRIBUTE}="field"]`) !== null) ??
    items.find((el) => el.getAttribute(DIALOG_PART_ATTRIBUTE) === 'apply') ??
    items.find((el) => el.closest(`[${DIALOG_PART_ATTRIBUTE}="header"]`) === null) ??
    items[0] ??
    null
  );
}

/**
 * Focus a dialog's initial field (`dialogInitialFocus`) after its fields mounted, unless a
 * control inside it took focus already. Focus on the `<dialog>` element itself does not count:
 * the frame holds it while the fields are still empty. @internal
 */
export function focusDialogInitialField(dialog: HTMLElement): void {
  const active = dialog.ownerDocument.activeElement;
  if (active && active !== dialog && dialog.contains(active)) return;
  dialogInitialFocus(dialog)?.focus({ preventScroll: true });
}

/**
 * True when a pointer event on a dialog element lies outside its box: on the backdrop, as
 * DOCX Editor's frame checks it. Both adapters use it. @internal
 */
export function isDialogBackdropPoint(
  dialog: Element,
  event: { readonly target: EventTarget | null; readonly clientX: number; readonly clientY: number }
): boolean {
  if (event.target !== dialog) return false;
  const r = dialog.getBoundingClientRect();
  return (
    event.clientX < r.left ||
    event.clientX > r.right ||
    event.clientY < r.top ||
    event.clientY > r.bottom
  );
}

/**
 * True when an event comes from a dialog opened inside `dialog` (the custom gradient editor in
 * Format Background): the outer dialog leaves its keys and its `cancel` to it. Works across
 * documents (an editor in an iframe), so it does not test `instanceof Element`. @internal
 */
export function fromNestedDialog(dialog: Element, target: EventTarget | null): boolean {
  const node = target as { closest?: (selector: string) => Element | null } | null;
  return typeof node?.closest === 'function' && node.closest('dialog') !== dialog;
}

/**
 * A part as a host writes it, for warning text: `PageSetupDialog.Field name="theme"` for the
 * key `field:theme`. Both adapters name parts this way. @internal
 */
export function dialogPartLabel(dialogName: string, key: string): string {
  // Split at the first ':' only: a field name may hold one.
  const cut = key.indexOf(':');
  const part = cut === -1 ? key : key.slice(0, cut);
  const field = cut === -1 ? undefined : key.slice(cut + 1);
  const name = `${dialogName}.${part.charAt(0).toUpperCase()}${part.slice(1)}`;
  if (field === undefined) return name;
  return field === 'undefined' || field === ''
    ? `${name} (without a name)`
    : `${name} name="${field}"`;
}

/**
 * Whether Return on `target` applies a dialog's draft, as OK does: an
 * input the user types or picks in, or a radio of a radio group. A button (also an input of a
 * button type), a link, a text area, and a file input keep their own Return. @internal
 */
export function returnAppliesFrom(target: EventTarget | null): boolean {
  const el = target as {
    localName?: string;
    type?: string;
    getAttribute?: (n: string) => string | null;
  } | null;
  if (!el || typeof el.getAttribute !== 'function') return false;
  if (el.getAttribute('role') === 'radio') return true;
  if (el.localName !== 'input') return false;
  return !['button', 'submit', 'reset', 'image', 'file'].includes(el.type ?? '');
}
