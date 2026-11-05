// The draft of the Paragraph, Font, Bullets and Numbering, and Columns dialogs: the Vue twin of
// the React adapter's form `useDialogDraft` (`dialog-form.tsx`). It seeds once, when the
// dialog's open component mounts, on core's draft (`createDialogDraft`), and it behaves as the
// React draft: any edit clears the error, a field problem names the field to mark, and a
// refused command keeps the dialog open with the engine's reason. Internal.

import { shallowRef, type ShallowRef } from 'vue';
import {
  formCommandRejection,
  type DialogError,
  type DialogErrors,
  type EditorCommand,
} from '@pptx-editor.dev/core/editor';
import { usePptxEditor, useTranslation } from './composables.ts';
import { useDialogDraft } from './dialog-parts.ts';

/** A form dialog's draft over its seed, with the React draft's error rules. @internal */
export interface FormDraft<D extends object> {
  /** The values the dialog opened with. */
  readonly seed: D;
  readonly values: ShallowRef<D>;
  /** The error shown, per field `name` (`form` for the whole draft). */
  readonly errors: ShallowRef<DialogErrors<D>>;
  setValue<K extends keyof D>(name: K, value: D[K]): void;
  /** Change several fields at once (a pair of controls over one value). */
  update(change: (draft: D) => D): void;
  /** Hide the error (a tab change), as React's `setError(null)`. */
  clearErrors(): void;
  apply(): boolean;
  reset(): void;
}

/**
 * The draft of a form dialog. `problem` is the first field OK cannot apply, or null; `command`
 * is what OK sends, or null when nothing changed (OK then only closes). @internal
 */
export function useFormDraft<D extends object>(options: {
  readonly seed: D;
  readonly problem: (values: D) => DialogError | null;
  readonly command: (values: D) => EditorCommand | null;
}): FormDraft<D> {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const base = useDialogDraft<D>({
    initial: options.seed,
    commit: (values) => {
      const command = options.command(values);
      if (!command) return { ok: true };
      const refused = formCommandRejection(t, editor.value?.exec(command));
      return refused === null ? { ok: true } : { ok: false, reason: refused };
    },
  });
  const none = {} as DialogErrors<D>;
  // The errors of the last OK; any edit clears them, as in the React twin.
  const errors = shallowRef<DialogErrors<D>>(none);
  const clearErrors = (): void => {
    if (errors.value !== none) errors.value = none;
  };
  return {
    seed: options.seed,
    values: base.values,
    errors,
    setValue(name, value) {
      base.setValue(name, value);
      clearErrors();
    },
    update(change) {
      const now = base.values.value;
      const next = change(now);
      for (const key of Object.keys(next) as (keyof D)[])
        if (!Object.is(now[key], next[key])) base.setValue(key, next[key]);
      clearErrors();
    },
    clearErrors,
    apply() {
      const problem = options.problem(base.values.value);
      if (problem) {
        errors.value = { [problem.field ?? 'form']: problem.message } as DialogErrors<D>;
        return false;
      }
      const ok = base.apply();
      errors.value = ok ? none : base.errors.value;
      return ok;
    },
    reset() {
      base.reset();
      clearErrors();
    },
  };
}
