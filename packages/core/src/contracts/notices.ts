// What the edit notice reports: a refused edit (`snapshot.lastRejection`) and the last Delete
// the user can undo (`snapshot.lastRemoval`). Moved out of `index.ts` (the 1000-line cap).

/** A Delete that the user can undo: what it removed and how many. @public */
export interface EditorRemoval {
  readonly kind: 'shapes' | 'slides';
  readonly count: number;
  /** Grows with each removal, so chrome tells two equal removals apart. */
  readonly serial: number;
}

/** A refused edit: the i18n key of the reason and an optional detail. @public */
export interface EditorRejection {
  /** The i18n key of the reason (`errors.*`, `disabled.*`). */
  readonly reason: string;
  /** What the reason names, for the `{reason}` or `{detail}` placeholder. */
  readonly detail?: string;
  /**
   * True when the command had nothing to do (`disabled.nothingToUndo`,
   * `disabled.nothingToRedo`, `disabled.noSelection`), not when an edit was refused. Chrome
   * shows no notice for it.
   */
  readonly noop?: boolean;
}

/** Whether a command can run now; when it cannot, the i18n key of the reason. @public */
export type CanResult = { readonly ok: true } | { readonly ok: false; readonly reason: string };

/**
 * The outcome of a command. A refusal with `noop` means the command had nothing to do (nothing
 * to undo, nothing selected): nothing changed, and nothing went wrong. @public
 */
export type CommandResult =
  | { readonly ok: true }
  | {
      readonly ok: false;
      readonly reason: string;
      readonly detail?: string;
      readonly noop?: boolean;
    };
