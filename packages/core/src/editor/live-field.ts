// The edit lifecycle of one side-panel field: the slide shows each valid entry as the user types,
// Enter or blur ends the edit, and Escape restores the value that the field showed at the start.
// All writes of one edit are one undo step. React and Vue bind the same controller, so both
// adapters write, refuse, and restore at the same moments.

import type { Editor, EditorExecOptions, HistoryGroup } from '../contracts/index.ts';

/** What a live field asks of its adapter. @internal */
export interface LiveFieldHost {
  /** The editor that receives the writes. Without one, the field writes without an undo step. */
  readonly editor: Pick<Editor, 'beginHistoryGroup' | 'exec'> | null | undefined;
  /** The document value that the field shows without a draft. Null is a mixed selection. */
  value(): string | null;
  /**
   * Write `text`, passing `options` to `exec`. Return null when the text applies, a message to
   * show when it does not, or undefined to wait without a message (a partial entry while the
   * user types). `final` is true for Enter, blur, and Escape.
   */
  write(text: string, options: EditorExecOptions, final: boolean): string | null | undefined;
  /**
   * Whether `value` is what the draft wrote, so the field keeps the typed text. Without it, the
   * texts must be equal. A number field accepts a rounded document value.
   */
  same?(draft: string, value: string | null): boolean;
}

/** One side-panel field's edit state. @internal */
export interface LiveField {
  /** The typed text, or null when the field shows the document value. */
  readonly draft: string | null;
  /** The message for the typed text, or null. */
  readonly error: string | null;
  /**
   * The user typed `text`: show it, and write it when it is valid. During an IME composition,
   * pass `composing`: the field then shows the text and writes it when the composition ends.
   */
  input(text: string, composing?: boolean): void;
  /** Enter or blur: write the draft and end the edit. */
  commit(): void;
  /** Escape: restore the value from the start of the edit and end it. */
  cancel(): void;
  /**
   * The document value or the field's scope changed. Keep the draft when this edit wrote the
   * change; otherwise drop the draft and end the edit, as another source changed the value.
   * Call it during render; it does not notify.
   */
  sync(): void;
  /** End the edit and keep its writes. Call it on unmount. */
  dispose(): void;
}

/**
 * Create the edit controller of one side-panel field. `changed` runs after each state change
 * that the adapter must render. @internal
 */
export function createLiveField(host: () => LiveFieldHost, changed: () => void): LiveField {
  let draft: string | null = null;
  let error: string | null = null;
  // True after this edit wrote: the document value then follows the draft, so keep the draft.
  let editing = false;
  let written = false;
  // True during a write: the document value changes before the write returns.
  let busy = false;
  let original: string | null | undefined;
  let group: HistoryGroup | undefined;
  const options = (): EditorExecOptions => {
    const editor = host().editor;
    if (!editor) return {};
    if (!group || group.state === 'closed') group = editor.beginHistoryGroup();
    return { historyGroup: group };
  };
  const end = () => {
    group?.end();
    group = undefined;
    editing = false;
    written = false;
    original = undefined;
  };
  const clear = () => {
    draft = null;
    error = null;
  };
  const attempt = (text: string, final: boolean): boolean => {
    const current = host().value();
    if (text === current) {
      error = null;
      return true;
    }
    busy = true;
    let result: string | null | undefined;
    try {
      result = host().write(text, options(), final);
    } finally {
      busy = false;
    }
    if (result === undefined) {
      error = null;
      return false;
    }
    if (result !== null) {
      error = result;
      return false;
    }
    error = null;
    editing = true;
    written = true;
    return true;
  };
  return {
    get draft() {
      return draft;
    },
    get error() {
      return error;
    },
    input(text, composing = false) {
      if (original === undefined) original = host().value();
      draft = text;
      if (!composing) attempt(text, false);
      changed();
    },
    commit() {
      if (draft === null) {
        end();
        return;
      }
      if (attempt(draft, true)) {
        end();
        clear();
      }
      changed();
    },
    cancel() {
      const restore = original;
      const undo = written && group?.state === 'open';
      if (restore !== undefined && restore !== null && written) {
        // Write the start value in the same step: one undo step with no net change.
        if (host().write(restore, options(), true) !== null && undo)
          host().editor?.exec({ type: 'undo' });
      } else if (undo) host().editor?.exec({ type: 'undo' });
      end();
      clear();
      changed();
    },
    sync() {
      if (busy || draft === null) return;
      const value = host().value();
      const same = host().same ?? ((a: string, b: string | null) => a === b);
      if (editing && same(draft, value)) return;
      end();
      clear();
    },
    dispose() {
      end();
    },
  };
}
