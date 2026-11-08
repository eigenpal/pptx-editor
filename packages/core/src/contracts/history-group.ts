// The public history group handle, as DOCX Editor's `packages/core/src/contracts/editor-scope.ts`
// defines it: an engine-owned identity of one gesture that a host passes to `exec`. DOCX Editor
// commit adf0e178.

declare const historyGroupBrand: unique symbol;

/**
 * The engine-owned identity of one gesture. Create it with `Editor.beginHistoryGroup()`.
 *
 * Every command that names one open handle becomes part of one undo step. Undo, redo, and a
 * command with another handle (or none) close the step, so a later command with the handle
 * starts a new one. A selection change, and a command that the editor refuses before it
 * writes, leave the step open. Redo restores the selection after the last command of the step.
 *
 * @public
 */
export interface HistoryGroup {
  readonly [historyGroupBrand]: true;
  /** Closed after `end()`, after a new document loads, and after `destroy()`. */
  readonly state: 'open' | 'closed';
  /** End this gesture without reverting any edits. Safe to call more than once. */
  end(): void;
}

/**
 * Options for a command that can join an explicit history group.
 *
 * @example
 * ```ts
 * const resize = editor.beginHistoryGroup();
 * editor.exec({ type: 'setTableGrid', columns: [914400, 914400] }, { historyGroup: resize });
 * editor.exec({ type: 'setTableGrid', columns: [1828800, 914400] }, { historyGroup: resize });
 * resize.end();
 * editor.exec({ type: 'undo' }); // reverts both commands
 * ```
 * @public
 */
export interface HistoryGroupOptions {
  /**
   * An open handle from `Editor.beginHistoryGroup` on this editor. A closed handle, or a
   * handle from another editor, refuses the command with `errors.unsupported` and changes
   * nothing.
   */
  readonly historyGroup?: HistoryGroup;
}

/** Options for one command. @public */
export type EditorExecOptions = HistoryGroupOptions;
