// Compile-time consumer checks for history group handles, adapted from DOCX Editor's
// `packages/core/src/__tests__/history-group.test-d.ts` at be94e1c51. `bun run typecheck` checks this file;
// a failed `@ts-expect-error` fails the gate.
import type { Editor, EditorCommand, HistoryGroup } from '../contracts/index.ts';
import type { HistoryGroup as StoreHistoryGroup, TransactOptions } from '../store/index.ts';
import { runToolbarCommand } from '../editor/chrome-controls.ts';

declare const editor: Editor;
declare const command: EditorCommand;
const historyGroup: HistoryGroup = editor.beginHistoryGroup();
editor.exec(command, { historyGroup });
editor.can(command, { historyGroup });
editor.exec(command);
const state: 'open' | 'closed' = historyGroup.state;
void state;
// Options come last, after the slot's value, as in DOCX Editor.
runToolbarCommand(editor, 'font.size', 24, { historyGroup });
runToolbarCommand(editor, 'text.bold', undefined, { historyGroup });
historyGroup.end();
// The store takes the token a handle stands for: a symbol, never a string key.
const token: StoreHistoryGroup = Symbol('drag');
const transact: TransactOptions = { historyGroup: token, selectionAfter: null };
void transact;
// @ts-expect-error a raw symbol is not an owned handle
editor.exec(command, { historyGroup: Symbol('gesture') });
// @ts-expect-error font size is a number of points
runToolbarCommand(editor, 'font.size', '24', { historyGroup });
// @ts-expect-error callers cannot forge ownership
const forged: HistoryGroup = { state: 'open', end() {} };
void forged;
// @ts-expect-error a string key is not a history group
const keyed: TransactOptions = { historyGroup: 'typing' };
void keyed;
// @ts-expect-error `mergeKey` is gone; groups replace it
const merge: TransactOptions = { mergeKey: 'typing' };
void merge;
