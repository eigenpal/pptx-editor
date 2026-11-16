// `getDocumentHandle()` and the `replaceDocument` command, after DOCX Editor's document handle
// and `replace-document-edit.ts` (docx-editor/packages/core/src/editor at b63cbea4a).
//
// The command reads the returned file through `readPptx` and its limits, takes the difference
// from the open package (`diffPackages`), and commits it as one store transaction. A change
// made on the store directly goes through the facade's store listener, which keeps the slide on
// screen and the selection valid, as for a host's own store edit. Parts the file did not change
// keep their identity, so the cascade, the layout, and the thumbnails reuse them: nothing on
// screen paints again for them, and nothing flashes.
//
// A document refresh that reloads applies the same difference in place, with no history entry,
// and then clears the history and the selection, as a load does (`reloadInPlace`). It gives the
// document a new load id. A file whose difference touches embedded fonts reloads through
// `load()` instead, because a session registers its embedded faces when it opens.

import type {
  DocumentHandle,
  ReplaceDocumentCommand,
  ReplaceDocumentOutcome,
  ReplaceDocumentRefusal,
} from '../contracts/index.ts';
import {
  diffPackages,
  readPptx,
  slideRefAt,
  type PptxOp,
  type PptxPackage,
} from '../store/index.ts';
import { markOf } from './history-marks.ts';
import { isCollaborating } from './modules.ts';
import type { DocumentSession } from './document-session.ts';
import type { PptxEditorInstance } from './pptx-editor-types.ts';
import { withChangeSource } from './change-source.ts';

type Translate = (key: string, params?: Record<string, string | number>) => string;
const translations = new WeakMap<object, Translate>();

/** The editor's `translate` option, for chrome that core draws itself. @internal */
export function translateOf(editor: object): Translate | undefined {
  return translations.get(editor);
}

const loadIds = new WeakMap<DocumentSession, string>();
let nextLoad = 0;
/** Editors with a refresh or a preview that blocks replacement (`busy`). */
const busyEditors = new WeakSet<object>();

/** The load id of a session: one per open, and a new one after a refresh that reloads. */
function loadIdOf(session: DocumentSession): string {
  let id = loadIds.get(session);
  if (!id) loadIds.set(session, (id = `load-${++nextLoad}`));
  return id;
}

/** The open document's handle. @internal */
export function documentHandleOf(editor: Pick<PptxEditorInstance, 'session'>): DocumentHandle {
  const session = editor.session;
  return session
    ? { loadId: loadIdOf(session), revision: session.store.revision }
    : { loadId: '', revision: 0 };
}

/** Mark `editor` busy while `run` replaces its document. @internal */
export async function whileReplacing<T>(editor: object, run: () => Promise<T>): Promise<T> {
  busyEditors.add(editor);
  try {
    return await run();
  } finally {
    busyEditors.delete(editor);
  }
}

const refusal = (
  code: ReplaceDocumentRefusal['code'],
  message: string,
  reason?: string
): ReplaceDocumentRefusal =>
  ({ ok: false, code, message, ...(reason ? { reason } : {}) }) as ReplaceDocumentRefusal;

/** The returned file as a package, or a refusal. */
function readReturned(bytes: unknown): PptxPackage | ReplaceDocumentRefusal {
  if (!(bytes instanceof Uint8Array) && !(bytes instanceof ArrayBuffer))
    return refusal('invalidRequest', 'pptx must be a Uint8Array or an ArrayBuffer.');
  const read = readPptx(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
  return read.ok ? read.pkg : refusal('invalidDocument', 'The bytes are not a valid PPTX file.');
}

/** Whether the ops change embedded fonts, which a session registers only when it opens. */
const touchesFonts = (ops: readonly PptxOp[]): boolean =>
  ops.some((op) => /^\/ppt\/fonts\//i.test(op.part));

interface Prepared {
  readonly session: DocumentSession;
  readonly ops: readonly PptxOp[];
}

/** Check the editor and the request, and compute the difference. */
function prepare(
  editor: PptxEditorInstance,
  request: Pick<ReplaceDocumentCommand, 'pptx' | 'expected'>,
  own = false
): Prepared | ReplaceDocumentRefusal {
  const session = editor.session;
  if (!session || editor.snapshot().isOpening)
    return refusal('unavailable', 'No document is mounted, or it is still opening.');
  // A refresh blocks other replacements, not its own.
  if (busyEditors.has(editor) && !own) return refusal('busy', 'A document refresh blocks edits.');
  if (editor.snapshot().mode === 'view' || session.store.blocked !== null)
    return refusal('editNotPermitted', 'The editor is in viewing mode.', 'viewing');
  const expected = request.expected;
  if (expected !== undefined) {
    if (!expected || typeof expected.loadId !== 'string' || typeof expected.revision !== 'number')
      return refusal('invalidRequest', 'expected must be a document handle.');
    const now = documentHandleOf(editor);
    if (expected.loadId !== now.loadId)
      return refusal('staleDocument', 'Another document was loaded.', 'otherDocument');
    if (expected.revision !== now.revision)
      return refusal('staleDocument', 'The document changed.', 'edited');
  }
  const next = readReturned(request.pptx);
  if ('ok' in next) return next;
  // A deck still opening parses the rest first: the difference reads every part.
  if (!editor.finishOpening())
    return refusal('unavailable', 'A slide of the open document did not parse.');
  const diff = diffPackages(session.store.package, next);
  if (!diff.ok)
    return diff.reason === 'different-main-part'
      ? refusal('editUnsupported', 'The file names a different main part.', 'differentMainPart')
      : refusal('editUnsupported', 'The difference is not an edit.', 'unsupportedDifference');
  return { session, ops: diff.ops };
}

/** Commit the difference. `history: false` records no undo step. */
function commit(
  editor: PptxEditorInstance,
  { session, ops }: Prepared,
  history: boolean
): ReplaceDocumentOutcome {
  if (ops.length === 0) return { ok: true, changed: false };
  const snap = editor.snapshot();
  const shown = slideRefAt(session.store.package, snap.currentSlide)?.partName;
  const mark = markOf(shown, shown, snap.selection);
  const r = session.store.transact(ops, {
    label: 'replaceDocument',
    exact: true,
    recordsHistory: history,
    selectionBefore: mark,
  });
  if (r.ok) return { ok: true, changed: true };
  // A collaboration room refuses through the store too; anything else is a difference the part
  // checks do not accept.
  return isCollaborating(editor) && r.rejection.code === 'unsupported'
    ? refusal('editUnsupported', 'The collaboration room refused the edit.', 'collaborationRefused')
    : refusal(
        'editUnsupported',
        'The difference cannot apply as an edit.',
        'unsupportedDifference'
      );
}

/** Run the `replaceDocument` command. @internal */
export function replaceDocument(
  editor: PptxEditorInstance,
  command: ReplaceDocumentCommand
): ReplaceDocumentOutcome {
  const prepared = prepare(editor, command);
  if ('ok' in prepared) return prepared;
  return withChangeSource(editor, 'refresh', () => commit(editor, prepared, true));
}

/**
 * Show a returned file as a reload does: the content is the file's, and the history and the
 * selection are reset. The difference applies in place, so unchanged slides and thumbnails do
 * not paint again and no loading screen shows. Returns `'load'` when the editor must open the
 * file with `load()` instead (embedded fonts changed, or the difference is no edit).
 *
 * @internal
 */
export function reloadInPlace(
  editor: PptxEditorInstance,
  bytes: Uint8Array,
  source: 'refresh' | 'recovery'
): ReplaceDocumentOutcome | 'load' {
  const prepared = prepare(editor, { pptx: bytes }, true);
  if ('ok' in prepared)
    return prepared.code === 'invalidDocument' || prepared.code === 'busy' ? prepared : 'load';
  if (touchesFonts(prepared.ops)) return 'load';
  return withChangeSource(editor, source, () => {
    const { session } = prepared;
    const index = editor.snapshot().currentSlide;
    editor.setSelection({ kind: 'none' });
    const outcome = commit(editor, prepared, false);
    if (!outcome.ok) return outcome;
    // A reload keeps the slide index, clamped, as a load with `startSlide` does.
    if (editor.snapshot().currentSlide !== index) editor.goToSlide(index);
    // A reload is a new document with no history: handles taken before it name another load.
    session.store.clearHistory();
    session.store.closeHistoryGroup();
    loadIds.set(session, `load-${++nextLoad}`);
    // Publish the empty history and the new load id.
    editor.setMode(editor.snapshot().mode);
    return outcome;
  });
}

/**
 * Add `getDocumentHandle()` and the `replaceDocument` command to an editor. The facade calls it
 * once, on the instance it returns.
 *
 * @internal
 */
export function withDocumentReplace(
  editor: PptxEditorInstance,
  translate?: Translate
): PptxEditorInstance {
  if (translate) translations.set(editor, translate);
  const exec = editor.exec;
  editor.exec = (command, options) => {
    if (command.type !== 'replaceDocument') return exec(command, options);
    if (options?.historyGroup !== undefined) {
      const outcome = refusal('invalidRequest', 'A replacement is always its own undo step.');
      return { ok: false, reason: 'errors.unsupported', replaceDocument: outcome };
    }
    const outcome = replaceDocument(editor, command);
    return outcome.ok
      ? { ok: true, replaceDocument: outcome }
      : { ok: false, reason: 'errors.unsupported', detail: outcome.code, replaceDocument: outcome };
  };
  return editor;
}
