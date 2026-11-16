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
// and then clears the history, as a load does (`reloadInPlace`). It gives the document a new
// load id. In both paths the editor follows the slide on screen by its part, also when slides
// before it come or go, and the selection stays when the nodes it names stay, as after an
// edit. A file whose difference touches embedded fonts, or that the part checks refuse as an
// edit, reloads through `load()` instead: a session registers its embedded faces when it opens,
// and a load accepts what an edit does not. For the same reason, the command refuses a
// difference that touches embedded fonts (`unsupportedDifference`).
//
// Both paths read the returned file as an open does: `readPptx`, then the sanitizers of
// `openStore` (`withoutEditorMarks`). So a replacement and a reload give the same package.
//
// A document refresh runs the steps in separate tasks, so no task holds the main thread for
// all of them: `preparseReturned` reads the file, `prediffReturned` takes the difference, and
// the commit then reuses both. The commit reuses a difference only while the open package is
// the package that the difference compares.

import type {
  DocumentHandle,
  ReplaceDocumentCommand,
  ReplaceDocumentOutcome,
  ReplaceDocumentRefusal,
} from '../contracts/index.ts';
import {
  diffPackages,
  presentationInfo,
  readPptx,
  slideRefAt,
  type PackageDiff,
  type PptxOp,
  type PptxPackage,
} from '../store/index.ts';
import { withoutEditorMarks } from '../automation/commands/table-edit-marks.ts';
import { markOf } from './history-marks.ts';
import { isCollaborating } from './modules.ts';
import type { DocumentSession } from './document-session.ts';
import type { PptxEditorInstance } from './pptx-editor-types.ts';
import { adoptLoadId, withChangeSource } from './change-source.ts';
import { focusSurfaceOf, surfaceInputOf } from './surface-registry.ts';

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
/** What runs when an editor is destroyed: a refresh controller's teardown. */
const teardowns = new WeakMap<object, Set<() => void>>();

/** Run `teardown` once when `editor` is destroyed. Returns the function that cancels it. @internal */
export function onEditorDestroy(editor: object, teardown: () => void): () => void {
  let set = teardowns.get(editor);
  if (!set) teardowns.set(editor, (set = new Set()));
  set.add(teardown);
  return () => {
    set.delete(teardown);
  };
}

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

/** The steps a refresh did already for its returned bytes: the package, and a difference. */
interface Steps {
  readonly pkg: PptxPackage;
  diff?: { readonly base: PptxPackage; readonly result: PackageDiff };
}
const steps = new WeakMap<Uint8Array, Steps>();

/**
 * Read `bytes` ahead of `reloadInPlace` or `replaceDocument` with the same array, which then
 * reuse the package. The caller must not change `bytes` after the call. Returns false when the
 * bytes are not a valid PPTX file. @internal
 */
export function preparseReturned(bytes: Uint8Array): boolean {
  if (steps.has(bytes)) return true;
  const read = readPptx(bytes);
  if (read.ok) steps.set(bytes, { pkg: withoutEditorMarks(read.pkg) });
  return read.ok;
}

/**
 * Take the difference from the open package to `bytes` (read with `preparseReturned`) ahead of
 * the commit. The commit reuses it only while the open package stays the same. @internal
 */
export function prediffReturned(editor: PptxEditorInstance, bytes: Uint8Array): void {
  const known = steps.get(bytes);
  const session = editor.session;
  if (!known || !session || editor.snapshot().isOpening || !editor.finishOpening()) return;
  const base = session.store.package;
  known.diff = { base, result: diffPackages(base, known.pkg) };
}

/**
 * The parts that the commit of `bytes` replaces whole (taken with `prediffReturned`): their new
 * node ids can name other nodes. Null when no difference from the open package is known.
 * @internal
 */
export function replacedWholeParts(
  editor: PptxEditorInstance,
  bytes: Uint8Array
): ReadonlySet<string> | null {
  const diff = steps.get(bytes)?.diff;
  if (!diff || diff.base !== editor.session?.store.package || !diff.result.ok) return null;
  const removed = new Set<string>();
  const out = new Set<string>();
  for (const op of diff.result.ops) {
    if (op.op === 'removePart') removed.add(op.part);
    else if (op.op === 'addPart' && removed.has(op.part)) out.add(op.part);
  }
  return out;
}

/** Forget the steps of `bytes`. @internal */
export function forgetReturned(bytes: Uint8Array): void {
  steps.delete(bytes);
}

/** The returned file as a package, or a refusal. */
function readReturned(bytes: unknown): PptxPackage | ReplaceDocumentRefusal {
  if (!(bytes instanceof Uint8Array) && !(bytes instanceof ArrayBuffer))
    return refusal('invalidRequest', 'pptx must be a Uint8Array or an ArrayBuffer.');
  const known = bytes instanceof Uint8Array ? steps.get(bytes) : undefined;
  if (known) return known.pkg;
  const read = readPptx(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
  // The sanitizers of an open: a structure mark in a file is forged (`openStore`).
  return read.ok
    ? withoutEditorMarks(read.pkg)
    : refusal('invalidDocument', 'The bytes are not a valid PPTX file.');
}

/**
 * Keep the keys reaching the editor after a replacement. The hidden text input has the focus
 * only for a text selection: when the replacement removed the text, the slide takes the focus,
 * so a key does not go to an input that writes nothing. @internal
 */
export function keepKeyboardFocus(editor: PptxEditorInstance): void {
  const input = surfaceInputOf(editor);
  if (!input || input.ownerDocument.activeElement !== input) return;
  if (editor.snapshot().selection.kind !== 'text') focusSurfaceOf(editor);
}

/**
 * Whether going from `base` to `next` changes embedded fonts, which a session registers only
 * when it opens: a font part, or the font list of the presentation (`p:embeddedFontLst`).
 */
const touchesFonts = (ops: readonly PptxOp[], base: PptxPackage, next: PptxPackage): boolean =>
  ops.some((op) => /^\/ppt\/fonts\//i.test(op.part)) ||
  JSON.stringify(presentationInfo(base).embeddedFonts) !==
    JSON.stringify(presentationInfo(next).embeddedFonts);

interface Prepared {
  readonly session: DocumentSession;
  readonly ops: readonly PptxOp[];
  /** The difference changes embedded fonts. */
  readonly fonts: boolean;
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
  const known = request.pptx instanceof Uint8Array ? steps.get(request.pptx)?.diff : undefined;
  const diff =
    known && known.base === session.store.package
      ? known.result
      : diffPackages(session.store.package, next);
  if (!diff.ok)
    return diff.reason === 'different-main-part'
      ? refusal('editUnsupported', 'The file names a different main part.', 'differentMainPart')
      : refusal('editUnsupported', 'The difference is not an edit.', 'unsupportedDifference');
  const fonts = diff.ops.length > 0 && touchesFonts(diff.ops, session.store.package, next);
  return { session, ops: diff.ops, fonts };
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
  // A session registers embedded faces only when it opens: a reload opens such a file whole.
  if (prepared.fonts)
    return refusal('editUnsupported', 'The file changes embedded fonts.', 'unsupportedDifference');
  const outcome = withChangeSource(editor, 'refresh', () => commit(editor, prepared, true));
  if (outcome.ok) keepKeyboardFocus(editor);
  return outcome;
}

/**
 * Show a returned file as a reload does: the content is the file's, and the history is reset.
 * The difference applies in place, so unchanged slides and thumbnails do not paint again and no
 * loading screen shows. The editor follows the slide on screen by its part, also to a new
 * index. The selection stays when the nodes it names stay, with a text selection
 * clamped into its text, as after an edit; else it is none, and the slide takes the keyboard
 * focus from the hidden text input. Returns `'load'` when the editor must open the file with
 * `load()` instead: embedded fonts changed, the difference is no edit, or the part checks refuse
 * it as an edit.
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
  if (prepared.fonts) return 'load';
  const outcome = withChangeSource(editor, source, (): ReplaceDocumentOutcome | 'load' => {
    const { session } = prepared;
    // A reload is a new document. The new load id applies before the commit, so the change
    // event of the commit names it, and the next edit is no load (`adoptLoadId`).
    const before = loadIdOf(session);
    const next = `load-${++nextLoad}`;
    loadIds.set(session, next);
    const result = commit(editor, prepared, false);
    if (!result.ok) {
      loadIds.set(session, before);
      // The part checks refuse the difference as an edit: a load opens the file whole. A
      // collaboration room never gets here: a reload refuses before it.
      return isCollaborating(editor) ? result : 'load';
    }
    adoptLoadId(editor, next);
    // The editor's store listener followed the slide on screen by its part, and kept its
    // selection when its nodes stay. A removed slide leaves the index, clamped.
    // Handles taken before the reload name another load: no undo step reaches them.
    session.store.clearHistory();
    session.store.closeHistoryGroup();
    // Publish the empty history and the new load id.
    editor.setMode(editor.snapshot().mode);
    return result;
  });
  if (outcome !== 'load' && outcome.ok) keepKeyboardFocus(editor);
  return outcome;
}

/**
 * The index that the slide on screen has in the returned file `bytes` (read with
 * `preparseReturned`): a reload with `load()` opens there, so the editor follows the slide when
 * slides before it come or go. The index on screen when the file names no such slide part.
 *
 * @internal
 */
export function followedSlide(editor: PptxEditorInstance, bytes: Uint8Array): number {
  const snap = editor.snapshot();
  const session = editor.session;
  const pkg = steps.get(bytes)?.pkg;
  if (!session || !pkg || snap.presentationView === 'slideMaster') return snap.currentSlide;
  const part = slideRefAt(session.store.package, snap.currentSlide)?.partName;
  const at = part ? presentationInfo(pkg).slides.findIndex((s) => s.partName === part) : -1;
  return at >= 0 ? at : snap.currentSlide;
}

/**
 * Add `getDocumentHandle()` and the `replaceDocument` command to an editor, and the teardowns
 * of `onEditorDestroy` to its `destroy()`. The facade calls it
 * once, on the instance it returns.
 *
 * @internal
 */
export function withDocumentReplace(
  editor: PptxEditorInstance,
  translate?: Translate
): PptxEditorInstance {
  if (translate) translations.set(editor, translate);
  const destroy = editor.destroy;
  editor.destroy = () => {
    const set = teardowns.get(editor);
    teardowns.delete(editor);
    for (const teardown of set ?? []) {
      try {
        teardown();
      } catch {
        /* A teardown cannot stop the destroy. */
      }
    }
    destroy.call(editor);
  };
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
