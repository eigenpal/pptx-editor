// The document refresh controller, after DOCX Editor's
// packages/core/src/editor/document-refresh.ts at b63cbea4a: the same lifecycle, sequence rules,
// refusal codes, and recovery.
//
// PPTX Editor shows an accepted result without a reset of the view. With `mode: 'edit'` it runs
// `replaceDocument`, one undoable transaction. A reload applies the same difference in place,
// with no history, and then clears the history and the selection, as a load does
// (`reloadInPlace`). In both modes parts the result did not change keep their identity, so the
// slide on screen and the thumbnails paint again only where the result changed them, the old
// paint stays until the new one replaces it, and the loading screen never shows. A result that
// changes embedded fonts reloads with a quiet `load()`, which keeps the open deck on screen
// until the new one is ready.
//
// Without `changes`, a registered `reviewModule()` computes the changes from the captured
// presentation to the returned one; without the module, the result reports
// `changeInformation: 'unavailable'`, as in DOCX Editor.

import { DocumentRefreshError } from './document-refresh-types.ts';
import type {
  DocumentRefresh,
  DocumentRefreshState,
  RefreshFailureCode,
  RefreshChangeInput,
  RefreshResult,
  RefreshSubmission,
  RefreshUpdate,
} from './document-refresh-types.ts';
import type { DocumentHandle, ReplaceDocumentRefusal } from '../contracts/index.ts';
import { readPptx, type PptxPackage } from '../store/index.ts';
import { findItem } from '../automation/commands/placement.ts';
import type { PptxEditorInstance } from './pptx-editor-types.ts';
import { isCollaborating, moduleRegistryOf } from './modules.ts';
import { reloadInPlace, whileReplacing } from './document-replace.ts';
import { expectLoad } from './change-source.ts';
import {
  isLocationShape,
  ownLocation,
  resolveRefreshChanges,
  type LocatedChange,
} from './document-refresh-changes.ts';
import { RefreshHighlights } from './refresh-highlights.ts';
import { editorHostDocument, surfaceRootOf } from './surface-registry.ts';
import { revealRailItem } from './refresh-navigation.ts';

export { DocumentRefreshError } from './document-refresh-types.ts';
export type * from './document-refresh-types.ts';
export { shapeParagraphTexts, shapesById } from './document-refresh-changes.ts';
export { describeRefreshChange } from './refresh-describe.ts';

const controllers = new WeakMap<PptxEditorInstance, DocumentRefresh>();

/** The `message` a refusal carries when nothing more specific explains it. */
const REFRESH_FAILURE_MESSAGES: Readonly<Record<RefreshFailureCode, string>> = {
  unavailable: 'No document is mounted, or it is still opening.',
  collaboration: 'A reload cannot replace the document while a collaboration session is attached.',
  busy: 'A replacement is running.',
  cancelled: 'Cancellation interrupted the update.',
  superseded: 'The result belongs to an inactive submission, or arrived after cancellation.',
  'document-changed': 'The document changed during the update.',
  'local-edits': 'The document changed after capture.',
  'out-of-order': 'A duplicate or older sequence arrived.',
  'invalid-result': 'The update object is malformed.',
  'invalid-document': 'The returned bytes are not a valid PPTX file.',
  'input-failed': 'Pending input or serialization failed.',
  'load-failed': 'A valid file failed to open.',
  'recovery-failed': 'Opening and restoration both failed.',
  'edit-not-permitted': 'The editor does not permit an undoable edit.',
  'edit-unsupported': 'The result cannot apply as an undoable edit.',
};
let nextController = 0;

/** Whether `update` has the documented shape. */
function wellFormed(update: RefreshUpdate): boolean {
  const changes = update.changes;
  return (
    Number.isSafeInteger(update.sequence) &&
    update.sequence >= 0 &&
    (update.bytes instanceof Uint8Array || update.bytes instanceof ArrayBuffer) &&
    (update.mode === undefined || update.mode === 'reload' || update.mode === 'edit') &&
    (update.failures === undefined ||
      (Array.isArray(update.failures) &&
        update.failures.length <= 10000 &&
        update.failures.every((id) => typeof id === 'string'))) &&
    (changes === undefined ||
      (Array.isArray(changes) &&
        changes.length <= 10000 &&
        new Set(changes.map((c) => c?.id)).size === changes.length &&
        changes.every(
          (c) =>
            !!c &&
            typeof c.id === 'string' &&
            c.id !== '' &&
            (c.description === undefined ||
              (typeof c.description === 'string' && c.description.length <= 1000)) &&
            (c.location === undefined || isLocationShape(c.location)) &&
            (c.unavailableReason === undefined ||
              c.unavailableReason === 'deleted' ||
              c.unavailableReason === 'unavailable')
        )))
  );
}

/** A copy of the update the caller cannot change while it waits in the queue. */
function owned(update: RefreshUpdate): RefreshUpdate {
  const bytes = update.bytes instanceof Uint8Array ? update.bytes : new Uint8Array(update.bytes);
  return {
    ...update,
    bytes: bytes.slice(),
    ...(update.changes
      ? {
          changes: update.changes.map((c) => ({
            id: c.id,
            ...(c.description !== undefined ? { description: c.description } : {}),
            ...(c.unavailableReason ? { unavailableReason: c.unavailableReason } : {}),
            ...(c.location ? { location: ownLocation(c.location) } : {}),
          })),
        }
      : {}),
    ...(update.failures ? { failures: update.failures.slice() } : {}),
  };
}

/**
 * The editor's text input while an IME composition runs in it, or null. Composition events
 * bubble, so one pair of listeners on the editor's document sees every input it mounts.
 */
function compositionTracker(editor: PptxEditorInstance): () => HTMLElement | null {
  const doc = editorHostDocument(editor);
  if (!doc) return () => null;
  let input: HTMLElement | null = null;
  const mine = (e: Event): HTMLElement | null => {
    const target = e.target as HTMLElement | null;
    const root = surfaceRootOf(editor);
    return target?.classList?.contains('pptx-surface__input') && root?.contains(target)
      ? target
      : null;
  };
  doc.addEventListener('compositionstart', (e) => void (input = mine(e) ?? input), true);
  doc.addEventListener('compositionend', (e) => void (mine(e) === input && (input = null)), true);
  return () => (input?.isConnected ? input : null);
}

/**
 * Get the engine-owned controller for updated PPTX files. React and Vue use the same controller.
 * Transport and processor jobs belong to your application. This is a host API, not an Office.js
 * presentation member.
 *
 * @example
 * ```ts
 * const refresh = createDocumentRefresh(editor);
 * const submission = await refresh.capture();
 * try {
 *   const bytes = await process(submission.bytes);
 *   const result = await refresh.applyUpdate({ submission, sequence: 1, bytes });
 *   if (result.ok) refresh.highlightChanges();
 * } finally {
 *   refresh.finish(submission);
 * }
 * ```
 *
 * @public
 */
export function createDocumentRefresh(editor: PptxEditorInstance): DocumentRefresh {
  const existing = controllers.get(editor);
  if (existing) return existing;
  const prefix = `refresh-${++nextController}`;
  let generation = 0;
  let active: {
    token: RefreshSubmission;
    /** The document an update must still find: taken at capture, moved on accept. */
    handle: DocumentHandle;
    /** The captured presentation, the base of computed changes. */
    captured: PptxPackage | null;
    sequence: number;
    seen: Map<string, string>;
  } | null = null;
  let busy = false;
  let queue = Promise.resolve();
  let recovery: { bytes: Uint8Array; loadId: string } | null = null;
  let located: LocatedChange[] = [];
  /** The handle the located changes name. Any other handle makes them stale. */
  let acceptedHandle: DocumentHandle | null = null;
  const listeners = new Set<() => void>();
  const resultListeners = new Set<(result: RefreshResult) => void>();
  let state: DocumentRefreshState = Object.freeze({
    phase: 'idle',
    result: null,
    changes: [],
    highlightsVisible: false,
    recoveryAvailable: false,
  });
  const notify = (patch: Partial<DocumentRefreshState>) => {
    state = Object.freeze({ ...state, ...patch });
    for (const listener of [...listeners]) {
      try {
        listener();
      } catch {
        /* A host observer cannot abort replacement. */
      }
    }
  };
  const composing = compositionTracker(editor);
  const highlights = new RefreshHighlights(editor, () => {
    if (state.highlightsVisible) notify({ highlightsVisible: false });
  });
  const handle = () => editor.getDocumentHandle();
  const same = (a: DocumentHandle | null, b: DocumentHandle) =>
    !!a && a.loadId === b.loadId && a.revision === b.revision;
  const publish = (result: RefreshResult, phase?: DocumentRefreshState['phase']): RefreshResult => {
    if (phase) notify({ phase, result, recoveryAvailable: recovery !== null });
    for (const listener of [...resultListeners]) {
      try {
        listener(result);
      } catch {
        /* Isolate host observers. */
      }
    }
    return result;
  };
  const failure = (
    resultId: string,
    code: Exclude<RefreshFailureCode, 'edit-not-permitted' | 'edit-unsupported'>,
    recovered?: boolean
  ): RefreshResult => ({
    ok: false,
    resultId,
    code,
    message: REFRESH_FAILURE_MESSAGES[code],
    ...(recovered === undefined ? {} : { recovered }),
  });
  /** A `replaceDocument` refusal in this controller's own codes, keeping its cause. */
  const editFailure = (resultId: string, refusal: ReplaceDocumentRefusal): RefreshResult => {
    const base = { ok: false as const, resultId, message: refusal.message };
    switch (refusal.code) {
      case 'editNotPermitted':
        return { ...base, code: 'edit-not-permitted', reason: 'viewing' };
      case 'editUnsupported':
        return {
          ...base,
          code: 'edit-unsupported',
          reason:
            refusal.reason === 'differentMainPart'
              ? 'different-main-part'
              : refusal.reason === 'collaborationRefused'
                ? 'collaboration-refused'
                : 'unsupported-difference',
        };
      case 'unavailable':
        return refusal.reason
          ? { ...base, code: 'unavailable', reason: 'collaboration-unavailable' }
          : { ...base, code: 'unavailable' };
      case 'staleDocument':
        return {
          ...base,
          code: refusal.reason === 'otherDocument' ? 'document-changed' : 'local-edits',
        };
      case 'invalidRequest':
        return { ...base, code: 'invalid-result' };
      case 'invalidDocument':
        return { ...base, code: 'invalid-document' };
      case 'busy':
        return { ...base, code: 'busy' };
    }
  };
  // A reload replaces the document for this editor only, which a collaboration room cannot
  // follow; an edit can.
  const refusal = (mode: RefreshUpdate['mode']): 'collaboration' | 'unavailable' | null =>
    isCollaborating(editor) && mode !== 'edit'
      ? 'collaboration'
      : !editor.session || editor.snapshot().isOpening
        ? 'unavailable'
        : null;
  const reset = () => {
    generation++;
    active = null;
    located = [];
    acceptedHandle = null;
    highlights.clear();
  };
  // Locations name one accepted document. Never paint stale places after another change.
  editor.subscribe(() => {
    if (busy || acceptedHandle === null || same(acceptedHandle, handle())) return;
    located = [];
    acceptedHandle = null;
    highlights.clear();
    if (state.changes.length > 0 || state.highlightsVisible)
      notify({ changes: [], highlightsVisible: false });
  });
  const rendered = (index: number, nodeId: string): boolean => {
    const layout = editor.session?.layout(index);
    return !!layout && !!findItem(layout.items, nodeId);
  };
  /** Wait for a running composition, then serialize with no await before the handle read. */
  const settle = async (current: number) => {
    const input = composing();
    if (input)
      await new Promise<void>((resolve) => {
        const done = () => (input.removeEventListener('compositionend', done), resolve());
        input.addEventListener('compositionend', done);
      });
    if (current !== generation) throw new Error('superseded');
    const bytes = editor.save();
    if (!bytes) throw new Error('save');
    return { bytes, handle: handle() };
  };
  /** Record an accepted result: advance the sequence, resolve its changes, and report. */
  const accept = (
    request: NonNullable<typeof active>,
    resultId: string,
    update: RefreshUpdate,
    mode: 'reload' | 'edit'
  ): RefreshResult => {
    highlights.clear();
    recovery = null;
    request.handle = handle();
    request.sequence = update.sequence;
    const pkg = editor.session!.store.package;
    const review = moduleRegistryOf(editor).review;
    let inputs: readonly RefreshChangeInput[] | null = update.changes ?? null;
    if (!inputs && review?.refreshChanges && request.captured) {
      try {
        inputs = review.refreshChanges(request.captured, pkg);
      } catch {
        inputs = null;
      }
    }
    located = inputs ? resolveRefreshChanges(pkg, resultId, inputs, request.seen, rendered) : [];
    for (const entry of located) request.seen.set(entry.change.id, entry.fingerprint);
    acceptedHandle = request.handle;
    const changes = Object.freeze(located.map((entry) => entry.change));
    const result: RefreshResult = Object.freeze({
      ok: true,
      resultId,
      mode,
      changes,
      changeInformation: inputs ? 'available' : 'unavailable',
      failures: Object.freeze(update.failures ?? []),
    });
    notify({ changes, result, phase: 'complete', recoveryAvailable: false });
    return result;
  };
  /** Show `bytes` as a reload: in place when it can, else with a quiet load. */
  const reload = async (
    bytes: Uint8Array,
    source: 'refresh' | 'recovery'
  ): Promise<'ok' | 'invalid' | 'failed'> => {
    const inPlace = reloadInPlace(editor, bytes, source);
    if (inPlace !== 'load')
      return inPlace.ok ? 'ok' : inPlace.code === 'invalidDocument' ? 'invalid' : 'failed';
    expectLoad(editor, source);
    const startSlide = editor.snapshot().currentSlide;
    const loaded = await editor.load(bytes, { startSlide, quiet: true });
    return loaded.ok ? 'ok' : 'failed';
  };
  const controller: DocumentRefresh = {
    async capture() {
      // Capture serves both modes, so only `applyUpdate()` refuses a reload in collaboration.
      const code = refusal('edit');
      if (recovery) throw new DocumentRefreshError('recovery-failed');
      if (code || busy) throw new DocumentRefreshError(code ?? 'busy');
      reset();
      recovery = null;
      const current = generation;
      notify({
        phase: 'capturing',
        result: null,
        changes: [],
        highlightsVisible: false,
        recoveryAvailable: false,
      });
      try {
        const captured = await settle(current);
        if (current !== generation) throw new Error('superseded');
        const token = Object.freeze({
          id: `${prefix}-${current}`,
          bytes: captured.bytes.slice().buffer as ArrayBuffer,
        });
        const pkg = editor.session?.store.package ?? null;
        active = { token, handle: captured.handle, captured: pkg, sequence: -1, seen: new Map() };
        notify({ phase: 'processing' });
        return token;
      } catch (error) {
        if (current === generation) notify({ phase: 'failed' });
        throw new DocumentRefreshError(
          current === generation ? 'input-failed' : 'superseded',
          error
        );
      }
    },
    applyUpdate(update) {
      const resultId = `${update?.submission?.id ?? 'unknown'}:${update?.sequence}`;
      let mine: RefreshUpdate;
      try {
        if (!update || !wellFormed(update)) throw new Error('invalid-result');
        mine = owned(update);
      } catch {
        return Promise.resolve(publish(failure(resultId, 'invalid-result')));
      }
      const apply = async (): Promise<RefreshResult> => {
        const request = active;
        const code = refusal(mine.mode);
        if (code) return publish(failure(resultId, code));
        if (!request || request.token !== mine.submission)
          return publish(failure(resultId, 'superseded'));
        if (mine.sequence <= request.sequence) return publish(failure(resultId, 'out-of-order'));
        const now = handle();
        if (now.loadId !== request.handle.loadId)
          return publish(failure(resultId, 'document-changed'));
        if (now.revision !== request.handle.revision)
          return publish(failure(resultId, 'local-edits'), 'failed');
        const valid = () => active === request;
        const bytes = mine.bytes as Uint8Array;
        busy = true;
        try {
          if (mine.mode === 'edit') {
            const result = editor.exec({ type: 'replaceDocument', pptx: bytes, expected: now });
            const outcome = result.replaceDocument;
            if (!outcome) return publish(failure(resultId, 'unavailable'), 'failed');
            if (!outcome.ok) return publish(editFailure(resultId, outcome), 'failed');
            return publish(accept(request, resultId, mine, 'edit'));
          }
          if (!readPptx(bytes).ok) return publish(failure(resultId, 'invalid-document'), 'failed');
          highlights.clear();
          located = [];
          acceptedHandle = null;
          notify({ phase: 'refreshing', changes: [], highlightsVisible: false });
          // Observers can cancel, switch documents, or edit when notified.
          if (!valid()) return publish(failure(resultId, 'cancelled'));
          if (!same(now, handle())) return publish(failure(resultId, 'local-edits'), 'failed');
          const previous = editor.save();
          if (!previous) return publish(failure(resultId, 'input-failed'), 'failed');
          recovery = { bytes: previous.slice(), loadId: now.loadId };
          const outcome = await whileReplacing(editor, () => reload(bytes, 'refresh'));
          if (outcome === 'invalid') {
            recovery = null;
            return publish(failure(resultId, 'invalid-document'), 'failed');
          }
          if (outcome === 'failed') {
            // A quiet load that fails keeps the open deck on screen: it is restored already.
            const restored = handle().loadId === now.loadId;
            if (restored) recovery = null;
            active = null;
            return publish(
              failure(resultId, restored ? 'load-failed' : 'recovery-failed', restored),
              'failed'
            );
          }
          if (!valid()) {
            // Accepted content stays accepted if cancellation races with completion.
            recovery = null;
            return publish(failure(resultId, 'cancelled'));
          }
          return publish(accept(request, resultId, mine, 'reload'));
        } catch {
          return publish(
            failure(resultId, valid() ? 'input-failed' : 'cancelled'),
            valid() ? 'failed' : undefined
          );
        } finally {
          busy = false;
        }
      };
      const result = queue.then(apply);
      queue = result.then(
        () => undefined,
        () => undefined
      );
      return result;
    },
    cancel() {
      const keepRecovery = state.phase === 'recovering' || recovery !== null;
      generation++;
      active = null;
      if (!keepRecovery) recovery = null;
      notify({
        phase:
          keepRecovery && !state.result?.ok ? 'failed' : state.result?.ok ? 'complete' : 'idle',
        result: keepRecovery || state.result?.ok ? state.result : null,
        recoveryAvailable: recovery !== null,
      });
    },
    finish(submission) {
      if (active?.token === submission && !busy) {
        active = null;
        notify({ phase: state.phase === 'processing' ? 'idle' : state.phase });
      }
    },
    snapshot: () => state,
    subscribe(listener) {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    onResult(listener) {
      resultListeners.add(listener);
      return () => {
        resultListeners.delete(listener);
      };
    },
    highlightChanges(options) {
      const settings = highlights.validate(options);
      if (busy || !same(acceptedHandle, handle())) return 0;
      const ids = options?.changeIds === undefined ? null : new Set(options.changeIds);
      const selected = located.filter(
        (entry) =>
          entry.change.status === 'available' &&
          entry.targets.length > 0 &&
          (ids ? ids.has(entry.change.id) : options?.includePrevious || entry.change.isNew)
      );
      highlights.show(
        selected.flatMap((entry) => entry.targets),
        settings
      );
      notify({ highlightsVisible: selected.length > 0 });
      return selected.length;
    },
    clearHighlights(options) {
      highlights.hide(options);
      notify({ highlightsVisible: false });
    },
    navigateToChange(id, options = {}) {
      if (options.focus !== undefined && typeof options.focus !== 'boolean')
        throw new TypeError('focus must be a boolean.');
      const block = options.block ?? 'center';
      if (!['start', 'center', 'centerIfNeeded', 'nearest'].includes(block))
        throw new TypeError("block must be 'start', 'center', 'centerIfNeeded', or 'nearest'.");
      const behavior = options.behavior ?? 'instant';
      if (behavior !== 'instant' && behavior !== 'smooth')
        throw new TypeError("behavior must be 'instant' or 'smooth'.");
      const offsetPx = options.offsetPx ?? 24;
      if (typeof offsetPx !== 'number' || !Number.isFinite(offsetPx) || offsetPx < 0)
        throw new RangeError('offsetPx must be a nonnegative finite number.');
      if (busy || !same(acceptedHandle, handle())) return false;
      const target = located.find((e) => e.change.id === id && e.change.status === 'available')
        ?.targets[0];
      if (!target) return false;
      if (editor.snapshot().currentSlide !== target.slideIndex) editor.goToSlide(target.slideIndex);
      revealRailItem(editor, target.slideIndex, { block, behavior, offsetPx });
      if (options.focus) {
        if (target.nodeId && target.text && !target.notes) {
          const at = { paragraph: target.text.paragraph, offset: target.text.start };
          editor.setSelection({ kind: 'text', nodeId: target.nodeId, anchor: at, focus: at });
        } else if (target.nodeId && !target.notes)
          editor.setSelection({ kind: 'shapes', nodeIds: [target.nodeId] });
        editor.focus();
      }
      return true;
    },
    recoveryBytes: () => (recovery ? (recovery.bytes.slice().buffer as ArrayBuffer) : null),
    async recover() {
      if (!recovery || recovery.loadId !== handle().loadId || busy) return false;
      busy = true;
      const saved = recovery;
      notify({ phase: 'recovering' });
      try {
        const outcome = await whileReplacing(editor, () => reload(saved.bytes, 'recovery'));
        const ok = outcome === 'ok';
        if (ok) recovery = null;
        notify({ phase: ok ? 'idle' : 'failed', recoveryAvailable: !ok });
        return ok;
      } finally {
        busy = false;
      }
    },
  };
  controllers.set(editor, controller);
  return controller;
}
