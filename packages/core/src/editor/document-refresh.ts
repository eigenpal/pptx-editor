// The document refresh controller, after DOCX Editor's
// packages/core/src/editor/document-refresh.ts at b63cbea4a: the same lifecycle, sequence rules,
// refusal codes, and recovery.
//
// PPTX Editor shows an accepted result without a reset of the view. With `mode: 'edit'` it runs
// `replaceDocument`, one undoable transaction. A reload applies the same difference in place,
// with no history, and then clears the history, as a load does (`reloadInPlace`). In both
// modes the selection stays when the nodes it names stay. In both modes parts the result did not change keep their identity, so the
// slide on screen and the thumbnails paint again only where the result changed them, the old
// paint stays until the new one replaces it, and the loading screen never shows. A result that
// changes embedded fonts, or that the part checks refuse as an edit, reloads with a quiet
// `load()`, which keeps the open deck on screen until the new one is ready.
//
// A result waits for input that runs when it arrives: a pointer gesture on the slide and an IME
// composition (`refresh-input.ts`). It then commits a notes draft. Input that changed the
// document makes the result refuse with `local-edits`, so a result never overwrites typing.
//
// An update runs in steps, each in a task of its own, so the main thread can paint between
// them: read the file, take the difference, commit and paint, then compute the changes. Each
// step after a yield checks again that the document is the captured one and that no
// cancellation came. The recovery bytes are the last bytes known to equal the document (the
// captured file, then each accepted file), so no save runs during an update. They stay only
// while no document shows: a load from outside drops them.
//
// Without `changes`, a registered `reviewModule()` computes the changes from the captured
// presentation to the returned one; without the module, the result reports
// `changeInformation: 'unavailable'`, as in DOCX Editor.

import { DocumentRefreshError } from './document-refresh-types.ts';
import type {
  ComputedRefreshChange,
  DocumentRefresh,
  DocumentRefreshState,
  RefreshFailureCode,
  RefreshChangeInput,
  RefreshResult,
  RefreshSubmission,
  RefreshUpdate,
} from './document-refresh-types.ts';
import type { DocumentHandle, ReplaceDocumentRefusal } from '../contracts/index.ts';
import type { PptxPackage } from '../store/index.ts';
import { findItem } from '../automation/commands/placement.ts';
import type { PptxEditorInstance } from './pptx-editor-types.ts';
import { isCollaborating, moduleRegistryOf } from './modules.ts';
import {
  followedSlide,
  forgetReturned,
  keepKeyboardFocus,
  onEditorDestroy,
  preparseReturned,
  prediffReturned,
  reloadInPlace,
  replacedWholeParts,
  whileReplacing,
} from './document-replace.ts';
import { expectLoad, forgetLoad } from './change-source.ts';
import {
  isComputedShape,
  isLocationShape,
  ownComputed,
  ownLocation,
  resolveRefreshChanges,
  type LocatedChange,
} from './document-refresh-changes.ts';
import { RefreshHighlights } from './refresh-highlights.ts';
import {
  REFRESH_INPUT_LIMITS,
  commitPendingInput,
  composing,
  gestureActive,
  trackRefreshInput,
  waitUntil,
  wakeRefreshWaits,
} from './refresh-input.ts';
import { revealRailItem } from './refresh-navigation.ts';

export { DocumentRefreshError } from './document-refresh-types.ts';
export type * from './document-refresh-types.ts';
export { shapeParagraphTexts, shapesById } from './document-refresh-changes.ts';
export { describeRefreshChange } from './refresh-describe.ts';

const controllers = new WeakMap<PptxEditorInstance, DocumentRefresh>();

/** Wait for the next task, so the browser can paint and handle input between steps. */
const nextTask = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 0));

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
              c.unavailableReason === 'unavailable') &&
            ((c as Partial<ComputedRefreshChange>).computed === undefined ||
              isComputedShape((c as Partial<ComputedRefreshChange>).computed))
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
          changes: update.changes.map((c) => {
            // A change from `comparePresentations()` keeps what it computed.
            const computed = (c as Partial<ComputedRefreshChange>).computed;
            return {
              id: c.id,
              ...(c.description !== undefined ? { description: c.description } : {}),
              ...(c.unavailableReason ? { unavailableReason: c.unavailableReason } : {}),
              ...(c.location ? { location: ownLocation(c.location) } : {}),
              ...(computed ? { computed: ownComputed(computed) } : {}),
            };
          }),
        }
      : {}),
    ...(update.failures ? { failures: update.failures.slice() } : {}),
  };
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
    /** A file equal to the document while its handle is `handle`: the recovery bytes. */
    current: Uint8Array;
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
    // A load from outside shows another document: the recovery bytes restore nothing now.
    if (!busy && recovery && editor.session && handle().loadId !== recovery.loadId) {
      recovery = null;
      notify({ recoveryAvailable: false });
    }
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
  /**
   * Wait for a running composition, at most `REFRESH_INPUT_LIMITS.compositionMs`, commit a notes
   * draft, then serialize with no await before the handle read.
   */
  const settle = async (current: number) => {
    const ended = await waitUntil(
      editor,
      () => current === generation && composing(editor),
      REFRESH_INPUT_LIMITS.compositionMs
    );
    if (current !== generation) throw new Error('superseded');
    if (!ended) throw new Error('composition');
    commitPendingInput(editor);
    const bytes = editor.save();
    if (!bytes) throw new Error('save');
    return { bytes, handle: handle() };
  };
  /**
   * Record an accepted result: advance the sequence, resolve its changes, and report. `pkg` and
   * `at` are the package and the handle right after the commit.
   */
  const accept = (
    request: NonNullable<typeof active>,
    resultId: string,
    update: RefreshUpdate,
    mode: 'reload' | 'edit',
    pkg: PptxPackage,
    at: DocumentHandle
  ): RefreshResult => {
    recovery = null;
    request.handle = at;
    request.current = update.bytes as Uint8Array;
    request.sequence = update.sequence;
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
    // A flash that showed moves to the new places of the same shapes and slides.
    const available = located.filter((entry) => entry.change.status === 'available');
    const highlightsVisible = highlights.adopt(available.flatMap((entry) => entry.targets));
    acceptedHandle = at;
    const changes = Object.freeze(located.map((entry) => entry.change));
    const result: RefreshResult = Object.freeze({
      ok: true,
      resultId,
      mode,
      changes,
      changeInformation: inputs ? 'available' : 'unavailable',
      failures: Object.freeze(update.failures ?? []),
    });
    notify({ changes, result, phase: 'complete', recoveryAvailable: false, highlightsVisible });
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
    // A load parses every part again: no node id of a held flash is sure to name its shape.
    highlights.forgetNodes(null);
    expectLoad(editor, source);
    const startSlide = followedSlide(editor, bytes);
    const loaded = await editor.load(bytes, { startSlide, quiet: true });
    // A load that failed or that another load replaced names no next load.
    forgetLoad(editor);
    if (loaded.ok) keepKeyboardFocus(editor);
    return loaded.ok ? 'ok' : 'failed';
  };
  /**
   * Let input that runs when a result arrives end, and commit a notes draft, so the revision
   * check sees it. A pointer gesture on the slide ends at its release; a lost release ends the
   * wait at its limit. False when a composition still runs at its limit.
   */
  const settleInput = async (): Promise<'ended' | 'composing' | 'cancelled'> => {
    const current = generation;
    const live = () => current === generation;
    await waitUntil(editor, () => live() && gestureActive(editor), REFRESH_INPUT_LIMITS.gestureMs);
    if (!live()) return 'cancelled';
    const ended = await waitUntil(
      editor,
      () => live() && composing(editor),
      REFRESH_INPUT_LIMITS.compositionMs
    );
    if (!live()) return 'cancelled';
    if (!ended) return 'composing';
    commitPendingInput(editor);
    return 'ended';
  };
  /**
   * The result of a reload whose load failed. A quiet load that fails keeps the open deck on
   * screen: it is restored already. A load from outside that replaced it shows another
   * document. With no document on screen, the recovery bytes open again.
   */
  const failedLoad = async (resultId: string, before: DocumentHandle): Promise<RefreshResult> => {
    active = null;
    const outside = () => !!editor.session && handle().loadId !== before.loadId;
    if (outside()) {
      recovery = null;
      return failure(resultId, 'document-changed');
    }
    let restored = !!editor.session;
    if (!restored && recovery) {
      const saved = recovery;
      notify({ phase: 'recovering' });
      restored = (await whileReplacing(editor, () => reload(saved.bytes, 'recovery'))) === 'ok';
      if (!restored && outside()) {
        recovery = null;
        return failure(resultId, 'document-changed');
      }
    }
    if (restored) recovery = null;
    return failure(resultId, restored ? 'load-failed' : 'recovery-failed', restored);
  };
  const controller: DocumentRefresh = {
    async capture() {
      // Capture serves both modes, so only `applyUpdate()` refuses a reload in collaboration.
      const code = refusal('edit');
      // A running replacement comes first: during a reload's quiet load, recovery bytes exist.
      if (busy) throw new DocumentRefreshError('busy');
      if (recovery) throw new DocumentRefreshError('recovery-failed');
      if (code) throw new DocumentRefreshError(code);
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
        active = {
          token,
          handle: captured.handle,
          captured: pkg,
          current: captured.bytes,
          sequence: -1,
          seen: new Map(),
        };
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
        // Input that runs now ends first: the checks below then see what it changed.
        if (active && active.token === mine.submission) {
          const input = await settleInput();
          if (input === 'cancelled') return publish(failure(resultId, 'cancelled'));
          if (input === 'composing') return publish(failure(resultId, 'local-edits'), 'failed');
        }
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
          // Step 1: read the file. An edit leaves a file that does not open to
          // `replaceDocument`, which reports it after its own checks.
          const parsed = preparseReturned(bytes);
          if (!parsed && mine.mode !== 'edit')
            return publish(failure(resultId, 'invalid-document'), 'failed');
          if (parsed) await nextTask();
          // Observers and other tasks can cancel, switch documents, or edit between steps.
          const unchanged = (): RefreshResult | null =>
            !valid()
              ? publish(failure(resultId, 'cancelled'))
              : handle().loadId !== now.loadId
                ? publish(failure(resultId, 'document-changed'))
                : !same(now, handle())
                  ? publish(failure(resultId, 'local-edits'), 'failed')
                  : null;
          let stop = unchanged();
          if (stop) return stop;
          // A flash that shows continues through the commit, at places that still exist.
          highlights.hold();
          if (mine.mode !== 'edit') {
            located = [];
            acceptedHandle = null;
            notify({ phase: 'refreshing', changes: [], highlightsVisible: highlights.visible });
            if ((stop = unchanged())) return stop;
          }
          // Step 2: the difference from the open package.
          if (parsed) {
            prediffReturned(editor, bytes);
            await nextTask();
            if ((stop = unchanged())) return stop;
          }
          // A part replaced whole gets new node ids, which can name other shapes: when the
          // commit changes the document, its canvas targets go until the result names the new
          // places.
          highlights.forgetNodes(replacedWholeParts(editor, bytes));
          // Step 3: commit and paint.
          let done: { pkg: PptxPackage; at: DocumentHandle };
          if (mine.mode === 'edit') {
            const result = editor.exec({ type: 'replaceDocument', pptx: bytes, expected: now });
            const outcome = result.replaceDocument;
            if (!outcome) return publish(failure(resultId, 'unavailable'), 'failed');
            if (!outcome.ok) return publish(editFailure(resultId, outcome), 'failed');
            done = { pkg: editor.session!.store.package, at: handle() };
          } else {
            recovery = { bytes: request.current.slice(), loadId: now.loadId };
            const outcome = await whileReplacing(editor, () => reload(bytes, 'refresh'));
            if (outcome === 'invalid') {
              recovery = null;
              return publish(failure(resultId, 'invalid-document'), 'failed');
            }
            if (outcome === 'failed') return publish(await failedLoad(resultId, now), 'failed');
            // The result shows: nothing is left to recover, also if a cancel comes next.
            recovery = null;
            done = { pkg: editor.session!.store.package, at: handle() };
          }
          // Step 4: the changes, after the new paint had a frame.
          await nextTask();
          if (!valid()) {
            // Accepted content stays accepted if cancellation races with completion.
            recovery = null;
            return publish(failure(resultId, 'cancelled'));
          }
          return publish(accept(request, resultId, mine, mine.mode ?? 'reload', done.pkg, done.at));
        } catch {
          return publish(
            failure(resultId, valid() ? 'input-failed' : 'cancelled'),
            valid() ? 'failed' : undefined
          );
        } finally {
          busy = false;
          forgetReturned(bytes);
          // No accepted result: a flash of the unchanged document shows on, else it goes.
          if (highlights.holding) {
            if (acceptedHandle && same(acceptedHandle, handle())) highlights.release();
            else {
              highlights.clear();
              if (state.highlightsVisible) notify({ highlightsVisible: false });
            }
          }
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
      // As in DOCX Editor: the recovery bytes stay only while no document shows.
      const keepRecovery = state.phase === 'recovering' || (!editor.session && recovery !== null);
      generation++;
      active = null;
      // A wait for a gesture or a composition ends now, and its apply returns `cancelled`.
      wakeRefreshWaits(editor);
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
      // The bytes restore the document they were taken from, or open where no document shows.
      if (!recovery || busy || (editor.session && recovery.loadId !== handle().loadId))
        return false;
      busy = true;
      const saved = recovery;
      notify({ phase: 'recovering' });
      try {
        const outcome = await whileReplacing(editor, () => reload(saved.bytes, 'recovery'));
        const ok = outcome === 'ok';
        if (ok || (editor.session && handle().loadId !== saved.loadId)) recovery = null;
        notify({ phase: ok ? 'idle' : 'failed', recoveryAvailable: recovery !== null });
        return ok;
      } finally {
        busy = false;
      }
    },
  };
  trackRefreshInput(editor);
  // A destroyed editor keeps no observer, subscription, or timer of the change flash.
  onEditorDestroy(editor, () => {
    reset();
    highlights.dispose();
    recovery = null;
    controllers.delete(editor);
  });
  controllers.set(editor, controller);
  return controller;
}
