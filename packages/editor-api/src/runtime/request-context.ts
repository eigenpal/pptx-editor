/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// After DOCX Editor's `packages/editor-api/src/runtime/request-context.ts` (commit 5993bd223945).
//
// One queue, one presentation, one sync at a time. A `sync()` called while another is in flight
// takes the calls queued so far at once, then waits for it: syncs reach the host in the order
// they were called, and a refused sync does not stop the ones after it. The presentation host takes result
// references, so the writes of a sync are ONE atomic host batch: an object a write creates can
// be read and written in the same sync. Reads go in earlier rounds when other calls wait for
// their answers, and `load('items/<name>')` on a collection reads the items in a later round,
// because the items are known only once the listing answers. Every later round is pinned to the
// revision the round before it left, so it reads the same presentation.

import type { AutomationHost, AutomationOperation } from '@pptx-editor.dev/core/automation';
import type { Application } from '../model/application.ts';
import type { Presentation } from '../model/presentation.ts';
import {
  awaitedReads,
  batchFailure,
  planBatch,
  readAfterWriteCut,
  settleBatch,
  withReplayedLookups,
} from './batch.ts';
import { markRefused, markReplayable } from './object-path.ts';
import { ClientObject } from './client-object.ts';
import {
  diagnosticOf,
  fail,
  failArgument,
  isPptxEditorError,
  nearest,
  PptxEditorError,
  withTraceMessages,
} from './errors.ts';
import {
  INTERNALS,
  LOADED_VALUES,
  PATH,
  REBIND,
  RELEASE,
  type ContextInternals,
} from './internals.ts';
import { resolveLoadOption, type LoadOption } from './load-options.ts';
import { ActionQueue, type QueuedAction } from './queue.ts';
import { TrackedObjects } from './tracked-objects.ts';
import { warnUnsynced } from './dev-warnings.ts';
import type { PresentationCapabilities } from './runtime.ts';

/**
 * A refusal of a read that went after the writes of its sync committed. Elsewhere a refusal
 * means nothing applied, so this one says that the writes stay.
 */
function readAfterCommit(error: unknown): unknown {
  if (!(error instanceof PptxEditorError)) return error;
  const diagnostic = diagnosticOf('read-after-commit');
  return new PptxEditorError({
    code: error.code,
    ...(error.target === undefined ? {} : { target: error.target }),
    ...(diagnostic === undefined ? {} : { diagnostic }),
  });
}

/**
 * A refusal in a later part of a split sync: the writes of the parts before it applied. The
 * diagnostic says so, so the caller does not apply them again.
 */
function partlyApplied(error: unknown): unknown {
  if (!(error instanceof PptxEditorError)) return error;
  const general = diagnosticOf('partly-applied');
  const own = error.debugInfo.innerError;
  // The cause keeps its own code; the hint adds that the writes before the read applied.
  const diagnostic = own
    ? { code: own.code, hint: `${own.message} ${general?.hint ?? ''}`.trim() }
    : general;
  return new PptxEditorError({
    code: error.code,
    ...(error.target === undefined ? {} : { target: error.target }),
    ...(diagnostic === undefined ? {} : { diagnostic }),
  });
}

/** What a context needs from the runtime that made it. */
export interface RuntimeSession {
  readonly host: AutomationHost;
  readonly capabilities: PresentationCapabilities;
  /** Told about each refused batch before it is thrown. A listener that throws is ignored. */
  readonly onRefusal?: (error: PptxEditorError, operations: readonly AutomationOperation[]) => void;
  /**
   * Told about each batch that threw an error other than a refusal, before it propagates: an
   * engine defect. A listener that throws is ignored.
   */
  readonly onException?: (error: unknown, operations: readonly AutomationOperation[]) => void;
  /** Identity for adoption checks. */
  readonly id: object;
  /** Refuse if the runtime has been disposed. */
  assertLive(target?: string): void;
  /** Build the presentation proxy for a context. Injected, so this module does not import it. */
  openPresentation(context: RequestContext): Presentation;
  /** Build the application proxy for a context. Injected for the same reason. */
  openApplication(context: RequestContext): Application;
  /** The Office.js type name of an object of the model, for `loadRecursive`. */
  typeName(object: ClientObject): string | undefined;
  /** Every type name `loadRecursive` takes. */
  readonly typeNames: readonly string[];
}

/**
 * Debug information about a request context, as Office.js `OfficeExtension.RequestContextDebugInfo`.
 *
 * @public
 */
export interface RequestContextDebugInfo {
  /**
   * What the next `sync()` sends, in order: the consumer-facing path of each queued call, such
   * as `presentation.slides.getItemAt(0).load`, and the message of each `trace` call. Never a
   * handle or presentation content.
   */
  pendingStatements: string[];
}

/**
 * The Office.js `OfficeExtension.ClientRequestContext`: here the same class as
 * {@link RequestContext}, under the name that code written for Office.js uses for the base of
 * every request context.
 *
 * @public
 */
export type ClientRequestContext = RequestContext;

/** The deepest `loadRecursive` goes. */
const MAX_LOAD_DEPTH = 32;

/** How many syncs every context has sent, for history groups that no two syncs share. */
let sentSyncs = 0;

/** A `trace` message, and how many calls were queued before it. */
interface Trace {
  readonly message: string;
  readonly at: number;
}

/** The messages of the traces queued before `failed`, or every message when it is unknown. */
function tracesBefore(
  actions: readonly QueuedAction[],
  traces: readonly Trace[],
  failed: QueuedAction | undefined
): string[] {
  const index = failed ? actions.indexOf(failed) : -1;
  return traces.filter((t) => index < 0 || t.at <= index).map((t) => t.message);
}

/**
 * What a `run` hands its callback: one queue, one presentation, one sync at a time.
 *
 * `sync()` sends every queued call as one atomic batch: all of it applies, or none of it. A
 * context that has READ remembers the revision it read at, and a later sync that writes is
 * conditional on it, failing `StaleDocument` if another writer moved the presentation. A context
 * that has read nothing writes unconditionally.
 *
 * @public
 */
export class RequestContext {
  readonly #session: RuntimeSession;
  readonly #queue = new ActionQueue();
  readonly #created = new Set<ClientObject>();
  readonly #tracked = new Set<ClientObject>();
  readonly #internals: ContextInternals;
  readonly #trackedObjects: TrackedObjects;
  #presentation: Presentation | undefined;
  #application: Application | undefined;
  /** `trace` calls since the last sync. */
  #traces: Trace[] = [];
  /** The action whose answer or settle failed last, for the trace messages of the error. */
  #failed: QueuedAction | undefined;
  /** The history group of the sync in progress (`AutomationBatchRequest.historyGroup`). */
  #historyGroup = '';
  #finished = false;
  #readRevision: number | null = null;
  /** Syncs called and not yet settled. */
  #syncs = 0;
  /** Settles when the last sync called settles; never rejects. */
  #tail: Promise<void> = Promise.resolve();
  #deferred: Promise<void>[] = [];

  private constructor(session: RuntimeSession) {
    this.#session = session;
    this.#internals = {
      host: session.host,
      capabilities: session.host.capabilities,
      queue: this.#queue,
      session: session.id,
      assertUsable: (target?: string) => {
        this.#session.assertLive(target);
        if (this.#finished)
          fail({ code: 'InvalidRequestContext', ...(target === undefined ? {} : { target }) });
      },
      isFinished: () => this.#finished,
      readRevision: () => this.#readRevision,
      register: (object) => void this.#created.add(object as ClientObject),
      track: (object) => void this.#tracked.add(object as ClientObject),
      untrack: (object) => void this.#tracked.delete(object as ClientObject),
      isTracked: (object) => this.#tracked.has(object as ClientObject),
      disown: (object) => {
        this.#created.delete(object as ClientObject);
        this.#tracked.delete(object as ClientObject);
      },
      defer: (work) => {
        // A later answer can fail before this work is awaited, and then the sync drops it:
        // mark it handled now, so its own failure is never an unhandled rejection.
        work.catch(() => {});
        this.#deferred.push(work);
      },
    };
    this.#trackedObjects = new TrackedObjects(this.#internals, (object) =>
      this.#created.has(object)
    );
  }

  /** The presentation this run is against: the same object for the life of the context. */
  get presentation(): Presentation {
    this.#internals.assertUsable('presentation');
    this.#presentation ??= this.#session.openPresentation(this);
    return this.#presentation;
  }

  /** The application the presentation is open in. It has no properties or methods. */
  get application(): Application {
    this.#internals.assertUsable('application');
    this.#application ??= this.#session.openApplication(this);
    return this.#application;
  }

  /** What the next `sync()` sends, for debugging. See {@link RequestContextDebugInfo}. */
  get debugInfo(): RequestContextDebugInfo {
    this.#internals.assertUsable('debugInfo');
    const labels = this.#queue.labels();
    const pendingStatements: string[] = [];
    let next = 0;
    for (let i = 0; i <= labels.length; i += 1) {
      while (next < this.#traces.length && this.#traces[next]!.at <= i)
        pendingStatements.push(this.#traces[next++]!.message);
      if (i < labels.length) pendingStatements.push(labels[i]!);
    }
    return { pendingStatements };
  }

  /** What the presentation host behind this context can do. */
  get capabilities(): PresentationCapabilities {
    return this.#session.capabilities;
  }

  /** Objects kept addressable past the run that created them. See {@link TrackedObjects}. */
  get trackedObjects(): TrackedObjects {
    return this.#trackedObjects;
  }

  /** Queue a load on `object`, as Office.js `context.load(object, option)`. */
  load<T extends ClientObject>(object: T, option?: LoadOption): T {
    return object.load(option);
  }

  /**
   * Queue loads on `object` and on the objects its loads reach, by type: each object whose
   * Office.js type name is a key of `options` loads that option, and the objects its load
   * fills in (navigation properties and collection items) go on, down to `maxDepth` levels
   * below `object` (at most 32, which is also the default). An object whose type is not a key
   * loads nothing and stops the walk there. Everything loads in the next `sync()`; each level
   * costs one more host batch.
   *
   * @example
   * ```ts
   * context.loadRecursive(context.presentation.slides, {
   *   SlideCollection: 'items',
   *   Slide: 'id,shapes',
   *   ShapeCollection: 'items',
   *   Shape: 'name',
   * });
   * await context.sync();
   * ```
   */
  loadRecursive(
    object: ClientObject,
    options: { [typeName: string]: LoadOption | string | string[] },
    maxDepth?: number
  ): void {
    const target = 'loadRecursive';
    this.#internals.assertUsable(target);
    if (!(object instanceof ClientObject)) failArgument(target, 'an object of this API');
    if (object.context !== this) fail({ code: 'InvalidObjectPath', target });
    if (typeof options !== 'object' || options === null || Array.isArray(options))
      failArgument(`${target}.options`, 'an object keyed by type name');
    const names = this.#session.typeNames;
    const table = new Map<string, LoadOption>();
    for (const [name, option] of Object.entries(options)) {
      if (!names.includes(name))
        failArgument(`${target}.options.${name}`, 'a type name of this API', nearest(name, names));
      // A bad option refuses now, before anything is queued.
      resolveLoadOption(option, `${target}.options.${name}`);
      table.set(name, option);
    }
    if (
      maxDepth !== undefined &&
      (typeof maxDepth !== 'number' ||
        !Number.isInteger(maxDepth) ||
        maxDepth < 0 ||
        maxDepth > MAX_LOAD_DEPTH)
    )
      failArgument(`${target}.maxDepth`, `a whole number from 0 to ${MAX_LOAD_DEPTH}`);
    const limit = maxDepth ?? MAX_LOAD_DEPTH;
    const seen = new WeakSet<ClientObject>();
    const visit = (current: ClientObject, depth: number): void => {
      if (seen.has(current) || current[PATH].isReleased || current[PATH].isNull) return;
      seen.add(current);
      const name = this.#session.typeName(current);
      const option = name === undefined ? undefined : table.get(name);
      if (option === undefined) return;
      current.load(option);
      if (depth >= limit) return;
      // Once the load above has answered, walk what it filled in. This read answers nothing of
      // its own; it settles after the reads queued before it, and its settle queues the next
      // level for the same sync.
      this.#queue.push({
        sort: 'read',
        label: `${current[PATH].label}.loadRecursive`,
        continues: true,
        plan: () => ({ op: 'getPresentation' }),
        settle: () =>
          this.#queue.continuing(() => {
            for (const value of current[LOADED_VALUES]())
              for (const child of Array.isArray(value) ? value : [value])
                if (child instanceof ClientObject) visit(child, depth + 1);
          }),
      });
    };
    visit(object, 0);
  }

  /**
   * Queue a trace message. A sync that fails gives the messages queued before the failed call
   * in `error.traceMessages`, as Office.js does. The message goes nowhere else.
   */
  trace(message: string): void {
    this.#internals.assertUsable('trace');
    if (typeof message !== 'string') failArgument('trace', 'a string');
    this.#traces.push({ message, at: this.#queue.size });
  }

  /**
   * Send everything queued as one atomic batch, then hydrate the answers. An empty queue is not
   * a round trip, so a defensive `sync()` at the end of a run costs nothing.
   *
   * A sync called while an earlier one is in flight takes the calls queued so far at once and
   * sends them when the earlier sync has settled, so syncs apply in the order they were called
   * and no call moves to another batch. A refused sync does not stop the syncs after it.
   *
   * As in Office.js, a refused sync that the caller does not await or catch is an unhandled
   * rejection: the returned promise is the only one that carries the refusal.
   */
  sync<T>(passThroughValue?: T): Promise<T> {
    try {
      this.#internals.assertUsable();
    } catch (error) {
      return Promise.reject(error);
    }
    const actions = this.#queue.take();
    const traces = this.#traces;
    this.#traces = [];
    const start = (): Promise<void> => this.#send(actions, traces);
    const result = this.#syncs === 0 ? start() : this.#tail.then(start);
    this.#syncs += 1;
    const done = (): void => {
      this.#syncs -= 1;
    };
    // The tail handles `result`; the caller gets a promise of its own, so a refusal that the
    // caller ignores stays unhandled instead of being lost.
    this.#tail = result.then(done, done);
    return result.then(() => passThroughValue as T);
  }

  /** Settles when every sync called so far has settled. */
  async #idle(): Promise<void> {
    while (this.#syncs > 0) await this.#tail;
  }

  async #send(queued: readonly QueuedAction[], traces: readonly Trace[]): Promise<void> {
    if (queued.length === 0) return;
    this.#failed = undefined;
    // Every batch of this sync carries one group, so a sync that splits is one undo step.
    this.#historyGroup = `sync-${(sentSyncs += 1)}`;
    try {
      await this.#sendRounds(queued);
    } catch (error) {
      // Trace positions count the calls as queued, before replayed lookups join them.
      throw withTraceMessages(error, tracesBefore(queued, traces, this.#failed));
    }
  }

  async #sendRounds(queued: readonly QueuedAction[]): Promise<void> {
    // A proxy from a refused sync is a path, as in Office.js: its lookup goes again.
    const actions = withReplayedLookups(queued);
    try {
      // A refused argument refuses the whole sync before anything is sent.
      const refused = actions.find((action) => action.refusal);
      if (refused) {
        this.#failed = refused;
        throw refused.refusal;
      }
      // Every read sees the writes queued before it. A read that a later call waits for, after a
      // write, splits the sync there: the calls before it go first, as their own batch.
      let rest: readonly QueuedAction[] = actions;
      let applied = false;
      while (rest.length > 0) {
        const cut = readAfterWriteCut(rest);
        const part = cut < 0 ? rest : rest.slice(0, cut);
        try {
          await this.#sendPart(part);
        } catch (error) {
          throw applied ? partlyApplied(error) : error;
        }
        applied ||= part.some((action) => action.sort === 'write');
        rest = cut < 0 ? [] : rest.slice(cut);
      }
    } catch (error) {
      // Only this sync's calls are dropped. Its lookups stay replayable, so a later batch that
      // names one of their objects looks it up again.
      for (const action of actions)
        if (action.sort === 'read') markReplayable(action);
        else markRefused(action);
      throw error;
    } finally {
      this.#queue.takeContinuations();
      this.#deferred = [];
    }
  }

  /** Send one part of a sync: its awaited reads in earlier rounds, then the rest at once. */
  async #sendPart(actions: readonly QueuedAction[]): Promise<void> {
    const hasWrites = actions.some((action) => action.sort === 'write');
    let expected = hasWrites && this.#readRevision !== null ? this.#readRevision : undefined;
    let remaining = actions;
    // Reads that later calls wait for go first, in rounds pinned to one revision; the writes
    // all go in the last round, so they stay one atomic batch.
    for (let ready = awaitedReads(remaining); ready.length > 0; ready = awaitedReads(remaining)) {
      await this.#dispatch(ready, expected);
      expected = this.#readRevision ?? undefined;
      const sent = new Set(ready);
      remaining = remaining.filter(
        (action) => !sent.has(action) && !(action.sort === 'read' && action.target?.isNull)
      );
    }
    // A listing that reads its items, ahead of every write: it and the reads before it go
    // first, with the item reads, so the writes stay one atomic batch at the end. Sent after
    // the writes, an item read of an object a later write removes would fail after the commit.
    const firstWrite = remaining.findIndex((action) => action.sort === 'write');
    const before = firstWrite < 0 ? [] : remaining.slice(0, firstWrite);
    if (before.some((action) => action.continues)) {
      await this.#dispatch(before, expected);
      await this.#continuations();
      expected = this.#readRevision ?? undefined;
      remaining = remaining
        .slice(firstWrite)
        .filter((action) => !(action.sort === 'read' && action.target?.isNull));
    }
    if (remaining.length > 0) {
      await this.#dispatch(remaining, expected);
      const wrote = remaining.some((action) => action.sort === 'write');
      try {
        await this.#continuations();
      } catch (error) {
        throw wrote ? readAfterCommit(error) : error;
      }
    }
  }

  /** Send the item reads that settled listings queued, until none is left. */
  async #continuations(): Promise<void> {
    for (
      let more = this.#queue.takeContinuations();
      more.length > 0;
      more = this.#queue.takeContinuations()
    )
      await this.#dispatch(more, this.#readRevision ?? undefined);
  }

  async #dispatch(
    actions: readonly QueuedAction[],
    expectedRevision: number | undefined
  ): Promise<void> {
    const planned = planBatch(actions);
    let response: ReturnType<AutomationHost['execute']>;
    try {
      response = this.#session.host.execute({
        operations: planned.operations,
        ...(expectedRevision === undefined ? {} : { expectedRevision }),
        historyGroup: this.#historyGroup,
      });
    } catch (error) {
      // An engine error, not a refusal: reported, then rethrown unchanged.
      if (!isPptxEditorError(error))
        try {
          this.#session.onException?.(error, planned.operations);
        } catch {
          /* an observer must never change what the caller sees */
        }
      throw error;
    }
    if (!response.ok) {
      const error = batchFailure(response, actions, expectedRevision, (action) => {
        this.#failed = action;
      });
      try {
        this.#session.onRefusal?.(error, planned.operations);
      } catch {
        /* an observer must never change what the caller sees */
      }
      throw error;
    }
    this.#readRevision = response.revision;
    settleBatch(actions, response, (action) => {
      this.#failed = action;
    });
    this.#failed = undefined;
    // Work the answers started: wait for all of it, then fail with the first refusal. After a
    // batch that wrote, the writes stay, and the error says so.
    const deferred = this.#deferred;
    this.#deferred = [];
    if (deferred.length === 0) return;
    const outcomes = await Promise.allSettled(deferred);
    const refused = outcomes.find((o): o is PromiseRejectedResult => o.status === 'rejected');
    if (!refused) return;
    throw response.changed ? readAfterCommit(refused.reason) : refused.reason;
  }

  /** @internal The seam proxies reach the context through. */
  get [INTERNALS](): ContextInternals {
    return this.#internals;
  }

  /** @internal Only `run` may build one, and only `run` may end one. */
  static begin(session: RuntimeSession): {
    context: RequestContext;
    adopt: (objects: readonly ClientObject[]) => void;
    idle: () => Promise<void>;
    /** `completed`: the callback returned, so calls still queued are a mistake worth a warning. */
    finish: (completed?: boolean) => void;
  } {
    const context = new RequestContext(session);
    return {
      context,
      adopt(objects) {
        // All or nothing: check every object before any moves, so a refusal leaves each where
        // it was.
        for (const object of objects) context.#checkAdoptable(object);
        for (const object of objects) context.#adopt(object);
      },
      idle() {
        return context.#idle();
      },
      finish(completed = false) {
        context.#finish(completed);
      },
    };
  }

  /**
   * Take over an object a PREVIOUS run tracked. Refused for another runtime's object
   * (`InvalidObjectPath`), for a run still in flight (`ObjectInUse`), and for an untracked
   * object (`InvalidObjectPath`). The revision the source read at comes along, so a write
   * computed from the object's loaded state stays conditional on it.
   */
  #checkAdoptable(object: ClientObject): void {
    const source = object.context;
    if (source === this) return;
    const internals = source[INTERNALS];
    if (internals.session !== this.#session.id)
      throw new PptxEditorError({ code: 'InvalidObjectPath' });
    if (!internals.isFinished()) throw new PptxEditorError({ code: 'ObjectInUse' });
    if (!internals.isTracked(object)) throw new PptxEditorError({ code: 'InvalidObjectPath' });
  }

  #adopt(object: ClientObject): void {
    const source = object.context;
    if (source === this) return;
    this.#checkAdoptable(object);
    const internals = source[INTERNALS];
    const carried = internals.readRevision();
    internals.disown(object);
    object[REBIND](this);
    this.#created.add(object);
    this.#tracked.add(object);
    if (carried !== null)
      this.#readRevision =
        this.#readRevision === null ? carried : Math.min(this.#readRevision, carried);
  }

  /**
   * End of the run. Queued actions are DROPPED, never flushed: a callback that returned without
   * syncing did not ask for its writes. Every object this context made is released except the
   * tracked ones.
   */
  #finish(completed: boolean): void {
    if (this.#finished) return;
    this.#finished = true;
    // Only writes count: a load left without a sync loses nothing the caller asked to change.
    const writes = this.#queue.take().filter((action) => action.sort === 'write');
    if (completed && writes.length > 0)
      warnUnsynced(
        writes.length,
        writes.flatMap((action) => (action.refusal ? [action.refusal] : []))
      );
    this.#queue.clear();
    this.#traces = [];
    for (const object of this.#created) if (!this.#tracked.has(object)) object[RELEASE]();
    this.#created.clear();
  }
}
