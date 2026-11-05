/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// After DOCX Editor's `packages/editor-api/src/runtime/batch.ts` (commit 5993bd223945).
//
// Planning, refusal mapping, and positional hydration of one batch.

import type {
  AutomationBatchResponse,
  AutomationError,
  AutomationErrorCode,
  AutomationOperation,
} from '@pptx-editor.dev/core/automation';
import { hostDiagnosticOf, PptxEditorError, type PptxEditorErrorCode } from './errors.ts';
import { NeedsProducer, WaitForRead } from './object-path.ts';
import type { QueuedAction } from './queue.ts';

export interface PlannedBatch {
  readonly operations: readonly AutomationOperation[];
  readonly hasRead: boolean;
  readonly hasWrite: boolean;
}

/**
 * Plan every action, in order. All or nothing before anything is sent: if one action cannot
 * name its target, the throw happens here and the host is never called.
 */
export function planBatch(actions: readonly QueuedAction[]): PlannedBatch {
  const index = new Map<object, number>();
  actions.forEach((action, i) => index.set(action, i));
  const operations: AutomationOperation[] = [];
  let hasRead = false;
  let hasWrite = false;
  for (const action of actions) {
    operations.push(action.plan((producer) => index.get(producer)));
    if (action.sort === 'write') hasWrite = true;
    else hasRead = true;
  }
  return { operations, hasRead, hasWrite };
}

/**
 * `actions` with the lookups they name from a refused sync put back, each just before the first
 * call that names it, so a proxy from a refused sync addresses its object again and the lookup
 * still sees the writes queued before that call. Planning errors other than `NeedsProducer` are
 * left for the real planning to raise.
 */
export function withReplayedLookups(actions: readonly QueuedAction[]): QueuedAction[] {
  let list = [...actions];
  for (let guard = 0; guard <= actions.length + 10_000; guard += 1) {
    const index = new Map<object, number>();
    list.forEach((action, i) => index.set(action, i));
    let needed: QueuedAction | undefined;
    let at = 0;
    for (const [i, action] of list.entries()) {
      try {
        action.plan((producer) => index.get(producer));
      } catch (error) {
        if (error instanceof NeedsProducer) {
          needed = error.producer as QueuedAction;
          at = i;
          break;
        }
      }
    }
    if (!needed) return list;
    list = [...list.slice(0, at), needed, ...list.slice(at)];
  }
  return list;
}

/**
 * Where a sync must split so that every read sees the writes queued before it: the position of
 * the first read that a later call waits for (`WaitForRead`) and that has a write before it in
 * the queue, or -1. The actions before it are sent first, as their own batch; the read then goes
 * in an earlier round of the rest, after those writes.
 */
export function readAfterWriteCut(actions: readonly QueuedAction[]): number {
  const index = new Map<object, number>();
  actions.forEach((action, i) => index.set(action, i));
  const firstWrite = actions.findIndex((action) => action.sort === 'write');
  if (firstWrite < 0) return -1;
  let cut = -1;
  for (const action of actions) {
    try {
      action.plan((producer) => index.get(producer));
    } catch (error) {
      if (!(error instanceof WaitForRead)) continue;
      const at = index.get(error.producer);
      if (at !== undefined && at > firstWrite && (cut < 0 || at < cut)) cut = at;
    }
  }
  return cut;
}

/**
 * The reads to send in an earlier round, in queue order: those that other actions of this batch
 * wait for (`WaitForRead`), with the reads they name by result reference. Empty when the whole
 * batch plans in one round. A read that waits on an object a write of the same batch creates
 * cannot go early, because the writes stay one atomic batch: it refuses with
 * `InvalidObjectPath`, and a `sync()` between the write and the read resolves it. Any other
 * planning refusal propagates.
 */
export function awaitedReads(actions: readonly QueuedAction[]): QueuedAction[] {
  const index = new Map<object, number>();
  actions.forEach((action, i) => index.set(action, i));
  const awaited: QueuedAction[] = [];
  const waiting = new Set<QueuedAction>();
  const uses = new Map<QueuedAction, QueuedAction[]>();
  for (const action of actions) {
    const used: QueuedAction[] = [];
    uses.set(action, used);
    try {
      action.plan((producer) => {
        const at = index.get(producer);
        if (at !== undefined) used.push(actions[at]!);
        return at;
      });
    } catch (error) {
      if (!(error instanceof WaitForRead)) throw error;
      awaited.push(error.producer as QueuedAction);
      waiting.add(action);
    }
  }
  if (awaited.length === 0) return [];
  const round = new Set<QueuedAction>();
  // The read and every action it names, if none of them waits or writes.
  const closure = (start: QueuedAction): Set<QueuedAction> | null => {
    const seen = new Set<QueuedAction>();
    const stack = [start];
    while (stack.length > 0) {
      const action = stack.pop()!;
      if (seen.has(action) || round.has(action)) continue;
      if (waiting.has(action)) return null;
      if (action.sort === 'write')
        throw new PptxEditorError({ code: 'InvalidObjectPath', target: start.label });
      seen.add(action);
      for (const used of uses.get(action) ?? []) stack.push(used);
    }
    return seen;
  };
  for (const read of awaited) for (const action of closure(read) ?? []) round.add(action);
  if (round.size === 0)
    throw new PptxEditorError({ code: 'InvalidObjectPath', target: [...waiting][0]?.label });
  return actions.filter((action) => round.has(action));
}

/**
 * The host's refusal codes, mapped onto this runtime's. A total mapping: a new protocol code
 * fails to compile here rather than reaching consumers as `GeneralException`.
 */
const HOST_CODES: Readonly<Record<AutomationErrorCode, PptxEditorErrorCode>> = Object.freeze({
  'stale-revision': 'StaleDocument',
  'invalid-handle': 'InvalidObjectPath',
  'invalid-argument': 'InvalidArgument',
  'unsupported-capability': 'NotSupported',
  disposed: 'RuntimeDisposed',
  'transaction-refused': 'GeneralException',
  'unknown-operation': 'NotSupported',
  'unsupported-content': 'NotSupported',
  'item-not-found': 'ItemNotFound',
  'document-unavailable': 'DocumentUnavailable',
});

export interface HostFailureContext {
  readonly target?: string;
  readonly expectedRevision?: number;
  readonly actualRevision?: number;
}

/**
 * One host error as one of ours. Nothing from `error.message` or `error.detail` survives: the
 * detail only selects a public diagnostic.
 */
export function hostFailure(
  error: AutomationError,
  context: HostFailureContext = {}
): PptxEditorError {
  const code = HOST_CODES[error.code] ?? 'GeneralException';
  const stale = code === 'StaleDocument';
  const diagnostic = hostDiagnosticOf(error.code, error.detail);
  return new PptxEditorError({
    code,
    ...(diagnostic === undefined ? {} : { diagnostic }),
    ...(context.target === undefined ? {} : { target: context.target }),
    ...(stale && context.expectedRevision !== undefined
      ? { expectedRevision: context.expectedRevision }
      : {}),
    ...(stale && context.actualRevision !== undefined
      ? { actualRevision: context.actualRevision }
      : {}),
  });
}

/** The error for a refused batch, named after the consumer call that failed. */
export function batchFailure(
  response: AutomationBatchResponse,
  actions: readonly QueuedAction[],
  expectedRevision?: number,
  onAction?: (action: QueuedAction) => void
): PptxEditorError {
  for (let i = 0; i < response.results.length; i += 1) {
    const result = response.results[i];
    if (result?.status !== 'error') continue;
    const label = actions[i]?.label;
    if (actions[i]) onAction?.(actions[i]!);
    return hostFailure(result.error, {
      ...(label === undefined ? {} : { target: label }),
      ...(expectedRevision === undefined ? {} : { expectedRevision }),
      actualRevision: response.revision,
    });
  }
  return new PptxEditorError({ code: 'GeneralException' });
}

/**
 * Hydrate a successful response into the actions that asked for it. Positional and
 * length-checked: a response that does not match the request cannot be matched up at all.
 */
export function settleBatch(
  actions: readonly QueuedAction[],
  response: AutomationBatchResponse,
  onAction?: (action: QueuedAction) => void
): void {
  if (response.results.length !== actions.length)
    throw new PptxEditorError({ code: 'GeneralException' });
  for (let i = 0; i < actions.length; i += 1) {
    const result = response.results[i]!;
    const action = actions[i]!;
    // The action a settle failure names, for the trace messages of the error.
    onAction?.(action);
    if (result.status !== 'ok')
      throw new PptxEditorError({ code: 'GeneralException', target: action.label });
    action.settle(result.value);
  }
}
