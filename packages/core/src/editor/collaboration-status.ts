// The collaboration status tracker and failure-code guard, after DOCX Editor's
// `packages/core/src/collaboration/failure.ts`. One cached snapshot, a new reference only when a
// value changes, so a React host can read it with `useSyncExternalStore`. Session
// implementations (the pro Yjs session, or a host's own) share it.

import type {
  CollaborationFailure,
  CollaborationFailureCode,
  CollaborationStatus,
  CollaborationStatusSnapshot,
} from './modules.ts';

const FAILURE_CODES: { readonly [K in CollaborationFailureCode]: true } = {
  'collaboration-session-destroyed': true,
  'invalid-document-id': true,
  'invalid-identity': true,
  'invalid-identity-color': true,
  'schema-version-mismatch': true,
  'remote-apply-failed': true,
  'unsafe-part-name': true,
  'transport-disconnected': true,
};

/** True when `value` is a {@link CollaborationFailureCode} member. @public */
export function isCollaborationFailureCode(value: string): value is CollaborationFailureCode {
  return Object.hasOwn(FAILURE_CODES, value);
}

/** Bounded detail length, so a peer-supplied message cannot grow a log line without limit. */
const MAX_DETAIL = 500;

function failureOf(code: CollaborationFailureCode, detail?: string): CollaborationFailure {
  const text = detail?.slice(0, MAX_DETAIL);
  return Object.freeze(text ? { code, detail: text } : { code });
}

function sameFailure(
  a: CollaborationFailure | undefined,
  b: CollaborationFailure | undefined
): boolean {
  return a === b || (!!a && !!b && a.code === b.code && a.detail === b.detail);
}

/** A session's status state. @public */
export interface CollaborationStatusTracker {
  /** The current status. */
  status(): CollaborationStatus;
  /** The cached snapshot: the same reference until a value changes. */
  snapshot(): CollaborationStatusSnapshot;
  /**
   * Set the status. `recovering` marks an `error` the session heals by itself. Returns whether
   * the snapshot changed.
   */
  set(
    status: CollaborationStatus,
    code?: CollaborationFailureCode,
    detail?: string,
    recovering?: boolean
  ): boolean;
}

/** Create a status tracker. @public */
export function createCollaborationStatusTracker(
  initial: CollaborationStatus = 'initializing'
): CollaborationStatusTracker {
  let status = initial;
  let reason: CollaborationFailure | undefined;
  let lastFailure: CollaborationFailure | undefined;
  let failureCount = 0;
  let recovering = false;
  let snapshot: CollaborationStatusSnapshot = Object.freeze({
    status,
    reason,
    lastFailure,
    failureCount,
    recovering,
  });
  return {
    status: () => status,
    snapshot: () => snapshot,
    set(next, code, detail, healing = false) {
      const nextReason = code === undefined ? undefined : failureOf(code, detail);
      const nextRecovering = next === 'error' && healing;
      if (status === next && sameFailure(reason, nextReason) && recovering === nextRecovering)
        return false;
      status = next;
      reason = nextReason;
      recovering = nextRecovering;
      if (next === 'error' && nextReason) {
        lastFailure = nextReason;
        failureCount += 1;
      }
      snapshot = Object.freeze({ status, reason, lastFailure, failureCount, recovering });
      return true;
    },
  };
}
