// The client side of one processing job, after DOCX Editor's
// examples/document-refresh/src/refresh-job.ts: capture, receive cumulative results, check
// their identity, apply them, and finish the submission.

import type {
  DocumentRefresh,
  RefreshChangeInput,
  RefreshResult,
  RefreshSubmission,
  RefreshUpdateMode,
} from '@pptx-editor.dev/react';

/** The id of the presentation that this example opens. */
export const DOCUMENT_ID = 'quarterly-review';

export interface ServerUpdate {
  documentId: string;
  submissionId: string;
  sequence: number;
  /** The complete PPTX file, as base64. */
  bytes: string;
  changes: RefreshChangeInput[];
  failures: string[];
}

export interface RefreshJobOptions {
  /** `'edit'` adds each result as one undo step. `'reload'` (default) resets undo history. */
  mode?: RefreshUpdateMode;
  /**
   * Omit the processor change locations, so a registered `reviewModule()` computes the changes.
   * Default: false.
   */
  computeChanges?: boolean;
}

/** Transport identity belongs to the application. Submission tokens stay in the browser. */
export async function runRefreshJob(
  refresh: DocumentRefresh,
  signal: AbortSignal,
  receive: (submission: RefreshSubmission) => AsyncIterable<ServerUpdate>,
  /** Receives each result with the server update it belongs to. */
  report: (result: RefreshResult, update: ServerUpdate) => void,
  { mode = 'reload', computeChanges = false }: RefreshJobOptions = {}
): Promise<void> {
  const submission = await refresh.capture();
  try {
    if (signal.aborted) return;
    for await (const update of receive(submission)) {
      if (signal.aborted) break;
      if (update.documentId !== DOCUMENT_ID || update.submissionId !== submission.id)
        throw new Error('identity-mismatch');
      const accepted = {
        submission,
        sequence: update.sequence,
        bytes: Uint8Array.from(atob(update.bytes), (c) => c.charCodeAt(0)),
        ...(computeChanges ? {} : { changes: update.changes }),
        failures: update.failures,
      };
      let result = await refresh.applyUpdate({ ...accepted, mode });
      // The presentation and its history did not change, so a reload of the same result can apply.
      if (!result.ok && result.code === 'edit-unsupported')
        result = await refresh.applyUpdate(accepted);
      if (signal.aborted) break;
      report(result, update);
      // A late result cannot replace newer output. Keep receiving the job's remaining results.
      if (!result.ok && result.code !== 'out-of-order') break;
    }
  } finally {
    refresh.finish(submission);
  }
}

/** Two complete cumulative files. Production transport can use polling or a result stream. */
export async function* receiveSampleUpdates(
  submission: RefreshSubmission,
  signal: AbortSignal,
  round: number,
  includeLateResult: boolean
): AsyncGenerator<ServerUpdate> {
  let first: ServerUpdate | undefined;
  for (const sequence of [1, 2]) {
    const response = await fetch(`/api/update?round=${round}&sequence=${sequence}`, {
      method: 'POST',
      headers: { 'X-Submission-Id': submission.id },
      signal,
      body: submission.bytes,
    });
    if (!response.ok) throw new Error('request-failed');
    const update: ServerUpdate = await response.json();
    first ??= update;
    yield update;
  }
  if (includeLateResult && first) yield first;
}
