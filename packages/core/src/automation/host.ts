// The one host implementation, after DOCX Editor's `packages/core/src/automation/host.ts`.
//
// INTERNAL composition. Both shipped hosts (`server-host.ts` here, and the browser host in the
// editor lane) call this with a port over a `PptxDocumentStore`, and the port commits through
// that store's `transact`, the only write path.
//
// THE BATCH RULE: run every operation in order against one working package, then commit every
// op in one transaction. If an operation is refused, nothing is committed, and the response
// names the failing operation and marks every other one `skipped`.

import type { TextMeasurer } from '../layout/index.ts';
import type { PptxOp, PptxPackage } from '../store/index.ts';
import { toBase64 } from './commands/base64.ts';
import { CommandRefusal, type OpBatch } from './commands/op-batch.ts';
import { createHandleTable } from './handles.ts';
import { AUTOMATION_COMMANDS, AUTOMATION_SELECTION_OPERATIONS } from './operations.ts';
import type { AutomationOperation } from './operations.ts';
import { imageTarget, type AutomationImageTarget } from './plan-images.ts';
import { planOperation, type AutomationSelectionPort, type BatchPlanContext } from './plan.ts';
import { sweepBindings, type AutomationBindings } from './plan-bindings.ts';
import type {
  AutomationBatchRequest,
  AutomationBatchResponse,
  AutomationCapabilities,
  AutomationChangeEvent,
  AutomationError,
  AutomationErrorCode,
  AutomationHost,
  AutomationImageRequest,
  AutomationImageResult,
  AutomationOperationResult,
  AutomationSaveResult,
  AutomationValue,
} from './protocol.ts';
import { AutomationFailure, NullReference } from './targets.ts';

/** The most operations one batch may carry. */
export const MAX_BATCH_OPERATIONS = 10_000;

/** What a host needs from the document it acts on. */
export interface AutomationDocumentPort {
  /** Monotonic revision of the document. */
  revision(): number;
  /**
   * Which document the port holds now, for a port that can change it (an editor that loads
   * another presentation). A change forgets every handle, so an object of the previous
   * document never addresses an object of the new one that has the same id.
   */
  document?(): number;
  /** The committed package, or null while there is none (an editor between mounts). */
  currentPackage(): PptxPackage | null;
  /** A working batch over `pkg` that keeps new part names unique for this document. */
  batch(pkg: PptxPackage): OpBatch;
  /**
   * Commit ops as one transaction. A port with an undo history makes consecutive commits of one
   * `historyGroup` a single undo step.
   */
  apply(
    ops: readonly PptxOp[],
    historyGroup?: string
  ):
    | { readonly ok: true; readonly changed: boolean }
    | { readonly ok: false; readonly reason: string };
  /**
   * Called once after a batch that changed the document has committed and moved the selection.
   * A browser port records the selection after the batch for redo.
   */
  committed?(): void;
  /** The document as `.pptx` bytes, or null when there is none. */
  save(): Uint8Array | null;
  /** Change notification for commits by anyone. */
  subscribe(listener: () => void): () => void;
  /** The measurer text writes re-fit with, when the host has one. */
  measurer?(): TextMeasurer | undefined;
  /**
   * True while the document is in a collaboration room. A table structure edit then marks the
   * table, as the editor's own table commands do, so a save repairs the merge markup that a
   * concurrent edit of a peer leaves.
   */
  collaborating?(): boolean;
  /** The reader's selection, for a browser host. */
  readonly selection?: AutomationSelectionPort;
  /**
   * A PNG picture of a slide or of one shape of it, from the committed package, for a host with
   * the `images` capability. Rejects with a `CommandRefusal` when it cannot draw.
   */
  renderImage?(target: AutomationImageTarget): Promise<Uint8Array>;
  /** Release the port. */
  dispose(): void;
}

/** The port and the capabilities a host claims. */
export interface AutomationHostComposition {
  readonly port: AutomationDocumentPort;
  readonly capabilities: AutomationCapabilities;
}

const SKIPPED: AutomationOperationResult = Object.freeze({ status: 'skipped' as const });

function automationError(
  code: AutomationErrorCode,
  message: string,
  detail?: string
): AutomationError {
  return Object.freeze(detail === undefined ? { code, message } : { code, message, detail });
}

/** A batch that answered nothing: the failure at its index, `skipped` everywhere else. */
function refuse(
  count: number,
  index: number,
  error: AutomationError,
  revision: number
): AutomationBatchResponse {
  const results: AutomationOperationResult[] = [];
  for (let at = 0; at < Math.max(count, 1); at += 1)
    results.push(at === index ? { status: 'error', error } : SKIPPED);
  return { ok: false, results: count === 0 ? [] : results, revision, changed: false };
}

/** A random id for the open presentation, for `Presentation.id`. */
function presentationId(): string {
  const bytes = globalThis.crypto.getRandomValues(new Uint8Array(8));
  let hex = '';
  for (const byte of bytes) hex += byte.toString(16).padStart(2, '0');
  return hex;
}

/** Build a host over a port. */
export function createAutomationHost(composition: AutomationHostComposition): AutomationHost {
  const { port } = composition;
  const capabilities = Object.freeze({ ...composition.capabilities });
  const handles = createHandleTable();
  // Bindings live with the host, in memory, and never enter the package (`plan-bindings.ts`).
  let bindings: AutomationBindings = new Map();
  const listeners = new Set<(event: AutomationChangeEvent) => void>();
  const id = presentationId();
  let disposed = false;
  let executing = 0;
  let pending: AutomationChangeEvent | null = null;
  let document = port.document?.();

  // Every change, by anyone (an automation batch, an edit in the editor, an undo), checks the
  // bindings, so a binding whose shape is removed is orphaned before a new shape takes its id.
  const sweep = (): void => {
    const pkg = port.currentPackage();
    if (pkg && bindings.size > 0) sweepBindings(bindings, pkg);
  };
  const unsubscribeBindings = port.subscribe(sweep);
  const unsubscribePort = capabilities.events
    ? port.subscribe(() => {
        if (disposed) return;
        const event = Object.freeze({ revision: port.revision() });
        if (executing > 0) pending = event;
        else for (const listener of [...listeners]) listener(event);
      })
    : () => {};

  const executeBatch = (request: AutomationBatchRequest): AutomationBatchResponse => {
    const operations: readonly AutomationOperation[] = Array.isArray(request?.operations)
      ? request.operations
      : [];
    const revision = port.revision();
    const count = operations.length;
    const now = port.document?.();
    if (now !== document) {
      document = now;
      handles.clear();
      bindings = new Map();
    }
    if (disposed)
      return refuse(count, 0, automationError('disposed', 'this host was disposed'), revision);
    if (count > MAX_BATCH_OPERATIONS)
      return refuse(
        count,
        0,
        automationError('invalid-argument', 'too many operations', 'operations'),
        revision
      );
    if (count === 0) return { ok: true, results: [], revision, changed: false };
    if (!capabilities.document)
      return refuse(
        count,
        0,
        automationError('unsupported-capability', 'this host has no document', 'document'),
        revision
      );
    const pkg = port.currentPackage();
    if (!pkg)
      return refuse(
        count,
        0,
        automationError('document-unavailable', 'this host holds no presentation right now'),
        revision
      );
    if (request.expectedRevision !== undefined && request.expectedRevision !== revision)
      return refuse(
        count,
        0,
        automationError(
          'stale-revision',
          'the presentation moved since that revision',
          `expected ${String(request.expectedRevision)}, at ${revision}`
        ),
        revision
      );

    const results: AutomationValue[] = [];
    const ctx: BatchPlanContext = {
      batch: port.batch(pkg),
      handles,
      results,
      presentationId: id,
      bindings: new Map(bindings),
      afterCommit: [],
      selectionAfter: {},
      ...(port.measurer?.() ? { measurer: port.measurer()! } : {}),
      ...(port.collaborating?.() ? { collaborating: true } : {}),
      ...(capabilities.selection && port.selection ? { selection: port.selection } : {}),
    };
    let firstCommand = -1;
    for (let index = 0; index < count; index += 1) {
      const operation = operations[index];
      const name = typeof operation === 'object' && operation !== null ? operation.op : undefined;
      const command = AUTOMATION_COMMANDS.has(name as AutomationOperation['op']);
      if (command && firstCommand < 0) firstCommand = index;
      if (
        AUTOMATION_SELECTION_OPERATIONS.has(name as AutomationOperation['op']) &&
        !capabilities.selection
      )
        return refuse(
          count,
          index,
          automationError('unsupported-capability', 'this host has no selection', 'selection'),
          revision
        );
      if (name === 'prepareImage' && !(capabilities.images && port.renderImage))
        return refuse(
          count,
          index,
          automationError('unsupported-capability', 'this host draws no pictures', 'images'),
          revision
        );
      try {
        if (!name) throw new AutomationFailure('unknown-operation', 'not an operation');
        const before = ctx.batch.pkg;
        results.push(planOperation(ctx, operation!));
        if (ctx.batch.pkg !== before && ctx.bindings.size > 0)
          sweepBindings(ctx.bindings, ctx.batch.pkg);
      } catch (e) {
        if (e instanceof NullReference) {
          if (command || name === 'exportSlides' || name === 'prepareImage')
            return refuse(
              count,
              index,
              automationError('invalid-handle', 'the object was not found', 'null-object'),
              revision
            );
          results.push({ kind: 'null' });
          continue;
        }
        if (e instanceof AutomationFailure)
          return refuse(count, index, automationError(e.code, e.message, e.detail), revision);
        // A planning bug must not commit a partial batch, and must not leak engine state.
        return refuse(
          count,
          index,
          automationError('transaction-refused', 'the operation failed', 'internal'),
          revision
        );
      }
    }

    let changed = false;
    if (ctx.batch.ops.length > 0) {
      const applied = port.apply(ctx.batch.ops, request.historyGroup);
      if (!applied.ok)
        return refuse(
          count,
          firstCommand < 0 ? 0 : firstCommand,
          automationError(
            'transaction-refused',
            'the store refused the transaction',
            applied.reason
          ),
          revision
        );
      changed = applied.changed;
    }
    bindings = ctx.bindings;
    // The commit made new elements: check the bindings against the committed package.
    sweep();
    for (const after of ctx.afterCommit) after();
    if (changed) port.committed?.();
    return {
      ok: true,
      results: results.map((value) => ({ status: 'ok', value })),
      revision: port.revision(),
      changed,
    };
  };

  /** Check the request against the committed package, then draw it after this call returns. */
  const imageOf = (request: AutomationImageRequest): AutomationImageTarget | AutomationError => {
    if (disposed) return automationError('disposed', 'this host was disposed');
    if (!capabilities.images || !port.renderImage)
      return automationError('unsupported-capability', 'this host draws no pictures', 'images');
    if (port.document?.() !== document)
      return automationError('invalid-handle', 'that object is gone', 'gone');
    const pkg = port.currentPackage();
    if (!pkg)
      return automationError('document-unavailable', 'this host holds no presentation right now');
    const target = handles.resolve(request?.target);
    if (!target) return automationError('invalid-handle', 'a handle of another kind or host');
    try {
      return imageTarget(pkg, target, request.width, request.height);
    } catch (e) {
      if (e instanceof AutomationFailure) return automationError(e.code, e.message, e.detail);
      return automationError('transaction-refused', 'the picture could not be drawn', 'internal');
    }
  };

  const renderImage = async (request: AutomationImageRequest): Promise<AutomationImageResult> => {
    const target = imageOf(request);
    if ('code' in target) return { ok: false, error: target };
    try {
      return { ok: true, base64: toBase64(await port.renderImage!(target)) };
    } catch (e) {
      // A refusal of the renderer has an engine reason; it selects one public detail only.
      const why = e instanceof CommandRefusal ? 'image-failed' : 'internal';
      return {
        ok: false,
        error: automationError('transaction-refused', 'the picture could not be drawn', why),
      };
    }
  };

  return {
    capabilities,
    revision: () => port.revision(),
    execute(request) {
      executing += 1;
      try {
        return executeBatch(request);
      } finally {
        executing -= 1;
        if (executing === 0 && pending) {
          const event = pending;
          pending = null;
          for (const listener of [...listeners]) listener(event);
        }
      }
    },
    save(): AutomationSaveResult {
      if (disposed)
        return { ok: false, error: automationError('disposed', 'this host was disposed') };
      if (!capabilities.save)
        return {
          ok: false,
          error: automationError('unsupported-capability', 'this host cannot save', 'save'),
        };
      const bytes = port.save();
      if (!bytes)
        return {
          ok: false,
          error: automationError(
            'document-unavailable',
            'this host holds no presentation right now'
          ),
        };
      return { ok: true, bytes };
    },
    renderImage: (request) => renderImage(request),
    subscribe(listener) {
      if (disposed || !capabilities.events) return () => {};
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      listeners.clear();
      unsubscribePort();
      unsubscribeBindings();
      port.dispose();
      handles.clear();
      bindings.clear();
    },
  };
}
