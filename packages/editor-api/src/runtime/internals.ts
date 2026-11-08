/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Adapted from DOCX Editor's `packages/editor-api/src/runtime/internals.ts` (commit 5993bd223945).
//
// The seam between a request context and the proxies it owns, on module-local symbols so no
// consumer can reach it by name.

import type { AutomationCapabilities, AutomationHost } from '@pptx-editor.dev/core/automation';
import type { ActionQueue } from './queue.ts';
import type { RequestContext } from './request-context.ts';

export const INTERNALS: unique symbol = Symbol('pptx-editor.runtime.internals');
export const RELEASE: unique symbol = Symbol('pptx-editor.runtime.release');
export const REBIND: unique symbol = Symbol('pptx-editor.runtime.rebind');
export const PATH: unique symbol = Symbol('pptx-editor.runtime.path');
export const LOAD_REQUEST: unique symbol = Symbol('pptx-editor.runtime.load-request');
export const DEFER_REFUSAL: unique symbol = Symbol('pptx-editor.runtime.defer-refusal');
export const LOADED_VALUES: unique symbol = Symbol('pptx-editor.runtime.loaded-values');

/** What a proxy may ask of the context it belongs to. */
export interface ContextInternals {
  readonly host: AutomationHost;
  readonly capabilities: AutomationCapabilities;
  readonly queue: ActionQueue;
  /**
   * Identity of the runtime behind this context. Compared, never called: it is how adoption
   * refuses an object minted by another runtime, whose handles name another presentation.
   */
  readonly session: object;
  /** Refuse if this context can no longer be used. `target` names what the caller reached for. */
  assertUsable(target?: string): void;
  /** Whether this context's run has ended. */
  isFinished(): boolean;
  /** The revision this context last read at, or `null` if it has never read. */
  readRevision(): number | null;
  register(object: RuntimeManagedObject): void;
  track(object: RuntimeManagedObject): void;
  untrack(object: RuntimeManagedObject): void;
  isTracked(object: RuntimeManagedObject): boolean;
  /** Give up every claim on an object: it belongs to another context now. */
  disown(object: RuntimeManagedObject): void;
  /**
   * Work that a settled answer starts and the sync must wait for (a picture the host draws).
   * The sync that dispatched the answer awaits it before it continues, and fails if it fails.
   */
  defer(work: Promise<void>): void;
}

/** What a context may do to a proxy it owns. */
export interface RuntimeManagedObject {
  /** @internal Drop this object's presentation reference; the run that owned it has ended. */
  [RELEASE](): void;
  /** @internal Adoption: this object now belongs to another run's context on the same runtime. */
  [REBIND](context: RequestContext): void;
}

/** The internals of a context, for the modules that hold one. */
export function internalsOf(context: RequestContext): ContextInternals {
  return context[INTERNALS];
}
