/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// After DOCX Editor's `packages/editor-api/src/runtime/runtime.ts` (commit 5993bd223945).
//
// A runtime: one presentation host, many runs. Runs are ISOLATED, not serialized: every `run`
// gets its own context and queue, so two runs cannot interleave into one batch.

import type { AutomationHost, AutomationOperation } from '@pptx-editor.dev/core/automation';
import { Application } from '../model/application.ts';
import { Presentation } from '../model/presentation.ts';
import { TYPE_NAMES, typeNameOf } from '../model/type-names.ts';
import { hostFailure } from './batch.ts';
import { ClientObject } from './client-object.ts';
import { fail, failArgument, type PptxEditorError } from './errors.ts';
import { RequestContext, type RuntimeSession } from './request-context.ts';

/**
 * What a `run` callback is given, and what it may answer with.
 *
 * @public
 */
export type RunCallback<T> = (context: RequestContext) => Promise<T>;

/**
 * Capabilities of a runtime, frozen for its lifetime.
 *
 * @public
 */
export interface PresentationCapabilities {
  /** There is a presentation to address at all. */
  readonly document: boolean;
  /** `save()` is offered: true for a server runtime, false for one borrowing an editor. */
  readonly save: boolean;
  /** The host raises change events. */
  readonly events: boolean;
  /** The host has a reader's selection to read or move. */
  readonly selection: boolean;
  /** Text writes re-fit text bodies by their autofit rules (ECMA-376 Part 1, 21.1.2.1). */
  readonly autofit: boolean;
  /**
   * `getImageAsBase64` draws pictures of slides and shapes: true for a runtime over a live editor
   * in a browser and for a server runtime created with a `rasterize` option, false otherwise.
   */
  readonly images: boolean;
}

/**
 * A runtime: one presentation host, many runs.
 *
 * Disposal is final: {@link PptxEditorRuntime.dispose} releases the host once and is safe to call
 * again, and every later `run` fails with `RuntimeDisposed`.
 *
 * @public
 */
export interface PptxEditorRuntime {
  /** What the presentation host behind this runtime can do. */
  readonly capabilities: PresentationCapabilities;
  /** Run one batch of work against the presentation. Answers with the callback's value. */
  run<T>(callback: RunCallback<T>): Promise<T>;
  /** Run one batch of work, adopting objects a previous run tracked. */
  run<T>(object: ClientObject | readonly ClientObject[], callback: RunCallback<T>): Promise<T>;
  /** Release the host. Idempotent. */
  dispose(): void;
}

/**
 * A runtime over PPTX bytes rather than a live editor: what `PptxEditor.createServer` answers.
 *
 * @public
 */
export interface PptxEditorServerRuntime extends PptxEditorRuntime {
  /** The current presentation as a fresh, caller-owned PPTX byte array. */
  save(): Promise<Uint8Array>;
}

export interface CreateRuntimeOptions {
  readonly host: AutomationHost;
  /** Whether this runtime offers `save()`. */
  readonly save: boolean;
  /** Told about each refused batch before it is thrown. */
  readonly onRefusal?: (error: PptxEditorError, operations: readonly AutomationOperation[]) => void;
  /** Told about each batch that threw an engine error, before it propagates. */
  readonly onException?: (error: unknown, operations: readonly AutomationOperation[]) => void;
}

export function createRuntime(
  options: CreateRuntimeOptions & { save: true }
): PptxEditorServerRuntime;
export function createRuntime(options: CreateRuntimeOptions): PptxEditorRuntime;
export function createRuntime(options: CreateRuntimeOptions): PptxEditorServerRuntime {
  const host = options.host;
  const capabilities: PresentationCapabilities = Object.freeze({
    document: host.capabilities.document,
    save: options.save && host.capabilities.save,
    events: host.capabilities.events,
    selection: host.capabilities.selection,
    autofit: host.capabilities.autofit,
    images: host.capabilities.images === true && typeof host.renderImage === 'function',
  });
  let disposed = false;
  const assertLive = (target?: string): void => {
    if (disposed) fail({ code: 'RuntimeDisposed', ...(target === undefined ? {} : { target }) });
  };
  const session: RuntimeSession = {
    host,
    capabilities,
    ...(options.onRefusal ? { onRefusal: options.onRefusal } : {}),
    ...(options.onException ? { onException: options.onException } : {}),
    id: {},
    assertLive,
    openPresentation: (context) => Presentation.open(context),
    openApplication: (context) => Application.open(context),
    typeName: typeNameOf,
    typeNames: TYPE_NAMES,
  };

  async function run<T>(
    first: RunCallback<T> | ClientObject | readonly ClientObject[],
    second?: RunCallback<T>
  ): Promise<T> {
    assertLive();
    const callback = typeof first === 'function' ? first : second;
    if (typeof callback !== 'function')
      failArgument('run', 'a callback, alone or after the objects to adopt');
    const adopted: readonly ClientObject[] =
      typeof first === 'function'
        ? []
        : Array.isArray(first)
          ? (first as readonly ClientObject[])
          : first instanceof ClientObject
            ? [first]
            : failArgument('run', 'an object of this API or an array of them, then a callback');
    if (!adopted.every((object) => object instanceof ClientObject))
      failArgument('run', 'objects of this API only');
    const { context, adopt, idle, finish } = RequestContext.begin(session);
    let completed = false;
    try {
      adopt(adopted);
      const value = await callback(context);
      completed = true;
      return value;
    } finally {
      // A sync the callback started and did not await still applies before the run ends.
      await idle();
      finish(completed);
    }
  }

  const runtime: PptxEditorRuntime = {
    capabilities,
    run,
    dispose() {
      if (disposed) return;
      disposed = true;
      host.dispose();
    },
  };
  if (!options.save) return runtime as PptxEditorServerRuntime;
  return {
    ...runtime,
    async save(): Promise<Uint8Array> {
      assertLive('save');
      if (!capabilities.save) fail({ code: 'NotSupported', target: 'save' });
      const saved = host.save();
      if (!saved.ok) throw hostFailure(saved.error, { target: 'save' });
      return saved.bytes;
    },
  };
}
