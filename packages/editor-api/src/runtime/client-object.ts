/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Adapted from DOCX Editor's `packages/editor-api/src/runtime/client-object.ts` (commit
// 5993bd223945).
//
// A proxy is three things: the context it belongs to, the path that says whether it can be
// addressed, and the properties a completed `load` filled in.
//
// LOADED VALUES LIVE IN A MAP. Property names reaching `load(...)` can come from a consumer's own
// data, and a name like `__proto__` assigned into a plain object is the prototype-pollution
// hazard the repository audits for. A `Map` has no such key.
//
// A RELEASED PROXY STILL ANSWERS WHAT IT ALREADY KNEW. A property loaded before the run ended is
// a copy the consumer already has. Anything that would talk to the presentation (`load`, a write,
// a method) refuses with `InvalidObjectPath`.

import type {
  AutomationOperation,
  AutomationRef,
  AutomationValue,
} from '@pptx-editor.dev/core/automation';
import { deferArgumentChecks } from './deferred-checks.ts';
import {
  diagnosticOf,
  fail,
  failArgument,
  type PptxEditorDiagnostic,
  type PptxEditorError,
} from './errors.ts';
import {
  DEFER_REFUSAL,
  INTERNALS,
  LOAD_REQUEST,
  LOADED_VALUES,
  PATH,
  REBIND,
  RELEASE,
  type ContextInternals,
  type RuntimeManagedObject,
} from './internals.ts';
import { resolveLoadOption, type LoadOption, type ResolvedLoadOptions } from './load-options.ts';
import type { IndexOf, ObjectPath } from './object-path.ts';
import type { ActionSort } from './queue.ts';
import type { RequestContext } from './request-context.ts';

/** A loaded value as plain data: a proxy becomes its own `toJSON()`. */
function plain(value: unknown): unknown {
  return value instanceof ClientObject ? value.toJSON() : value;
}

/**
 * The base every presentation proxy extends.
 *
 * A released proxy still answers what it already knew: a property loaded before the run ended is
 * served from memory. Anything that would talk to the presentation (`load`, a write, a method)
 * refuses with `InvalidObjectPath`.
 *
 * @public
 */
export abstract class ClientObject implements RuntimeManagedObject {
  #context: RequestContext;
  #internals: ContextInternals;
  readonly #path: ObjectPath;
  readonly #nullable: boolean;
  readonly #loaded = new Map<string, unknown>();
  #nestedTaken = false;
  #loadingOrNull = false;
  /** Why a load through a collection skipped this object, when it did. */
  #skipped: PptxEditorDiagnostic | undefined;

  /** @internal */
  protected constructor(
    context: RequestContext,
    path: ObjectPath,
    options: { readonly nullable?: boolean } = {}
  ) {
    this.#context = context;
    this.#internals = context[INTERNALS];
    this.#path = path;
    this.#nullable = options.nullable === true;
    this.#internals.register(this);
    // Argument checks of this class's calls and property writes go to the sync.
    deferArgumentChecks(Object.getPrototypeOf(this) as object, ClientObject.prototype);
  }

  /** The request context this object currently belongs to. */
  get context(): RequestContext {
    return this.#context;
  }

  /**
   * Whether this object turned out not to exist. Only an answer, never a guess: an object from a
   * `getItemOrNullObject` has no verdict until the sync that looked for it, and reading one
   * before then is `PropertyNotLoaded`.
   */
  get isNullObject(): boolean {
    if (this.#path.isReleased) fail({ code: 'InvalidObjectPath', target: this.#path.label });
    if (!this.#nullable) return false;
    if (this.#path.isPending)
      fail({ code: 'PropertyNotLoaded', target: `${this.#path.label}.isNullObject` });
    return this.#path.isNull;
  }

  /**
   * Queue the reads that fill in the selected properties. Returns `this`; the values are
   * readable after the next `sync()`.
   */
  load(option?: LoadOption): this {
    if (this.#path.isReleased) fail({ code: 'InvalidObjectPath', target: this.#path.label });
    this.#internals.assertUsable(this.#path.label);
    return this[LOAD_REQUEST](resolveLoadOption(option, this.#path.label));
  }

  /**
   * @internal Queue the reads of a resolved request: what `load(...)` does after it reads its
   * option, and what a navigation property or a collection item gets from the path above it.
   * All or nothing: a refusal drops what this call queued.
   */
  [LOAD_REQUEST](request: ResolvedLoadOptions): this {
    if (this.#path.isReleased) fail({ code: 'InvalidObjectPath', target: this.#path.label });
    this.#internals.assertUsable(this.#path.label);
    if (request.items !== undefined && !this.loadsItemPaths)
      failArgument(`${this.#path.label}.items`, "'items/' paths on a collection only");
    // A null object has nothing to load; Office.js lets the load pass.
    if (this.#path.isNull) return this;
    const restore = this.#internals.queue.checkpoint();
    this.#nestedTaken = false;
    const previous = this.#loadingOrNull;
    this.#loadingOrNull = request.orNull === true;
    try {
      this.onLoad(request);
      if (!this.#nestedTaken) {
        const deeper =
          request.nested.keys().next().value ?? request.items?.nested.keys().next().value;
        if (deeper !== undefined)
          failArgument(
            `${this.#path.label}.${deeper}`,
            'a path through the navigation properties of this object'
          );
      }
    } catch (error) {
      restore();
      throw error;
    } finally {
      this.#loadingOrNull = previous;
    }
    return this;
  }

  /**
   * @internal True while this object plans a load that a collection's items started: a text
   * read then answers null for a shape that cannot hold text, and fills nothing.
   */
  protected get loadingOrNull(): boolean {
    return this.#loadingOrNull;
  }

  /**
   * @internal Say that this load handed its deeper paths on. A load whose planner does not call
   * this refuses any deeper path, so no path is dropped without a word.
   */
  protected takeNestedPaths(): void {
    this.#nestedTaken = true;
  }

  /** The loaded properties as a plain object, as Office.js `toJSON()` gives them. */
  toJSON(): Record<string, unknown> {
    const out: Record<string, unknown> = {};
    for (const [name, value] of this.#loaded) {
      if (name === '__proto__' || name === 'constructor' || name === 'prototype') continue;
      out[name] = Array.isArray(value) ? value.map(plain) : plain(value);
    }
    return out;
  }

  /** @internal What this kind of object does with a resolved load request. */
  protected abstract onLoad(request: ResolvedLoadOptions): void;

  /** @internal Whether `load(...)` accepts `items/<property>` paths. Only collections do. */
  protected get loadsItemPaths(): boolean {
    return false;
  }

  /** @internal This object's address, or the promise standing in until a sync resolves it. */
  protected get path(): ObjectPath {
    return this.#path;
  }

  /** @internal The owning context's internal surface. */
  protected get internals(): ContextInternals {
    return this.#internals;
  }

  /** @internal Refuse a released or null object, and a finished context. */
  protected requireUsable(): void {
    if (this.#path.isReleased) fail({ code: 'InvalidObjectPath', target: this.#path.label });
    if (this.#path.isNull) {
      const diagnostic = diagnosticOf('null-object');
      fail({
        code: 'InvalidObjectPath',
        target: this.#path.label,
        ...(diagnostic ? { diagnostic } : {}),
      });
    }
    this.#internals.assertUsable(this.#path.label);
  }

  /**
   * @internal Queue one action. `plan` gets the ref of this object (a handle, or the result
   * reference of the action that creates it in the same batch).
   */
  protected enqueue(
    sort: ActionSort,
    name: string,
    plan: (ref: (path: ObjectPath) => AutomationRef) => AutomationOperation,
    settle: (value: AutomationValue) => void = () => {},
    produces?: ObjectPath,
    refable = true,
    continues = false
  ): void {
    this.requireUsable();
    const label = name ? `${this.#path.label}.${name}` : this.#path.label;
    const action = {
      sort,
      label,
      target: this.#path,
      ...(continues ? { continues } : {}),
      plan: (indexOf: IndexOf) => plan((path) => path.ref(indexOf)),
      settle,
    };
    produces?.producedBy(action, refable);
    this.#internals.queue.push(action);
  }

  /** @internal Read a property a completed `load` filled in; refuses if none did. */
  protected loadedProperty<T>(name: string): T {
    if (!this.#loaded.has(name)) {
      if (this.#path.isReleased) fail({ code: 'InvalidObjectPath', target: this.#path.label });
      fail({
        code: 'PropertyNotLoaded',
        target: `${this.#path.label}.${name}`,
        ...(this.#skipped ? { diagnostic: this.#skipped } : {}),
      });
    }
    return this.#loaded.get(name) as T;
  }

  /**
   * @internal A load through a collection skipped this object, for `reason` (a shape with no
   * text frame, a shape that is not a group): a read of what it did not load says why.
   */
  protected markSkipped(reason: PptxEditorDiagnostic | undefined): void {
    this.#skipped = reason;
  }

  /** @internal Whether a completed `load` filled this property in. */
  protected hasLoadedProperty(name: string): boolean {
    return this.#loaded.has(name);
  }

  /** @internal Record a value a completed load produced. */
  protected setLoadedProperty(name: string, value: unknown): void {
    this.#loaded.set(name, value);
  }

  /** @internal The values completed loads filled in, for `context.loadRecursive`. */
  [LOADED_VALUES](): readonly unknown[] {
    return [...this.#loaded.values()];
  }

  /** @internal The path of this object, for the model objects that address another one. */
  get [PATH](): ObjectPath {
    return this.#path;
  }

  /**
   * @internal Queue the refusal of a call whose argument this API does not accept, so the next
   * `sync()` refuses as a whole and applies nothing (`deferred-checks.ts`).
   */
  [DEFER_REFUSAL](error: PptxEditorError): void {
    this.requireUsable();
    this.#internals.queue.push({
      sort: 'write',
      label: error.debugInfo.errorLocation ?? this.#path.label,
      target: this.#path,
      refusal: error,
      plan: () => {
        throw error;
      },
      settle: () => {},
    });
  }

  /** @internal Drop this object's presentation reference; the run that owned it has ended. */
  [RELEASE](): void {
    this.#path.release();
  }

  /** @internal Adopt this object into another run's context. */
  [REBIND](context: RequestContext): void {
    this.#context = context;
    this.#internals = context[INTERNALS];
  }
}
