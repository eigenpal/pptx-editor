/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// After DOCX Editor's `packages/editor-api/src/runtime/object-path.ts` (commit 5993bd223945).
//
// What the host is told to look at. A proxy's path is pending until the read that names it
// answers, then resolved to a handle, null (a `getItemOrNullObject` that found nothing), or
// released (its run ended and nothing tracked it).
//
// ONE BATCH, NOT ROUNDS. The presentation host takes a result reference (`{ result: i }`) for
// the handle operation `i` of the same batch answers. So a pending path whose producing action
// is in the batch being planned is addressed by that reference, and a consumer can create a
// shape and write its text in one `sync()`. A pending path whose producer is not in the batch
// (an earlier sync failed) cannot be addressed, and the plan refuses with `InvalidObjectPath`.

import type { AutomationHandle, AutomationRef } from '@pptx-editor.dev/core/automation';
import { diagnosticOf, fail } from './errors.ts';

/** Refuse a path that cannot be addressed, with the reason when it is a null object. */
function unaddressable(state: ObjectPathState, target: string, refused = false): never {
  const diagnostic =
    state.status === 'null'
      ? diagnosticOf('null-object')
      : refused
        ? diagnosticOf('refused-sync')
        : undefined;
  fail({ code: 'InvalidObjectPath', target, ...(diagnostic ? { diagnostic } : {}) });
}

export type ObjectPathState =
  /** Promised: created by a queued action that has not answered yet. */
  | { readonly status: 'pending' }
  /** Addressable: the host has named this object. */
  | { readonly status: 'resolved'; readonly handle: AutomationHandle }
  /** A `get…OrNullObject` that found nothing. Not an error, and never addressable. */
  | { readonly status: 'null' }
  /** Its run ended without tracking it. Terminal. */
  | { readonly status: 'released' };

/** The batch index of an action, while a batch is planned. */
export type IndexOf = (producer: object) => number | undefined;

/**
 * Thrown while planning: the object is named by a read in this batch whose answer is not a
 * handle (the reader's selection, a null-object check). The sync sends that read first, in an
 * earlier round at the same revision, and plans the call again.
 */
export class WaitForRead {
  constructor(readonly producer: object) {}
}

/**
 * The lookups of a sync that was refused or dropped. A proxy is a path, as in Office.js: when a
 * later batch names an object whose lookup was in that sync, the batch sends the lookup again
 * (`NeedsProducer`), so the same proxy works in a retry while the object exists.
 */
const replayable = new WeakSet<object>();

/** Mark a read action of a refused sync, so a later batch that needs its answer resends it. */
export function markReplayable(action: object): void {
  replayable.add(action);
}

/**
 * The calls of a refused sync that would have created an object (`addTextBox`, `add`). Their
 * objects never existed, so a later call on one refuses with its own diagnostic.
 */
const refusedCreators = new WeakSet<object>();

/** Mark a write action of a refused sync: the object it would have created does not exist. */
export function markRefused(action: object): void {
  refusedCreators.add(action);
}

/** Thrown while planning: the object's lookup was in a refused sync. Send it again first. */
export class NeedsProducer {
  constructor(readonly producer: object) {}
}

export class ObjectPath {
  readonly label: string;
  #state: ObjectPathState;
  /**
   * The path whose fate this one shares. `shape.fill` is addressable exactly when the shape is,
   * and released when the shape is.
   */
  readonly #parent: ObjectPath | null;
  /** The queued action whose answer names this object, while it is pending. */
  #producer: object | null = null;
  /** Whether the producer's answer is this object's handle, so a result reference names it. */
  #refable = true;

  private constructor(label: string, state: ObjectPathState, parent: ObjectPath | null = null) {
    this.label = label;
    this.#state = state;
    this.#parent = parent;
  }

  /** A path that is addressable from the moment it exists: a root, or an item just hydrated. */
  static of(label: string, handle: AutomationHandle): ObjectPath {
    return new ObjectPath(label, { status: 'resolved', handle });
  }

  /** A path a queued action will fill in, or mark null. */
  static pending(label: string): ObjectPath {
    return new ObjectPath(label, { status: 'pending' });
  }

  /**
   * A pending path that the same queued action as this one settles, for an object that its
   * answer names or finds absent. A call in the same batch waits for that read.
   */
  sibling(label: string): ObjectPath {
    const path = new ObjectPath(label, { status: 'pending' });
    const origin = this.origin;
    if (origin.#producer) {
      path.#producer = origin.#producer;
      path.#refable = false;
    }
    return path;
  }

  /** A path that is whatever its owner's path is, under its own name. */
  static derived(label: string, parent: ObjectPath): ObjectPath {
    return new ObjectPath(label, { status: 'pending' }, parent);
  }

  get state(): ObjectPathState {
    return this.#parent ? this.#parent.state : this.#state;
  }

  /** The path this one shares its fate with, followed to the end. */
  get origin(): ObjectPath {
    return this.#parent ? this.#parent.origin : this;
  }

  get isAddressable(): boolean {
    return this.state.status === 'resolved';
  }

  get isPending(): boolean {
    return this.state.status === 'pending';
  }

  get isNull(): boolean {
    return this.state.status === 'null';
  }

  get isReleased(): boolean {
    return this.state.status === 'released';
  }

  /**
   * Name the action whose answer resolves this path. `refable` is false when the answer is not
   * the handle itself, so a call in the same batch waits for that read.
   */
  producedBy(action: object, refable = true): void {
    this.origin.#producer = action;
    this.origin.#refable = refable;
  }

  /**
   * What to put in an operation: the handle, or the result reference of the producing action
   * when it is in the batch being planned. Anything else refuses with `InvalidObjectPath`.
   */
  ref(indexOf: IndexOf): AutomationRef {
    const state = this.state;
    if (state.status === 'resolved') return state.handle;
    if (state.status === 'pending') {
      const producer = this.origin.#producer;
      const index = producer ? indexOf(producer) : undefined;
      if (index !== undefined) {
        if (this.origin.#refable) return { result: index };
        throw new WaitForRead(producer!);
      }
      if (producer && replayable.has(producer)) throw new NeedsProducer(producer);
      if (producer && refusedCreators.has(producer)) unaddressable(state, this.label, true);
    }
    unaddressable(state, this.label);
  }

  /** The handle, once resolved. */
  handle(): AutomationHandle {
    const state = this.state;
    if (state.status !== 'resolved') unaddressable(state, this.label);
    return state.handle;
  }

  /** Hydration: the read answered, and this is the object it named. A released path stays so. */
  resolveTo(handle: AutomationHandle): void {
    if (this.#parent || this.#state.status === 'released') return;
    this.#state = { status: 'resolved', handle };
    this.#producer = null;
  }

  /** Hydration: the read answered, and there was nothing there. */
  resolveNull(): void {
    if (this.#parent || this.#state.status === 'released') return;
    this.#state = { status: 'null' };
    this.#producer = null;
  }

  /** The run ended and nothing kept this object alive. A derived path follows its owner. */
  release(): void {
    if (this.#parent) return;
    this.#state = { status: 'released' };
    this.#producer = null;
  }
}
