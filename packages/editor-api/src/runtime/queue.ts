/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// After DOCX Editor's `packages/editor-api/src/runtime/queue.ts` (commit 5993bd223945).
//
// The action queue: what a proxy call leaves behind instead of doing something. A queued action
// is how to say it to the host (`plan`) and what to do with the answer (`settle`), plus the label
// an error may name. Hydration is positional: the answer at index i belongs to the action at i.
//
// `take()` REMOVES the actions and nothing puts them back, including when the batch fails. A
// failed sync loses its batch and says so with a typed error; what is queued afterwards is a new
// batch that can succeed on its own.

import type { AutomationOperation, AutomationValue } from '@pptx-editor.dev/core/automation';
import type { IndexOf, ObjectPath } from './object-path.ts';

/** Whether an action reads or writes. Drives the conditional-revision rule in `sync()`. */
export type ActionSort = 'read' | 'write';

export interface QueuedAction {
  readonly sort: ActionSort;
  /** The consumer-facing name of what this action is for, for errors. Never a handle. */
  readonly label: string;
  /**
   * The object the action is on. A read on an object that an earlier round of the same sync
   * found null is dropped: a null object has nothing to read.
   */
  readonly target?: ObjectPath;
  /**
   * True for a listing whose answer queues more reads (`load('items/<name>')`). In a sync that
   * writes, such a listing ahead of every write goes in an earlier round, with its item reads,
   * so the writes stay last and atomic.
   */
  readonly continues?: boolean;
  /**
   * The host operation for this action, called once at dispatch. `indexOf` gives the batch
   * index of a producing action, for result references. May throw `InvalidObjectPath`, which
   * refuses the batch before anything is sent.
   */
  plan(indexOf: IndexOf): AutomationOperation;
  /**
   * A call whose argument this API refused. A sync that holds one refuses with it before it
   * sends anything, so the whole batch applies nothing.
   */
  readonly refusal?: Error;
  /** The host's answer for this action, in batch order. */
  settle(value: AutomationValue): void;
}

export class ActionQueue {
  #actions: QueuedAction[] = [];
  /** Where `push` puts an action while `continuing` runs. */
  #sink: QueuedAction[] | null = null;
  /** Reads queued while a sync settled an answer, sent by that same sync. */
  #continuations: QueuedAction[] = [];

  get size(): number {
    return this.#actions.length;
  }

  /** The labels of the actions queued now, in order, for `context.debugInfo`. */
  labels(): string[] {
    return this.#actions.map((action) => action.label);
  }

  push(action: QueuedAction): void {
    (this.#sink ?? this.#actions).push(action);
  }

  /**
   * Run `work` while a sync settles an answer, and keep what it queued for that same sync. This
   * is how `load('items/name')` reads each item: the items are known only once the listing has
   * answered.
   */
  continuing(work: () => void): void {
    const previous = this.#sink;
    const sink: QueuedAction[] = [];
    this.#sink = sink;
    try {
      work();
    } finally {
      this.#sink = previous;
    }
    for (const action of sink) this.#continuations.push(action);
  }

  /**
   * Remember where the queue is now. The function returned drops what was queued since, so a
   * `load(...)` that refuses part way leaves nothing behind.
   */
  checkpoint(): () => void {
    const target = this.#sink ?? this.#actions;
    const length = target.length;
    return () => {
      target.length = Math.min(target.length, length);
    };
  }

  /** Hand over the reads `continuing` kept, and forget them. */
  takeContinuations(): readonly QueuedAction[] {
    const taken = this.#continuations;
    this.#continuations = [];
    return taken;
  }

  /** Hand over everything queued and forget it. Never replayed. */
  take(): readonly QueuedAction[] {
    const taken = this.#actions;
    this.#actions = [];
    return taken;
  }

  /** Drop everything queued: what a finished run does with actions nobody synced. */
  clear(): void {
    this.#actions = [];
    this.#continuations = [];
  }
}
