// How the editor tells its `subscribe` listeners about a published state. Listeners hear of it
// synchronously, except during `deferListeners`: a document refresh commits a large difference
// there, and the adapters render the whole chrome when they hear of it. The editor paints its
// own slide in the commit's task; its listeners hear of the commit in the next task, so the
// commit and the chrome render are two tasks, not one long task.
//
// A deferred notification keeps the cause of the change (`withChangeSource`), so its `change`
// event reports it. A notification that is due while one waits delivers the waiting one first:
// listeners see every state, in order.

import type { DocumentChange } from '../contracts/index.ts';
import { changeSourceOf, withChangeSource } from './change-source.ts';

const deferring = new WeakMap<object, number>();
/** Editors in `deferListeners`, so an editor that defers nothing reads no owner. */
let deferringCount = 0;

/**
 * Run `run`; the `subscribe` listeners of `editor` hear of the states it publishes in the next
 * task, once. @internal
 */
export function deferListeners<T>(editor: object, run: () => T): T {
  deferring.set(editor, (deferring.get(editor) ?? 0) + 1);
  deferringCount += 1;
  try {
    return run();
  } finally {
    deferringCount -= 1;
    const depth = (deferring.get(editor) ?? 1) - 1;
    if (depth > 0) deferring.set(editor, depth);
    else deferring.delete(editor);
  }
}

/** The notifier of one editor's listeners. @internal */
export interface ListenerRelay {
  /** Add a listener. Returns the function that removes it. */
  add(listener: () => void): () => void;
  /** Tell the listeners, now or, during `deferListeners`, in the next task. */
  notify(): void;
  /** Drop the listeners and a waiting notification. */
  clear(): void;
}

/**
 * A relay with no listeners. `owner` returns the editor instance; it is read only while some
 * editor defers, so a call before the instance exists is safe. @internal
 */
export function createListenerRelay(owner: () => object): ListenerRelay {
  const listeners = new Set<() => void>();
  let waiting: {
    readonly timer: ReturnType<typeof setTimeout>;
    readonly source: DocumentChange['source'];
  } | null = null;
  const tell = (): void => {
    for (const listener of [...listeners]) listener();
  };
  const flush = (): void => {
    if (!waiting) return;
    const { timer, source } = waiting;
    waiting = null;
    clearTimeout(timer);
    if (source) withChangeSource(owner(), source, tell);
    else tell();
  };
  return {
    add(listener) {
      listeners.add(listener);
      return () => void listeners.delete(listener);
    },
    notify() {
      if (deferringCount > 0 && deferring.has(owner())) {
        // One notification for the commit; it reads the state when it runs.
        if (!waiting) waiting = { timer: setTimeout(flush, 0), source: changeSourceOf(owner()) };
        return;
      }
      flush();
      tell();
    },
    clear() {
      if (waiting) clearTimeout(waiting.timer);
      waiting = null;
      listeners.clear();
    },
  };
}
