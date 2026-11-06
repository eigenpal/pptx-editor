/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// One point for argument checks: the sync, as in Office.js.
//
// The object model checks its arguments when a call is made. A method call or property write
// that this API refuses for its argument (`InvalidArgument`, `NotImplemented`, `NotSupported`,
// `ItemNotFound`) does not throw at the call: the refusal is queued, and the next `sync()`
// refuses with it as a whole, so nothing of that batch applies and no earlier call stays half
// queued. A method that answers an object answers a stand-in, which takes further calls and
// queues nothing. What still throws at the call: a read before its load (`PropertyNotLoaded`),
// `load(...)` options, and a proxy or context used outside its lifetime.

import { isPptxEditorError, type PptxEditorError, type PptxEditorErrorCode } from './errors.ts';
import { DEFER_REFUSAL } from './internals.ts';

/** The refusals of an argument, which the sync gives. */
const DEFERRED: ReadonlySet<PptxEditorErrorCode> = new Set([
  'InvalidArgument',
  'NotImplemented',
  'NotSupported',
  'ItemNotFound',
]);

/**
 * Members that are not calls of the object model: its machinery, and the calls that keep
 * throwing at the call (`load`, `toJSON`, and the lifetime members).
 */
export const NOT_DEFERRED: ReadonlySet<string> = new Set([
  'constructor',
  'load',
  'toJSON',
  'track',
  'untrack',
  'at',
  'command',
  'deleteLines',
  'enqueue',
  'find',
  'groupOf',
  'handlesOf',
  'hasLoadedProperty',
  'indexArg',
  'itemOf',
  'keyArg',
  'listing',
  'loadBackground',
  'loadCellFormat',
  'loadedProperty',
  'loadScalars',
  'lookup',
  'make',
  'markSkipped',
  'memberPath',
  'navigation',
  'onLoad',
  'requireUsable',
  'selection',
  'setLoadedProperty',
  'takeNestedPaths',
  'writeBackground',
  'writeSize',
]);

/** How deep in wrapped calls the current call is: only the outermost queues a refusal. */
let depth = 0;

/** What a refused call that answers an object answers: it takes any use and queues nothing. */
function standIn(): unknown {
  const proxy: unknown = new Proxy(function standIn() {}, {
    get: (_target, key) => (key === 'then' || typeof key === 'symbol' ? undefined : proxy),
    set: () => true,
    apply: () => proxy,
  });
  return proxy;
}

/** What a wrapped call needs from its object: a way to queue the refusal. */
interface Deferring {
  [DEFER_REFUSAL](error: PptxEditorError): void;
}

function deferring<A extends unknown[], R>(
  call: (this: Deferring, ...args: A) => R,
  answers: boolean
): (this: Deferring, ...args: A) => R {
  return function (this: Deferring, ...args: A): R {
    depth += 1;
    try {
      return call.apply(this, args);
    } catch (error) {
      if (depth > 1 || !isPptxEditorError(error) || !DEFERRED.has(error.code)) throw error;
      this[DEFER_REFUSAL](error);
      return (answers ? standIn() : undefined) as R;
    } finally {
      depth -= 1;
    }
  };
}

const done = new WeakSet<object>();

/** The member names that `deferArgumentChecks` wrapped, for the check against Office.js. */
export const DEFERRED_MEMBERS = new Set<string>();

/**
 * Move the argument checks of the calls and property writes of `proto` and the prototypes above
 * it, up to `stop`, to the sync. Each prototype is changed once. `ClientObject` calls this for
 * each object it constructs, so every class of the object model follows the rule.
 */
export function deferArgumentChecks(proto: object | null, stop: object): void {
  while (proto && proto !== stop && !done.has(proto)) {
    done.add(proto);
    for (const name of Object.getOwnPropertyNames(proto)) {
      if (NOT_DEFERRED.has(name)) continue;
      const d = Object.getOwnPropertyDescriptor(proto, name)!;
      if (typeof d.value === 'function') {
        Object.defineProperty(proto, name, { ...d, value: deferring(d.value, true) });
        DEFERRED_MEMBERS.add(name);
      } else if (d.set) {
        Object.defineProperty(proto, name, { ...d, set: deferring(d.set, false) });
        DEFERRED_MEMBERS.add(name);
      }
    }
    proto = Object.getPrototypeOf(proto) as object | null;
  }
}
