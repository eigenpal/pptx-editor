// What caused a `change` event (`DocumentChange.source`), after DOCX Editor's
// `refresh-edit-source.ts` (docx-editor/packages/core/src/editor at b63cbea4a).
//
// The editor publishes every change through `subscribe`, synchronously. A replacement runs its
// commit inside `withChangeSource`, so the event that the commit fires reads the source. A load
// gives the document a new load id: the first `change` event with a new load id has
// `source: 'load'`, unless a refresh named the load. A reload in place gives a new load id
// too, but its commit reports the refresh itself (`adoptLoadId`): the next edit has no source.

import type { DocumentChange } from '../contracts/index.ts';

const active = new WeakMap<object, NonNullable<DocumentChange['source']>>();
/** The cause of the load that replaces load id `from`: a refresh that reloads with `load()`. */
const nextLoad = new WeakMap<
  object,
  { readonly from: string | undefined; readonly source: NonNullable<DocumentChange['source']> }
>();

/** Load ids that a reload in place gave, per editor: no `change` event reports them as a load. */
const adopted = new WeakMap<object, Set<string>>();

/**
 * Record that a reload in place gave the document load id `loadId`. Its commit reported the
 * cause with `withChangeSource`, so a change after it is no load. @internal
 */
export function adoptLoadId(editor: object, loadId: string): void {
  let ids = adopted.get(editor);
  if (!ids) adopted.set(editor, (ids = new Set()));
  ids.add(loadId);
}

/** Name the cause of the editor's next load. @internal */
export function expectLoad(editor: object, source: NonNullable<DocumentChange['source']>): void {
  nextLoad.set(editor, { from: loadIdOf(editor), source });
}

/**
 * Drop the cause that `expectLoad` named: the load ended. A load that failed, or that another
 * load replaced, leaves the load id as it was, so the next load must not read that cause.
 * @internal
 */
export function forgetLoad(editor: object): void {
  nextLoad.delete(editor);
}

/** Run `run` with `source` as the cause of the changes it publishes. @internal */
export function withChangeSource<T>(
  editor: object,
  source: NonNullable<DocumentChange['source']>,
  run: () => T
): T {
  const before = active.get(editor);
  active.set(editor, source);
  try {
    return run();
  } finally {
    if (before) active.set(editor, before);
    else active.delete(editor);
  }
}

/** The load id of the editor's handle, when it has one. @internal */
export function loadIdOf(editor: object): string | undefined {
  const handle = (editor as { getDocumentHandle?: () => { loadId: string } }).getDocumentHandle;
  return typeof handle === 'function' ? handle.call(editor).loadId : undefined;
}

/**
 * A tracker of the cause of each `change` event of one subscriber: it remembers the load id the
 * subscriber saw last. @internal
 */
export function changeSourceTracker(editor: object): (revision: number) => DocumentChange {
  let loadId = loadIdOf(editor);
  return (revision) => {
    const now = loadIdOf(editor);
    const loaded = now !== loadId && !(now !== undefined && adopted.get(editor)?.has(now));
    const named = nextLoad.get(editor);
    const expected = loaded && named && named.from === loadId ? named.source : undefined;
    loadId = now;
    const source = active.get(editor) ?? (loaded ? (expected ?? 'load') : undefined);
    return source ? { source, revision } : { revision };
  };
}
