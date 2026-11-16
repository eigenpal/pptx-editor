// What caused a `change` event (`DocumentChange.source`), after DOCX Editor's
// `refresh-edit-source.ts` (docx-editor/packages/core/src/editor at b63cbea4a).
//
// The editor publishes every change through `subscribe`, synchronously. A replacement runs its
// commit inside `withChangeSource`, so the event that the commit fires reads the source. A load
// gives the document a new load id: the first `change` event with a new load id has
// `source: 'load'`, unless a refresh named the load.

import type { DocumentChange } from '../contracts/index.ts';

const active = new WeakMap<object, NonNullable<DocumentChange['source']>>();
/** The cause of the load that replaces load id `from`: a refresh that reloads with `load()`. */
const nextLoad = new WeakMap<
  object,
  { readonly from: string | undefined; readonly source: NonNullable<DocumentChange['source']> }
>();

/** Name the cause of the editor's next load. @internal */
export function expectLoad(editor: object, source: NonNullable<DocumentChange['source']>): void {
  nextLoad.set(editor, { from: loadIdOf(editor), source });
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

/** The editor's handle, when it has one. */
function loadIdOf(editor: object): string | undefined {
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
    const loaded = now !== loadId;
    const named = nextLoad.get(editor);
    const expected = loaded && named && named.from === loadId ? named.source : undefined;
    loadId = now;
    const source = active.get(editor) ?? (loaded ? (expected ?? 'load') : undefined);
    return source ? { source, revision } : { revision };
  };
}
