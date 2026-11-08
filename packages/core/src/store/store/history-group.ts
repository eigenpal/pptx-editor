// History groups and history-free origins of the document store.
//
// The same model as DOCX Editor's `packages/core/src/store/store/history-group.ts` and
// `tree-store.ts`: history is scoped by intent, never by a timer. The caller states the scope.
// A transaction is one entry; consecutive transactions that carry the same group token are one
// entry; a projection or awareness commit is no entry. DOCX Editor commit adf0e178.

/**
 * The identity of one continuous user gesture, for history grouping.
 *
 * A drag, a typing burst, or a live picker applies every intermediate value, so the slide
 * updates under the pointer, and each of those writes is a transaction in its own right: it
 * must render, and a collaborator must receive it. They are not separate undo steps. The caller
 * mints one token when the gesture starts (`Symbol('drag')`) and passes it with every write
 * until the gesture ends.
 *
 * A symbol, not a string, so two unrelated callers cannot collide by naming their gesture the
 * same way, and a token cannot outlive its gesture through serialization.
 *
 * @public
 */
export type HistoryGroup = symbol;

/**
 * What made a transaction.
 *
 * - `user`: a local edit. It records a history entry.
 * - `remote`: ops from a collaboration peer. No history entry; the local undo and redo stacks
 *   are cleared, because undo restores whole snapshots and would revert the peer's edit.
 * - `projection`: the view reconciled with state the store already holds. It publishes a
 *   revision so consumers re-derive, but it is not a user intent: no history entry, and the
 *   undo and redo stacks stay. A collaboration replica does not send it to peers, so commit
 *   with it only what each replica derives on its own; a peer never sees any other change.
 * - `awareness`: presence-style state, handled as `projection`.
 *
 * @public
 */
export type TransactOrigin = 'user' | 'remote' | 'projection' | 'awareness';
