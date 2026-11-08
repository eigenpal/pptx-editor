// The change a store delivers to its listeners now. Code that runs inside a listener (the
// editor's surface, during an IME composition) reads it to tell a peer's transaction (origin
// `remote`) from a local one. Internal: the store index does not export it.

import type { StoreChange } from './document-store.ts';

const current = new WeakMap<object, StoreChange>();

/** Record that `store` delivers `change` now; `undefined` when the delivery ends. */
export function setDelivering(store: object, change: StoreChange | undefined): void {
  if (change) current.set(store, change);
  else current.delete(store);
}

/** The change `store` delivers now, or undefined outside a delivery. */
export function deliveringChange(store: object): StoreChange | undefined {
  return current.get(store);
}
