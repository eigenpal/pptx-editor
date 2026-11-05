// A bounded least-recently-used cache for measured advances, in two generations: new entries
// go into the young map, a hit in the old map moves the entry back to the young one, and when
// the young map is full it becomes the old map and the previous old map is dropped. An entry
// used within the last `limit / 2` insertions is never dropped, so the text being edited stays
// measured while a large deck's other runs cycle through. Every operation is O(1).
//
// A single Map that deletes its oldest entry is not used: in JavaScriptCore, deleting from the
// front of a Map and asking for its first key again costs time in proportion to the deleted
// entries, about 15 µs per insertion at this size.

/** A two-generation LRU cache holding at most `limit` entries. Internal to the layout lane. */
export class AdvanceCache<V> {
  #young = new Map<string, V>();
  #old = new Map<string, V>();
  readonly #generation: number;

  constructor(limit: number) {
    this.#generation = Math.max(1, Math.floor(limit / 2));
  }

  get(key: string): V | undefined {
    const young = this.#young.get(key);
    if (young !== undefined) return young;
    const old = this.#old.get(key);
    if (old !== undefined) {
      this.#old.delete(key);
      this.#add(key, old);
    }
    return old;
  }

  set(key: string, value: V): void {
    this.#add(key, value);
  }

  /** Entries held now. */
  get size(): number {
    return this.#young.size + this.#old.size;
  }

  #add(key: string, value: V): void {
    if (this.#young.size >= this.#generation) {
      this.#old = this.#young;
      this.#young = new Map();
    }
    this.#young.set(key, value);
  }
}
