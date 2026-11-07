// A bounded least-recently-used cache for measured advances, in two generations: new entries
// go into the young map, a hit in the old map moves the entry back to the young one, and when
// the young map is full it becomes the old map and the previous old map is dropped. The young
// map is full at `limit / 2` entries or at `weightLimit / 2` weight. An entry used within the
// last `limit / 2` insertions and the last `weightLimit / 2` of inserted weight is never
// dropped, so the text being edited stays measured while a large deck's other runs cycle
// through. Every operation is O(1).
//
// An entry's weight is the caller's size of it. A run's advances grow with its text, and every
// keystroke in a long paragraph adds a new entry: a count limit alone keeps that many copies of
// the paragraph.
//
// A single Map that deletes its oldest entry is not used: in JavaScriptCore, deleting from the
// front of a Map and asking for its first key again costs time in proportion to the deleted
// entries, about 15 µs per insertion at this size.

interface Entry<V> {
  readonly value: V;
  readonly weight: number;
}

/**
 * A two-generation LRU cache holding at most `limit` entries and at most `weightLimit` weight.
 * Internal to the layout lane.
 */
export class AdvanceCache<V> {
  #young = new Map<string, Entry<V>>();
  #old = new Map<string, Entry<V>>();
  #youngWeight = 0;
  #oldWeight = 0;
  readonly #generation: number;
  readonly #generationWeight: number;

  constructor(limit: number, weightLimit = Infinity) {
    this.#generation = Math.max(1, Math.floor(limit / 2));
    this.#generationWeight = weightLimit / 2;
  }

  get(key: string): V | undefined {
    const young = this.#young.get(key);
    if (young !== undefined) return young.value;
    const old = this.#old.get(key);
    if (old !== undefined) {
      this.#old.delete(key);
      this.#oldWeight -= old.weight;
      this.#add(key, old);
    }
    return old?.value;
  }

  /** Keep `value` under `key`; `weight` is its size in the unit of `weightLimit`. */
  set(key: string, value: V, weight = 0): void {
    // An entry larger than a generation would empty the cache for itself alone.
    if (weight > this.#generationWeight) return;
    const young = this.#young.get(key);
    if (young) {
      this.#young.delete(key);
      this.#youngWeight -= young.weight;
    }
    this.#add(key, { value, weight });
  }

  /** Entries held now. */
  get size(): number {
    return this.#young.size + this.#old.size;
  }

  /** The weight held now. */
  get weight(): number {
    return this.#youngWeight + this.#oldWeight;
  }

  #add(key: string, entry: Entry<V>): void {
    if (
      this.#young.size >= this.#generation ||
      this.#youngWeight + entry.weight > this.#generationWeight
    ) {
      this.#old = this.#young;
      this.#oldWeight = this.#youngWeight;
      this.#young = new Map();
      this.#youngWeight = 0;
    }
    this.#young.set(key, entry);
    this.#youngWeight += entry.weight;
  }
}
