// The part map of a package that an edit derives from another one.
//
// A transaction changes a few parts of a package that can hold thousands (three or more per
// slide). Copying the whole map for each commit, and for each `can` probe, made one keystroke
// cost time in proportion to the slide count. A derived map instead keeps the base map as its
// root and records only its changes, so a derivation costs the size of its changes. When the
// changes grow past a bound, the next derivation copies everything into a new root, so reads
// stay at a few map lookups.
//
// The map reads exactly as a new `Map` of the base entries, then edited, would: the same
// entries, the same size, and the same iteration order (a root entry keeps its place when it
// is replaced, and an added entry, or a removed root entry that is added again, goes last).

import type { PackagePart } from './package.ts';

/** Changes past this bound make the next derivation copy every entry into a new root. */
const MAX_CHANGES = 512;

/** A part map that records its changes on top of a shared root map. @internal */
export class DerivedParts implements ReadonlyMap<string, PackagePart> {
  readonly #root: ReadonlyMap<string, PackagePart>;
  /** Replaced root entries that keep their place. */
  readonly #over: Map<string, PackagePart>;
  /** Root entries removed (or removed and added again, last). */
  readonly #gone: Set<string>;
  /** Entries after the root's, in insertion order. */
  readonly #tail: Map<string, PackagePart>;
  #sealed = false;
  /** Every entry in order, built on the first pass over a sealed map (`#entryMap`). */
  #flat: Map<string, PackagePart> | null = null;

  private constructor(
    root: ReadonlyMap<string, PackagePart>,
    over: Map<string, PackagePart>,
    gone: Set<string>,
    tail: Map<string, PackagePart>
  ) {
    this.#root = root;
    this.#over = over;
    this.#gone = gone;
    this.#tail = tail;
  }

  /**
   * A writable map with the entries of `base`, for one transaction. It costs the number of
   * changes `base` records, or one copy of `base` when they pass the bound. Call `seal` before
   * the map becomes part of a package.
   */
  static from(base: ReadonlyMap<string, PackagePart>): DerivedParts {
    if (base instanceof DerivedParts && base.changeCount() <= MAX_CHANGES)
      return new DerivedParts(
        base.#root,
        new Map(base.#over),
        new Set(base.#gone),
        new Map(base.#tail)
      );
    return new DerivedParts(
      base instanceof DerivedParts ? (base.#flat ?? new Map(base)) : base,
      new Map(),
      new Set(),
      new Map()
    );
  }

  /** The number of entries that differ from the root. */
  changeCount(): number {
    return this.#over.size + this.#gone.size + this.#tail.size;
  }

  /** End writes: the map is immutable from here on, as every package's part map is. */
  seal(): this {
    this.#sealed = true;
    return this;
  }

  #writable(): void {
    if (this.#sealed) throw new Error('a package part map is immutable');
  }

  set(name: string, part: PackagePart): this {
    this.#writable();
    if (this.#tail.has(name) || this.#gone.has(name) || !this.#root.has(name))
      this.#tail.set(name, part);
    else this.#over.set(name, part);
    return this;
  }

  delete(name: string): boolean {
    this.#writable();
    if (this.#tail.delete(name)) return true;
    if (this.#gone.has(name) || !this.#root.has(name)) return false;
    this.#gone.add(name);
    this.#over.delete(name);
    return true;
  }

  get(name: string): PackagePart | undefined {
    const added = this.#tail.get(name);
    if (added !== undefined) return added;
    if (this.#gone.has(name)) return undefined;
    return this.#over.get(name) ?? this.#root.get(name);
  }

  has(name: string): boolean {
    return this.get(name) !== undefined;
  }

  get size(): number {
    return this.#root.size - this.#gone.size + this.#tail.size;
  }

  /**
   * The entries in one native map, for a pass over every part. A sealed map builds it on its
   * first pass and keeps it, so a pass costs what it cost on a copied map, and a transaction
   * that reads only a few parts builds none.
   */
  #entryMap(): ReadonlyMap<string, PackagePart> {
    if (this.changeCount() === 0) return this.#root;
    if (this.#flat) return this.#flat;
    const flat = new Map<string, PackagePart>();
    for (const [name, part] of this.#root) {
      if (this.#gone.has(name)) continue;
      flat.set(name, this.#over.get(name) ?? part);
    }
    for (const [name, part] of this.#tail) flat.set(name, part);
    if (this.#sealed) this.#flat = flat;
    return flat;
  }

  entries(): MapIterator<[string, PackagePart]> {
    return this.#entryMap().entries();
  }

  keys(): MapIterator<string> {
    return this.#entryMap().keys();
  }

  values(): MapIterator<PackagePart> {
    return this.#entryMap().values();
  }

  [Symbol.iterator](): MapIterator<[string, PackagePart]> {
    return this.entries();
  }

  forEach(
    fn: (part: PackagePart, name: string, map: ReadonlyMap<string, PackagePart>) => void,
    thisArg?: unknown
  ): void {
    for (const [name, part] of this.entries()) fn.call(thisArg, part, name, this);
  }

  get [Symbol.toStringTag](): string {
    return 'Map';
  }

  /**
   * The names whose entries may differ between `a` and `b`, when both are derived from one
   * root (or one is the other's root): a superset of the changed names, at the cost of the
   * changes both record. Undefined when the maps share no root.
   */
  static candidates(
    a: ReadonlyMap<string, PackagePart>,
    b: ReadonlyMap<string, PackagePart>
  ): Iterable<string> | undefined {
    const root = (m: ReadonlyMap<string, PackagePart>) => (m instanceof DerivedParts ? m.#root : m);
    if (root(a) !== root(b)) return undefined;
    const out = new Set<string>();
    for (const m of [a, b]) {
      if (!(m instanceof DerivedParts)) continue;
      for (const name of m.#over.keys()) out.add(name);
      for (const name of m.#gone) out.add(name);
      for (const name of m.#tail.keys()) out.add(name);
    }
    return out;
  }
}
