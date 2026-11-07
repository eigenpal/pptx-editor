// The id → location index of a part, cached per root identity.
//
// Internal to the store lane: `tree-edit.ts` exposes `nodeIndex`, and a transaction
// (`working-tree.ts`) hands over the index it kept in step with its edits, so the next reader of
// the committed root does not walk the whole tree again.

import type { Element, Node, NodeId, XmlPart } from '../package/tree.ts';

/** Where a node sits: its parent id (null for the root) and index among the parent's children. */
export interface Location {
  readonly node: Node;
  readonly parent: NodeId | null;
  readonly index: number;
}

const indexCache = new WeakMap<Element, ReadonlyMap<NodeId, Location>>();

// Work counters for tests, as DOCX Editor's `nodeIndexTestRecorder` (store/package/ooxml-edit.ts):
// a count is exact on every machine, a wall-clock limit is not.
let nodeIndexCompleteBuilds = 0;
let nodeIndexCompleteVisits = 0;

/** @internal Counts the complete index builds of `nodeIndex` and the nodes they visit. */
export function nodeIndexTestRecorder(): {
  readonly completeBuilds: number;
  readonly completeVisits: number;
  reset(): void;
} {
  return {
    get completeBuilds() {
      return nodeIndexCompleteBuilds;
    },
    get completeVisits() {
      return nodeIndexCompleteVisits;
    },
    reset() {
      nodeIndexCompleteBuilds = 0;
      nodeIndexCompleteVisits = 0;
    },
  };
}

/** The id → location index of a part, built once per root identity. */
export function nodeIndex(part: XmlPart): ReadonlyMap<NodeId, Location> {
  const hit = indexCache.get(part.root);
  if (hit) return hit;
  const map = new Map<NodeId, Location>();
  const stack: Location[] = [{ node: part.root, parent: null, index: 0 }];
  while (stack.length > 0) {
    const item = stack.pop()!;
    nodeIndexCompleteVisits += 1;
    map.set(item.node.id, item);
    if (item.node.type === 'element') {
      const kids = item.node.children;
      for (let i = kids.length - 1; i >= 0; i -= 1)
        stack.push({ node: kids[i], parent: item.node.id, index: i });
    }
  }
  nodeIndexCompleteBuilds += 1;
  indexCache.set(part.root, map);
  return map;
}

/** How many layers an index may stack before a commit flattens it into one map. */
const MAX_LAYERS = 8;

/**
 * An index as the changes of one transaction over the index of the root before it. A command
 * that commits many small edits to a large part (Select All, then a format) then costs what its
 * edits touch, not the size of the part for each commit. A read walks at most `MAX_LAYERS`
 * layers; iteration flattens the layers once.
 */
class LayeredIndex implements ReadonlyMap<NodeId, Location> {
  #flat: Map<NodeId, Location> | null = null;

  /** The size of the bottom map: the scale that decides when the changes are large. */
  readonly scale: number;

  constructor(
    readonly base: ReadonlyMap<NodeId, Location>,
    readonly changes: ReadonlyMap<NodeId, Location | undefined>,
    readonly depth: number
  ) {
    this.scale = base instanceof LayeredIndex ? base.scale : base.size;
  }

  get(id: NodeId): Location | undefined {
    return this.changes.has(id) ? this.changes.get(id) : this.base.get(id);
  }

  has(id: NodeId): boolean {
    return this.get(id) !== undefined;
  }

  /** The layers as one map, built once. */
  flat(): Map<NodeId, Location> {
    if (this.#flat) return this.#flat;
    // Copy the bottom map once, then apply each layer's changes from the bottom up.
    const layers: LayeredIndex[] = [];
    let bottom: ReadonlyMap<NodeId, Location> = this;
    while (bottom instanceof LayeredIndex) {
      if (bottom.#flat) break;
      layers.push(bottom);
      bottom = bottom.base;
    }
    const map = new Map(bottom instanceof LayeredIndex ? bottom.#flat! : bottom);
    for (let i = layers.length - 1; i >= 0; i -= 1)
      for (const [id, loc] of layers[i]!.changes) {
        if (loc) map.set(id, loc);
        else map.delete(id);
      }
    return (this.#flat = map);
  }

  get size(): number {
    return this.flat().size;
  }
  keys() {
    return this.flat().keys();
  }
  values() {
    return this.flat().values();
  }
  entries() {
    return this.flat().entries();
  }
  forEach(f: (value: Location, key: NodeId, map: ReadonlyMap<NodeId, Location>) => void): void {
    this.flat().forEach((v, k) => f(v, k, this));
  }
  [Symbol.iterator]() {
    return this.flat()[Symbol.iterator]();
  }
}

/**
 * The index after a transaction: `changes` (an entry, or undefined for a removed node) over
 * `base`, the index of the root the transaction opened. Layered while the stack is short and
 * the changes are small, else one map. The caller gives up `changes`. @internal
 */
export function layeredNodeIndex(
  base: ReadonlyMap<NodeId, Location>,
  changes: Map<NodeId, Location | undefined>
): ReadonlyMap<NodeId, Location> {
  const depth = base instanceof LayeredIndex ? base.depth + 1 : 1;
  const layered = new LayeredIndex(base, changes, depth);
  return depth > MAX_LAYERS || changes.size * 4 > layered.scale ? flatten(layered) : layered;
}

function flatten(index: LayeredIndex): Map<NodeId, Location> {
  const map = index.flat();
  nodeIndexFlattenWrites += map.size;
  return map;
}

let nodeIndexFlattenWrites = 0;

/** @internal The entries that flattened layered indexes wrote, for the work tests. */
export function nodeIndexFlattenCount(): number {
  return nodeIndexFlattenWrites;
}

/**
 * Record an index built for `root` elsewhere. The caller gives up the map: nothing may change
 * it afterwards. Ignored when the root already has one.
 */
export function adoptNodeIndex(root: Element, index: ReadonlyMap<NodeId, Location>): void {
  if (!indexCache.has(root)) indexCache.set(root, index);
}
