// One part's tree during a transaction.
//
// A transaction can apply thousands of ops to one slide. Rebuilding the id index after each op
// walks the whole slide every time (2.6 s for 1,000 ops on a 60,000-node slide). This keeps ONE
// index per part for the whole transaction and updates only what an edit touches: the path to
// the root, the subtree that came or went, and the siblings that moved. When the transaction
// commits, the index is handed to the cache for the new root.
//
// Edits are persistent: every untouched subtree is shared with the previous revision.

import type { Element, Node, NodeId, NodeKind, XmlPart } from '../package/tree.ts';
import { deriveKind, isTransparentWrapper } from '../package/node-kinds.ts';
import { adoptNodeIndex, nodeIndex, type Location } from './node-index.ts';

// Work counters for tests, beside `nodeIndexTestRecorder`: a transaction copies the index once
// per part it opens, and its edits then write only the entries they touch. The kept index is a
// `KeptIndex`, so every write to it is counted, whichever code makes it.
let workingTreeOpens = 0;
let workingTreeSubtreeVisits = 0;
let workingTreeIndexWrites = 0;

/**
 * @internal Counts the working trees that transactions open, the nodes that edits add to or
 * remove from a kept index, and every entry written to or removed from a kept index (the copy a
 * working tree makes when it opens writes each entry once).
 */
export function workingTreeTestRecorder(): {
  readonly opens: number;
  readonly subtreeVisits: number;
  readonly indexWrites: number;
  reset(): void;
} {
  return {
    get opens() {
      return workingTreeOpens;
    },
    get subtreeVisits() {
      return workingTreeSubtreeVisits;
    },
    get indexWrites() {
      return workingTreeIndexWrites;
    },
    reset() {
      workingTreeOpens = 0;
      workingTreeSubtreeVisits = 0;
      workingTreeIndexWrites = 0;
    },
  };
}

/**
 * A transaction's id index: a plain map behind methods that count each entry written to it or
 * removed from it. It wraps the map and does not extend `Map`: a subclass that overrides `set`
 * loses the engine's fast path for `new Map(map)` (in JavaScriptCore, about 25 times slower
 * on a 60,000-node slide). Only `release()` hands the map out, when the tree is done and no
 * longer writes to it, so every write before that is counted.
 */
class KeptIndex {
  readonly #map: Map<NodeId, Location>;

  constructor(source: ReadonlyMap<NodeId, Location>) {
    this.#map = new Map(source);
    workingTreeIndexWrites += this.#map.size;
  }

  get(id: NodeId): Location | undefined {
    return this.#map.get(id);
  }

  set(id: NodeId, loc: Location): void {
    workingTreeIndexWrites += 1;
    this.#map.set(id, loc);
  }

  delete(id: NodeId): void {
    workingTreeIndexWrites += 1;
    this.#map.delete(id);
  }

  /** The map itself, for the cache. The caller gives up this index: nothing may change it. */
  release(): ReadonlyMap<NodeId, Location> {
    return this.#map;
  }
}

export class WorkingTree {
  readonly partName: string;
  readonly #source: XmlPart;
  readonly #index: KeptIndex;
  #root: Element;
  #nextId: number;

  constructor(part: XmlPart) {
    this.partName = part.partName;
    this.#source = part;
    this.#root = part.root;
    this.#nextId = part.nextId;
    // A copy: the cached index belongs to the previous root, which history still holds.
    this.#index = new KeptIndex(nodeIndex(part));
    workingTreeOpens += 1;
  }

  /** Where a node sits now, or undefined. */
  get(id: NodeId): Location | undefined {
    return this.#index.get(id);
  }

  /** The element with `id`, or undefined when there is no such element. */
  element(id: NodeId): Element | undefined {
    const node = this.#index.get(id)?.node;
    return node?.type === 'element' ? node : undefined;
  }

  /** The id counter for new nodes. */
  get nextId(): number {
    return this.#nextId;
  }

  set nextId(value: number) {
    this.#nextId = value;
  }

  /** Mint one new node id. */
  mintId(): NodeId {
    return `n${this.#nextId++}`;
  }

  /** The nesting level of a node: the root element is level 1. */
  level(id: NodeId): number {
    let level = 1;
    for (let p = this.#index.get(id)?.parent ?? null; p !== null; p = this.#index.get(p)!.parent)
      level += 1;
    return level;
  }

  /** Whether `ancestor` is `id` or one of its ancestors. */
  isSelfOrAncestor(ancestor: NodeId, id: NodeId): boolean {
    for (let cur: NodeId | null = id; cur !== null; cur = this.#index.get(cur)?.parent ?? null)
      if (cur === ancestor) return true;
    return false;
  }

  /**
   * The kind context children of `id` are typed under: the element's kind, or for a
   * markup-compatibility wrapper, the context of its own parent.
   */
  childContext(id: NodeId): NodeKind | null {
    let cur: NodeId | null = id;
    while (cur !== null) {
      const loc = this.#index.get(cur);
      if (!loc || loc.node.type !== 'element') return null;
      if (!isTransparentWrapper(loc.node.ns, loc.node.local)) return loc.node.kind;
      cur = loc.parent;
    }
    return null;
  }

  /** Replace an element by a copy with the same id and the same children (an attribute edit). */
  replace(id: NodeId, next: Element): void {
    this.#commit(id, next);
  }

  /** Replace all children of an element. */
  replaceChildren(id: NodeId, kids: readonly Node[]): void {
    const el = this.element(id)!;
    for (const old of el.children) this.#forget(old);
    kids.forEach((kid, i) => this.#learn(kid, id, i));
    this.#commit(id, { ...el, children: kids });
  }

  /** Insert `node` (not yet in the tree) as child `index` of `parentId`. */
  insertChild(parentId: NodeId, index: number, node: Node): void {
    const el = this.element(parentId)!;
    const kids = el.children.slice();
    kids.splice(index, 0, node);
    for (let i = index + 1; i < kids.length; i += 1)
      this.#index.set(kids[i].id, { node: kids[i], parent: parentId, index: i });
    this.#learn(node, parentId, index);
    this.#commit(parentId, { ...el, children: kids });
  }

  /** Remove the node `id` (not the root) with its subtree. */
  remove(id: NodeId): void {
    const loc = this.#index.get(id)!;
    const parentId = loc.parent!;
    const el = this.element(parentId)!;
    const kids = el.children.slice();
    kids.splice(loc.index, 1);
    this.#forget(loc.node);
    for (let i = loc.index; i < kids.length; i += 1)
      this.#index.set(kids[i].id, { node: kids[i], parent: parentId, index: i });
    this.#commit(parentId, { ...el, children: kids });
  }

  /** Whether any edit changed this part. */
  get changed(): boolean {
    return this.#root !== this.#source.root || this.#nextId !== this.#source.nextId;
  }

  /** The edited part. Its index is handed to the cache; this tree must not be used after. */
  result(): XmlPart {
    if (!this.changed) return this.#source;
    adoptNodeIndex(this.#root, this.#index.release());
    return { partName: this.partName, root: this.#root, nextId: this.#nextId };
  }

  /** Swap in `next` for the node with its id, and copy the path up to the root. */
  #commit(id: NodeId, next: Node): void {
    let loc = this.#index.get(id)!;
    let current: Node = next;
    this.#index.set(id, { node: next, parent: loc.parent, index: loc.index });
    while (loc.parent !== null) {
      const parentLoc: Location = this.#index.get(loc.parent)!;
      const parent = parentLoc.node as Element;
      const kids = parent.children.slice();
      kids[loc.index] = current;
      current = { ...parent, children: kids };
      this.#index.set(parent.id, {
        node: current,
        parent: parentLoc.parent,
        index: parentLoc.index,
      });
      loc = parentLoc;
    }
    this.#root = current as Element;
  }

  #learn(node: Node, parent: NodeId, index: number): void {
    const stack: Location[] = [{ node, parent, index }];
    while (stack.length > 0) {
      const item = stack.pop()!;
      workingTreeSubtreeVisits += 1;
      this.#index.set(item.node.id, item);
      if (item.node.type === 'element') {
        const kids = item.node.children;
        for (let i = 0; i < kids.length; i += 1)
          stack.push({ node: kids[i], parent: item.node.id, index: i });
      }
    }
  }

  #forget(node: Node): void {
    const stack: Node[] = [node];
    while (stack.length > 0) {
      const item = stack.pop()!;
      workingTreeSubtreeVisits += 1;
      this.#index.delete(item.id);
      if (item.type === 'element') for (const kid of item.children) stack.push(kid);
    }
  }
}

/** The height of an element subtree: a leaf element is 1. */
export function subtreeHeight(node: Element): number {
  let height = 0;
  const stack: [Element, number][] = [[node, 1]];
  while (stack.length > 0) {
    const [el, level] = stack.pop()!;
    if (level > height) height = level;
    for (const kid of el.children) if (kid.type === 'element') stack.push([kid, level + 1]);
  }
  return height;
}

/**
 * Derive kinds again for a subtree placed under a new context, as the reader would on reopen.
 * Subtrees whose kinds do not change are shared, not copied.
 */
export function rekind(node: Element, context: NodeKind | null): Element {
  const { kind, childContext } = deriveKind(node.ns, node.local, context);
  let kids: Node[] | null = null;
  node.children.forEach((kid, i) => {
    if (kid.type !== 'element') return;
    const next = rekind(kid, childContext);
    if (next === kid) return;
    kids ??= node.children.slice();
    kids[i] = next;
  });
  if (kind === node.kind && kids === null) return node;
  return { ...node, kind, children: kids ?? node.children };
}
