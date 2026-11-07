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
import { adoptNodeIndex, layeredNodeIndex, nodeIndex, type Location } from './node-index.ts';

// Work counters for tests, beside `nodeIndexTestRecorder`: a transaction copies the index once
// per part it commits, and its edits write only the entries they touch. The kept index is a
// `KeptIndex`, so every write to it is counted, whichever code makes it.
let workingTreeOpens = 0;
let workingTreeSubtreeVisits = 0;
let workingTreeIndexWrites = 0;

/**
 * @internal Counts the working trees that transactions open, the nodes that edits add to or
 * remove from a kept index, and every entry written to or removed from a kept index (the copy a
 * working tree makes when it commits writes each entry once).
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
 * A transaction's id index: the entries that its edits write or remove, over the index of the
 * part it opened, behind methods that count each write. The edits never copy the part's index:
 * a transaction that never commits, such as the `can` probe of a toolbar button, costs what its
 * edits touch. `release()` hands the changes to the next root as a layered index
 * (`layeredNodeIndex`), which flattens into one map only when the layers grow.
 */
class KeptIndex {
  /** The index of the part the transaction opened. It is shared and never written. */
  readonly #base: ReadonlyMap<NodeId, Location>;
  /** The entries that edits wrote or removed (`undefined`), over `#base`. */
  readonly #changes = new Map<NodeId, Location | undefined>();

  constructor(source: ReadonlyMap<NodeId, Location>) {
    this.#base = source;
  }

  get(id: NodeId): Location | undefined {
    return this.#changes.has(id) ? this.#changes.get(id) : this.#base.get(id);
  }

  set(id: NodeId, loc: Location): void {
    workingTreeIndexWrites += 1;
    this.#changes.set(id, loc);
  }

  delete(id: NodeId): void {
    workingTreeIndexWrites += 1;
    this.#changes.set(id, undefined);
  }

  /** The whole index, for the cache. The caller gives up this index: nothing may change it. */
  release(): ReadonlyMap<NodeId, Location> {
    return layeredNodeIndex(this.#base, this.#changes);
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

  /**
   * Parents whose children from an index on hold stale positions in the index: an insert or a
   * removal shifts the siblings after it, and the index learns their new positions only when
   * one of them is read (`#loc`) or the tree is done (`result`). A run of inserts into one
   * large parent then costs what it inserts, not the number of siblings for each insert.
   */
  readonly #stale = new Map<NodeId, number>();

  /**
   * Elements this tree made by a path copy, with a child array of their own. No earlier
   * revision shares them, so a later edit below one writes into it in place, and the edits of
   * one paragraph of 50,000 runs do not copy its child list once per run.
   */
  readonly #owned = new WeakSet<Node>();

  /**
   * Nodes that a caller holds: read with `element`, or given to an edit. Their subtrees stay
   * as the caller saw them: an edit below one copies the path again (which the tree owns).
   */
  readonly #held = new WeakSet<Node>();

  /** Where a node sits now, or undefined. */
  get(id: NodeId): Location | undefined {
    return this.#loc(id);
  }

  /** The index entry of `id`, with the positions of its siblings brought up to date first. */
  #loc(id: NodeId): Location | undefined {
    const loc = this.#index.get(id);
    // A sibling before the first shifted place still sits where the index says.
    const from = loc && loc.parent !== null ? this.#stale.get(loc.parent) : undefined;
    if (from === undefined || loc!.index < from) return loc;
    this.#refresh(loc!.parent!);
    return this.#index.get(id);
  }

  /** Write the current positions of the children of `parentId` from its stale index on. */
  #refresh(parentId: NodeId): void {
    const from = this.#stale.get(parentId);
    if (from === undefined) return;
    this.#stale.delete(parentId);
    // A parent that an edit removed since has no children to bring up to date.
    const parent = this.#loc(parentId)?.node;
    if (parent?.type !== 'element') return;
    const kids = parent.children;
    for (let i = from; i < kids.length; i += 1)
      this.#index.set(kids[i].id, { node: kids[i], parent: parentId, index: i });
  }

  /**
   * Mark the children of `parentId` from `index` on as holding stale positions. `index` is the
   * place of the insert or the removal: every sibling that moved has a new place after it and,
   * unless an earlier mark already covers it, a stored place at it or after it.
   */
  #shifted(parentId: NodeId, index: number): void {
    const was = this.#stale.get(parentId);
    if (was === undefined || index < was) this.#stale.set(parentId, index);
  }

  /**
   * The element with `id`, or undefined when there is no such element. The caller may keep it:
   * later edits never change it.
   */
  element(id: NodeId): Element | undefined {
    const el = this.#element(id);
    if (el) this.#held.add(el);
    return el;
  }

  #element(id: NodeId): Element | undefined {
    const node = this.#loc(id)?.node;
    return node?.type === 'element' ? node : undefined;
  }

  /**
   * Whether the element `id` is one this tree may write in place: an element it made, with no
   * held node at it or above it.
   */
  #writable(id: NodeId): boolean {
    const loc = this.#loc(id);
    if (!loc || !this.#owned.has(loc.node)) return false;
    for (
      let at: Location | undefined = loc;
      at;
      at = at.parent === null ? undefined : this.#loc(at.parent)
    )
      if (this.#held.has(at.node)) return false;
    return true;
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
    for (let p = this.#loc(id)?.parent ?? null; p !== null; p = this.#loc(p)!.parent) level += 1;
    return level;
  }

  /** Whether `ancestor` is `id` or one of its ancestors. */
  isSelfOrAncestor(ancestor: NodeId, id: NodeId): boolean {
    for (let cur: NodeId | null = id; cur !== null; cur = this.#loc(cur)?.parent ?? null)
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
      const loc = this.#loc(cur);
      if (!loc || loc.node.type !== 'element') return null;
      if (!isTransparentWrapper(loc.node.ns, loc.node.local)) return loc.node.kind;
      cur = loc.parent;
    }
    return null;
  }

  /** Replace an element by a copy with the same id and the same children (an attribute edit). */
  replace(id: NodeId, next: Element): void {
    // The caller made `next`, and it may share the child array of the element it replaces.
    this.#held.add(next);
    this.#commit(id, next);
  }

  /** Replace all children of an element. */
  replaceChildren(id: NodeId, kids: readonly Node[]): void {
    const el = this.#element(id)!;
    // Every child is learned again at its position below.
    this.#stale.delete(id);
    for (const old of el.children) this.#forget(old);
    kids.forEach((kid, i) => this.#learn(kid, id, i));
    // The child array is the caller's.
    const next = { ...el, children: kids };
    this.#held.add(next);
    this.#commit(id, next);
  }

  /** Insert `node` (not yet in the tree) as child `index` of `parentId`. */
  insertChild(parentId: NodeId, index: number, node: Node): void {
    // The caller made `node`, or a move took it from a subtree that a caller may hold.
    this.#held.add(node);
    const el = this.#element(parentId)!;
    const writable = this.#writable(parentId);
    const kids = writable ? (el.children as Node[]) : el.children.slice();
    kids.splice(index, 0, node);
    // From the insert place: the sibling stored there moved too.
    if (index + 1 < kids.length) this.#shifted(parentId, index);
    this.#learn(node, parentId, index);
    if (!writable) this.#commit(parentId, this.#own({ ...el, children: kids }));
  }

  /** Remove the node `id` (not the root) with its subtree. */
  remove(id: NodeId): void {
    const loc = this.#loc(id)!;
    const parentId = loc.parent!;
    const el = this.#element(parentId)!;
    // A moved node keeps its subtree, which a caller may hold through the old place.
    this.#held.add(loc.node);
    const writable = this.#writable(parentId);
    const kids = writable ? (el.children as Node[]) : el.children.slice();
    kids.splice(loc.index, 1);
    this.#forget(loc.node);
    if (loc.index < kids.length) this.#shifted(parentId, loc.index);
    if (!writable) this.#commit(parentId, this.#own({ ...el, children: kids }));
  }

  /** Whether any edit changed this part. */
  get changed(): boolean {
    return this.#root !== this.#source.root || this.#nextId !== this.#source.nextId;
  }

  /** The edited part. Its index is handed to the cache; this tree must not be used after. */
  result(): XmlPart {
    if (!this.changed) return this.#source;
    for (const parentId of [...this.#stale.keys()]) this.#refresh(parentId);
    adoptNodeIndex(this.#root, this.#index.release());
    return { partName: this.partName, root: this.#root, nextId: this.#nextId };
  }

  #own(el: Element): Element {
    this.#owned.add(el);
    return el;
  }

  /**
   * Swap in `next` for the node with its id, and copy the path up to the root. The copy stops
   * at the first element the tree may write in place (`#writable`): the elements above it
   * already hold it.
   */
  #commit(id: NodeId, next: Node): void {
    let loc = this.#loc(id)!;
    this.#index.set(id, { node: next, parent: loc.parent, index: loc.index });
    // The path to the root, and the highest held node on it: the path up to that node is
    // copied, so a caller that holds it sees no change.
    const path: Location[] = [];
    let held = -1;
    for (let at = loc; at.parent !== null;) {
      at = this.#loc(at.parent)!;
      if (this.#held.has(at.node)) held = path.length;
      path.push(at);
    }
    let current: Node = next;
    for (let i = 0; i < path.length; i += 1) {
      const parentLoc = path[i];
      const parent = parentLoc.node as Element;
      if (i > held && this.#owned.has(parent)) {
        (parent.children as Node[])[loc.index] = current;
        return;
      }
      const kids = parent.children.slice();
      kids[loc.index] = current;
      current = this.#own({ ...parent, children: kids });
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
