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

/**
 * Record an index built for `root` elsewhere. The caller gives up the map: nothing may change
 * it afterwards. Ignored when the root already has one.
 */
export function adoptNodeIndex(root: Element, index: ReadonlyMap<NodeId, Location>): void {
  if (!indexCache.has(root)) indexCache.set(root, index);
}
