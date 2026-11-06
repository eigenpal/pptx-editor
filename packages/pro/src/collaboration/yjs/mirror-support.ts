/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Helpers of the package mirror (`mirror.ts`) that read no mirror state: the key-list sync of a
// write, and the listings a read skipped (`PackageMirror.dropHiddenListings`).

import type * as Y from 'yjs';
import type {
  Element,
  Node,
  NodeId,
  PackagePart,
  ReplicaElementSpec,
} from '@pptx-editor.dev/core/store';
import { specKeys } from './move-rescue.ts';

/** Every node of a tree by id. */
export function indexTree(root: Element | undefined): Map<NodeId, Node> {
  const out = new Map<NodeId, Node>();
  if (!root) return out;
  const stack: Node[] = [root];
  while (stack.length > 0) {
    const n = stack.pop()!;
    out.set(n.id, n);
    if (n.type === 'element') for (const c of n.children) stack.push(c);
  }
  return out;
}

/**
 * The listings of one child list that a read skipped, because the read showed their key at an
 * earlier place: index → key, and the length of the list when it was read.
 */
export interface SkippedListings {
  readonly length: number;
  readonly at: Map<number, string>;
}

/** Longest common subsequence of two key lists, as index sets. Small lists only. */
function lcs(
  a: readonly string[],
  b: readonly string[]
): { keepA: Set<number>; keepB: Set<number> } {
  const n = a.length;
  const m = b.length;
  const table = new Uint32Array((n + 1) * (m + 1));
  for (let i = n - 1; i >= 0; i--)
    for (let j = m - 1; j >= 0; j--)
      table[i * (m + 1) + j] =
        a[i] === b[j]
          ? table[(i + 1) * (m + 1) + j + 1] + 1
          : Math.max(table[(i + 1) * (m + 1) + j], table[i * (m + 1) + j + 1]);
  const keepA = new Set<number>();
  const keepB = new Set<number>();
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (a[i] === b[j]) {
      keepA.add(i++);
      keepB.add(j++);
    } else if (table[(i + 1) * (m + 1) + j] >= table[i * (m + 1) + j + 1]) i++;
    else j++;
  }
  return { keepA, keepB };
}

/** Make `list` hold `desired`, keeping the longest common run of keys in place. */
export function syncList(list: Y.Array<string>, desired: readonly string[]): void {
  const current = list.toArray();
  if (current.length === desired.length && current.every((k, i) => k === desired[i])) return;
  if (current.length * desired.length > 250_000) {
    list.delete(0, current.length);
    list.insert(0, [...desired]);
    return;
  }
  const { keepA, keepB } = lcs(current, desired);
  for (let i = current.length - 1; i >= 0; i--) if (!keepA.has(i)) list.delete(i, 1);
  let pos = 0;
  desired.forEach((key, j) => {
    if (!keepB.has(j)) list.insert(pos, [key]);
    pos += 1;
  });
}

/**
 * The skipped listings of the parents that `spec` shows. A parent the read did not show (a
 * hidden copy, a dropped piece) keeps its list as the room holds it.
 */
export function shownParentsOnly(
  skipped: Map<string, SkippedListings>,
  spec: ReplicaElementSpec
): Map<string, SkippedListings> {
  if (skipped.size === 0) return skipped;
  const shown = specKeys(spec);
  return new Map([...skipped].filter(([parent]) => shown.has(parent)));
}

/** Each node's place in a part: the ids of its ancestors, from its parent up, joined. */
function placesOf(part: PackagePart | undefined): Map<NodeId, string> {
  const out = new Map<NodeId, string>();
  if (part?.kind !== 'xml') return out;
  const stack: [Node, string][] = [[part.xml.root, '']];
  while (stack.length > 0) {
    const [n, place] = stack.pop()!;
    out.set(n.id, place);
    if (n.type === 'element') for (const c of n.children) stack.push([c, `${n.id}/${place}`]);
  }
  return out;
}

/**
 * Whether a local change (`before` → `after`) keeps the place where a node shows: the node and
 * each of its ancestors keep their parents. `idOf` maps a key to this replica's node id.
 */
export function shownPlaceKept(
  idOf: ReadonlyMap<string, NodeId> | undefined,
  before: PackagePart | undefined,
  after: PackagePart | undefined
): (key: string) => boolean {
  let was: Map<NodeId, string> | undefined;
  let now: Map<NodeId, string> | undefined;
  return (key) => {
    const id = idOf?.get(key);
    if (id === undefined) return false;
    was ??= placesOf(before);
    now ??= placesOf(after);
    const a = was.get(id);
    return a !== undefined && a === now.get(id);
  };
}

/**
 * The key of the parent that a node shows in before a local change (`before`): `idOf` and
 * `keyOf` map keys and this replica's node ids.
 */
export function shownParentKey(
  idOf: ReadonlyMap<string, NodeId> | undefined,
  keyOf: ReadonlyMap<NodeId, string> | undefined,
  before: PackagePart | undefined
): (key: string) => string | undefined {
  let parents: Map<NodeId, NodeId> | undefined;
  return (key) => {
    const id = idOf?.get(key);
    if (id === undefined || before?.kind !== 'xml') return undefined;
    if (!parents) {
      const found = new Map<NodeId, NodeId>();
      const stack: Node[] = [before.xml.root];
      while (stack.length > 0) {
        const n = stack.pop()!;
        if (n.type === 'element')
          for (const c of n.children) {
            found.set(c.id, n.id);
            stack.push(c);
          }
      }
      parents = found;
    }
    const parent = parents.get(id);
    return parent === undefined ? undefined : keyOf?.get(parent);
  };
}
