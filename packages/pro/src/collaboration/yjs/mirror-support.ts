/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Helpers of the package mirror (`mirror.ts`) that read no mirror state: the text and key-list
// syncs of a write, the listings a read skipped (`PackageMirror.dropHiddenListings`), and the
// runs a read showed by the move rescue (`rescuedParents`).

import * as Y from 'yjs';
import { readText, writeCase } from './text-case.ts';
import type {
  Element,
  Node,
  NodeId,
  PackagePart,
  ReplicaElementSpec,
} from '@pptx-editor.dev/core/store';
import { specKeys } from './move-rescue.ts';
import { inSchemaOrder } from './singleton-children.ts';
import { liveIds, type TextWrite } from './text-identity.ts';

/** One part's write: its key mapping, and its nodes before the change. */
export interface PartWrite {
  readonly partName: string;
  readonly caseChange: boolean;
  readonly ids: { readonly keyOf: Map<NodeId, string>; readonly idOf: Map<string, NodeId> };
  readonly before: Map<NodeId, Node>;
  /** What the write did to each text record's `Y.Text` (`text-identity.ts`). */
  readonly texts: Map<string, TextWrite>;
  readonly materialize: Set<string>;
  readonly force: Set<NodeId>;
}

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

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

/**
 * Make `text` hold `value` with one delete and one insert around the common ends, or, when the
 * changed span keeps its length, one replace per changed character. Returns what it deleted
 * and inserted, for the identity of moved text (`text-identity.ts`).
 */
export function syncText(
  text: Y.Text,
  value: string,
  created: boolean,
  caseChange = false
): TextWrite | undefined {
  const old = readText(text);
  if (old === value) return undefined;
  if (caseChange && !created && writeCase(text, old, value)) return { created, deleted: [] };
  let start = 0;
  while (start < old.length && start < value.length && old[start] === value[start]) start++;
  let end = 0;
  while (
    end < old.length - start &&
    end < value.length - start &&
    old[old.length - 1 - end] === value[value.length - 1 - end]
  )
    end++;
  // The common ends count UTF-16 units: never cut a surrogate pair at either edge, or Yjs
  // splits the pair and a peer reads two U+FFFD.
  const high = (c: number) => c >= 0xd800 && c <= 0xdbff;
  const low = (c: number) => c >= 0xdc00 && c <= 0xdfff;
  if (start > 0 && high(old.charCodeAt(start - 1))) start--;
  if (end > 0 && low(old.charCodeAt(old.length - end))) end--;
  const removed = old.length - start - end;
  const added = value.length - start - end;
  // The same length: replace each changed character in place, so a peer's
  // concurrent insert between two characters stays between them, not at the span's edge.
  // A span with a surrogate pair is replaced whole, so a replace never splits a pair.
  if (
    removed === added &&
    !/[\uD800-\uDFFF]/.test(old.slice(start, start + removed) + value.slice(start, start + added))
  ) {
    // The ids of the replaced characters, read in one walk before the first replace.
    const was = liveIds(text, start, removed);
    const deleted: Y.ID[] = [];
    for (let i = start; i < start + removed; i++) {
      if (old[i] === value[i]) continue;
      const id = was[i - start];
      if (id) deleted.push(id);
      // The new character goes before the old one, then the old one goes. A `Y.Text` insert
      // skips the tombstones to its right, so a character a peer types after a new one lands
      // after the old one's tombstone, and an undo, which puts the old character back right
      // before its tombstone, keeps the typed character after it.
      text.insert(i, value[i]!);
      text.delete(i + 1, 1);
    }
    return { created, deleted };
  }
  const deleted = removed > 0 ? liveIds(text, start, removed) : [];
  if (removed > 0) text.delete(start, removed);
  const doc = text.doc;
  const clock = doc ? Y.getState(doc.store, doc.clientID) : 0;
  if (added > 0) text.insert(start, value.slice(start, value.length - end));
  return {
    created,
    deleted,
    ...(added > 0 && doc ? { inserted: { client: doc.clientID, clock, length: added } } : {}),
  };
}

/** The inline nodes that the move rescue put back in one part (`rescuedParents`). */
export interface RescuedIn {
  /** Each node, with the parent the read shows it in. */
  readonly inParent: ReadonlyMap<string, string>;
  /** Each parent that shows one, and that parent's parent: a write there can take the node. */
  readonly watch: ReadonlySet<string>;
}

/**
 * Each inline node (a run, a line break, a field) that the move rescue put back (`placed` of
 * `rescueMoved`), with the parent the read shows it in.
 */
export function rescuedParents(spec: ReplicaElementSpec, placed: ReadonlySet<string>): RescuedIn {
  const inParent = new Map<string, string>();
  const watch = new Set<string>();
  if (placed.size === 0) return { inParent, watch };
  const up = new Map<string, string>();
  const stack: ReplicaElementSpec[] = [spec];
  while (stack.length > 0) {
    const s = stack.pop()!;
    for (const c of s.children ?? []) {
      if ('text' in c) continue;
      up.set(c.key, s.key);
      if (placed.has(c.key) && c.ns === NS_A && INLINE.has(c.local)) inParent.set(c.key, s.key);
      stack.push(c);
    }
  }
  for (const parent of inParent.values()) {
    watch.add(parent);
    const grand = up.get(parent);
    if (grand !== undefined) watch.add(grand);
  }
  return { inParent, watch };
}

const INLINE = new Set(['r', 'br', 'fld']);

/**
 * The child list to write: `desired`, unless it differs from the room's list only by the
 * schema order that a read puts children in (`inSchemaOrder`). Then the room's order stays,
 * and each new key goes before the next key of `desired` that the room holds. A write of a pure reorder
 * would move a peer's child, and an undo of that write would bring back a child that the
 * peer deleted in the meantime.
 */
export function inRoomOrder(
  parent: Element,
  current: readonly string[],
  desired: readonly string[],
  nameOf: (key: string) => { ns: string; local: string } | undefined
): readonly string[] {
  const inDesired = new Set(desired);
  const inCurrent = new Set(current);
  const common = current.filter((k) => inDesired.has(k));
  const desiredCommon = desired.filter((k) => inCurrent.has(k));
  const same = (x: readonly string[], y: readonly string[]) =>
    x.length === y.length && x.every((k, i) => k === y[i]);
  if (same(common, desiredCommon)) return desired;
  if (!same(inSchemaOrder(parent, common, nameOf), desiredCommon)) return desired;
  // Each new key goes before the next key of `desired` that the room holds, so it keeps its
  // place against a peer's child that the read moved; a key with none after it goes last.
  const before = new Map<string, string[]>();
  const tail: string[] = [];
  let next: string | undefined;
  for (let i = desired.length - 1; i >= 0; i--) {
    const k = desired[i]!;
    if (inCurrent.has(k)) next = k;
    else if (next === undefined) tail.unshift(k);
    else {
      let list = before.get(next);
      if (!list) before.set(next, (list = []));
      list.unshift(k);
    }
  }
  const out: string[] = [];
  for (const k of common) {
    for (const n of before.get(k) ?? []) out.push(n);
    out.push(k);
  }
  for (const n of tail) out.push(n);
  return out;
}
