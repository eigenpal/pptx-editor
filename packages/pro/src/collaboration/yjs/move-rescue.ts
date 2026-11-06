/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The move rescue of the text lineage (openspec change `collaboration-text-lineage`): a node
// that a peer moved into a subtree another peer removed without having seen it goes back to its
// old place, or to an earlier one when that place is gone too. When every old place is gone, it
// follows the children that moved out of the removed node (a join of a paragraph whose own
// join is in flight). Adapted from DOCX Editor
// `packages/pro/src/collaboration/document/materialize-orphans.ts` at 5ffe9674. The fields it
// reads are written by `markLineage` in `text-lineage.ts`. Everything read here comes from the
// room and is untrusted.

import * as Y from 'yjs';
import type { ReplicaElementSpec } from '@pptx-editor.dev/core/store';
import { F, type Room } from './schema.ts';
import { inSchemaOrder } from './singleton-children.ts';
import {
  isCount,
  isRemovalField,
  isRemoved,
  LINEAGE,
  MAX_EARLIER,
  MAX_PATH,
  MAX_REMOVERS,
  MAX_SNAPSHOT,
  MOVES_KEY,
  type PartLineage,
  type ReadRecord,
  type Spec,
} from './text-lineage.ts';

/** The most moved nodes one part read puts back; a hostile room may list more. */
const MAX_RESCUES = 1_024;

/** An ancestor's removal snapshot for a read: its keys, `all` when too large, null when none. */
type Snapshot = ReadonlySet<string> | 'all' | null;

/**
 * The keys that any removal of a record saw: the union of every writer's snapshot. A snapshot
 * or a count of writers larger than any writer makes reads as `all`; a value of a wrong type
 * is ignored.
 */
function removalSnapshot(rec: unknown): Snapshot {
  if (!(rec instanceof Y.Map)) return null;
  let out: Set<string> | null = null;
  let removers = 0;
  for (const [field, list] of (rec as Y.Map<unknown>).entries()) {
    if (!isRemovalField(field) || !Array.isArray(list)) continue;
    if (list.length > MAX_SNAPSHOT + 1 || (removers += 1) > MAX_REMOVERS) return 'all';
    out ??= new Set();
    for (const k of list as unknown[]) if (typeof k === 'string') out.add(k);
  }
  return out;
}

/**
 * The ancestor whose removal did not see a moved node, or null when the removal saw it: the
 * walk up its destination chain. A chain or a snapshot longer than any writer makes counts as
 * seen (no rescue). `snapshots` keeps each ancestor's snapshot as a set, so a read builds each
 * one once.
 */
function removedUnseen(
  room: Room,
  key: string,
  to: readonly unknown[],
  reachable: ReadonlyMap<string, number>,
  snapshots: Map<string, Snapshot>
): string | null {
  if (to.length > MAX_PATH) return null;
  for (const ancestor of to) {
    if (typeof ancestor !== 'string' || reachable.has(ancestor)) return null;
    let held = snapshots.get(ancestor);
    if (held === undefined) {
      held = removalSnapshot(room.nodes.get(ancestor));
      snapshots.set(ancestor, held);
    }
    if (held === 'all') return null;
    if (held) return held.has(key) ? null : ancestor;
  }
  return null;
}

/** One order of keys on every replica. */
const byKey = (x: string, y: string): number => (x < y ? -1 : x > y ? 1 : 0);

/** A set that also holds every key of `base`, without a copy of it. */
class Overlay extends Set<string> {
  readonly #base: ReadonlySet<string>;
  constructor(base: ReadonlySet<string>) {
    super();
    this.#base = base;
  }
  override has(key: string): boolean {
    return this.#base.has(key) || super.has(key);
  }
}

/**
 * A part's spec with the nodes put back that a peer moved into a subtree another peer removed
 * without having seen them: each goes to its old parent, if that parent is in the part, at its
 * old index, in schema order. A node removed after its move stays removed. A node inside
 * another rescued node goes with it, and no key is placed twice. The key of each node put back
 * is added to `placed`.
 */
export function rescueMoved(
  room: Room,
  partName: string,
  spec: ReplicaElementSpec,
  read: ReadRecord,
  placed?: Set<string>
): ReplicaElementSpec {
  const moves = room.doc.getMap(MOVES_KEY);
  if (moves.size === 0) return spec;
  const depthOf = new Map<string, number>();
  // Each shown key's parent and its place in document order, for a node that follows a join.
  const parentOf = new Map<string, string>();
  const orderOf = new Map<string, number>();
  const collect = (s: Spec, depth: number, parent: string | null): void => {
    depthOf.set(s.key, depth);
    orderOf.set(s.key, orderOf.size);
    if (parent !== null) parentOf.set(s.key, parent);
    if (!('text' in s)) for (const c of s.children ?? []) collect(c, depth + 1, s.key);
  };
  collect(spec, 0, null);
  // `from` is the parent it goes into. A node that follows a join also has `after`, the key it
  // goes right after, and `seq`, its place in the order such nodes are placed.
  type Candidate = {
    key: string;
    from: string;
    index: number;
    clock: number;
    to: unknown[];
    after?: string;
    seq?: number;
  };
  const candidates: Candidate[] = [];
  // Nodes whose old parents are all gone: `last` is the parent they left, `removed` the
  // ancestor whose removal did not see them.
  type Orphan = {
    key: string;
    last: string;
    removed: string;
    index: number;
    clock: number;
    to: unknown[];
  };
  const orphans: Orphan[] = [];
  const snapshots = new Map<string, Snapshot>();
  // Built on first use: each old parent's last moved-out key that the read shows.
  let lastOut: Map<string, string> | undefined;
  const lastShownOutOf = (parent: string): string | undefined => {
    if (!lastOut) {
      lastOut = new Map();
      for (const key of moves.keys()) {
        if (!parentOf.has(key)) continue;
        const rec = room.nodes.get(key);
        const old = rec instanceof Y.Map ? rec.get(LINEAGE.movedFrom) : undefined;
        if (typeof old !== 'string') continue;
        const was = lastOut.get(old);
        if (was === undefined || orderOf.get(key)! > orderOf.get(was)!) lastOut.set(old, key);
      }
    }
    return lastOut.get(parent);
  };
  for (const [key, part] of moves.entries()) {
    if (part !== partName || depthOf.has(key)) continue;
    const rec = room.nodes.get(key);
    if (!(rec instanceof Y.Map) || rec.get(F.part) !== partName) continue;
    if (isRemoved(rec as Y.Map<unknown>)) continue;
    const last = rec.get(LINEAGE.movedFrom);
    const to = rec.get(LINEAGE.movedTo);
    if (typeof last !== 'string' || !Array.isArray(to)) continue;
    // The parent it came from, or, when that one is gone too (a node moved twice), the latest
    // earlier parent the part still holds, at its end.
    const earlier = rec.get(LINEAGE.movedEarlier);
    const from = depthOf.has(last)
      ? last
      : (Array.isArray(earlier) ? (earlier as unknown[]).slice(0, MAX_EARLIER) : []).find(
          (k): k is string => typeof k === 'string' && depthOf.has(k)
        );
    const at = rec.get(LINEAGE.movedIndex);
    const q = rec.get(LINEAGE.movedClock);
    const clock = typeof q === 'number' && Number.isFinite(q) ? q : 0;
    const removed = removedUnseen(room, key, to, depthOf, snapshots);
    if (removed === null) continue;
    // More than any writer makes in one read: a hostile room. Put none back.
    if (candidates.length + orphans.length >= MAX_RESCUES) return spec;
    if (from !== undefined) {
      const index = from === last && isCount(at) ? at : Infinity;
      candidates.push({ key, from, index, clock, to });
    } else orphans.push({ key, last, removed, index: isCount(at) ? at : Infinity, clock, to });
  }
  // Every parent an orphan came from is gone too. When children that moved out of the removed
  // node show elsewhere (the remover's join moved them; the room does not record who moved a
  // node), the orphan goes with them: right after the last one shown, or
  // after the last orphan placed so that left the removed node, in the order it left its own
  // parent. A peer joined that paragraph into the one before it while this node's join moved
  // it in; with more joins at once, the orphans form a chain. Each
  // round places the orphans whose anchor is known, so an anchor is placed before its
  // followers. At most one round per orphan.
  const placedOutOf = new Map<string, { key: string; into: string }>();
  let pending = orphans;
  let seq = 0;
  while (pending.length > 0) {
    const round: (Orphan & { after: string; into: string })[] = [];
    const rest: Orphan[] = [];
    for (const o of pending) {
      const follower = placedOutOf.get(o.removed);
      const shown = follower ? undefined : lastShownOutOf(o.removed);
      const into = follower ? follower.into : shown === undefined ? undefined : parentOf.get(shown);
      const after = follower ? follower.key : shown;
      if (after === undefined || into === undefined) rest.push(o);
      else round.push({ ...o, after, into });
    }
    if (round.length === 0) break;
    round.sort((x, y) => x.clock - y.clock || x.index - y.index || byKey(x.key, y.key));
    for (const o of round) {
      candidates.push({ ...o, from: o.into, seq: seq++ });
      placedOutOf.set(o.last, { key: o.key, into: o.into });
    }
    pending = rest;
  }
  if (candidates.length === 0) return spec;
  // One order on every replica: a candidate that moved into another candidate comes after it,
  // so one read takes it with its new parent; then by key.
  const keys = new Set(candidates.map((c) => c.key));
  const outer = new Map(
    candidates.map((c) => [c.key, c.to.filter((k) => typeof k === 'string' && keys.has(k)).length])
  );
  candidates.sort((a, b) => outer.get(a.key)! - outer.get(b.key)! || byKey(a.key, b.key));
  // One read of each candidate, all sharing one set of keys already placed: a node inside an
  // earlier candidate reads as null and goes with it, so no key is placed twice.
  const seen = new Overlay(new Set(depthOf.keys()));
  type Addition = {
    node: ReplicaElementSpec;
    index: number;
    clock: number;
    after?: string;
    seq: number;
  };
  const additions = new Map<string, Addition[]>();
  for (const c of candidates) {
    const node = read(c.key, seen, depthOf.get(c.from)! + 1);
    if (!node) continue;
    placed?.add(node.key);
    const entry: Addition = {
      node,
      index: c.index,
      clock: c.clock,
      after: c.after,
      seq: c.seq ?? 0,
    };
    const list = additions.get(c.from);
    if (list) list.push(entry);
    else additions.set(c.from, [entry]);
  }
  if (additions.size === 0) return spec;
  const place = (s: ReplicaElementSpec): ReplicaElementSpec => {
    const extra = additions.get(s.key);
    const children = (s.children ?? []).map((c) => ('text' in c ? c : place(c)));
    if (!extra) return { ...s, children };
    // The moves undone in reverse: each index was taken in the list as it was before its
    // write, so the latest write (the mover's clock) goes back first, and the nodes of one
    // write go back in the order of their indexes. Then schema order.
    const placed: Spec[] = [...children];
    const order = extra
      .filter((x) => x.after === undefined)
      .sort((x, y) => y.clock - x.clock || x.index - y.index || byKey(x.node.key, y.node.key));
    for (const { node, index } of order) placed.splice(Math.min(index, placed.length), 0, node);
    // A node that follows a join goes right after its anchor, after the nodes before it that
    // follow the same anchor, in the order the rounds above placed them.
    const following = extra.filter((x) => x.after !== undefined).sort((x, y) => x.seq - y.seq);
    const lastAfter = new Map<string, string>();
    for (const { node, after } of following) {
      const anchor = lastAfter.get(after!) ?? after!;
      const at = placed.findIndex((k) => k.key === anchor);
      placed.splice(at < 0 ? placed.length : at + 1, 0, node);
      lastAfter.set(after!, node.key);
    }
    const name = { ns: s.ns, local: s.local };
    return {
      ...s,
      children: [...inSchemaOrder(name, placed, (k) => ('text' in k ? undefined : k))],
    };
  };
  return place(spec);
}

/** Every key of a read spec. */
export function specKeys(spec: ReplicaElementSpec): Set<string> {
  const out = new Set<string>();
  const stack: Spec[] = [spec];
  while (stack.length > 0) {
    const s = stack.pop()!;
    out.add(s.key);
    if (!('text' in s)) for (const c of s.children ?? []) stack.push(c);
  }
  return out;
}

const sameSet = (x: ReadonlySet<string>, y: ReadonlySet<string>): boolean =>
  x.size === y.size && [...x].every((k) => y.has(k));

/** Whether two part lineages read alike: a read after a rescue needs no second pass then. */
export function sameLineage(x: PartLineage, y: PartLineage): boolean {
  return sameSet(x.lost, y.lost) && sameSet(x.unseenWinners, y.unseenWinners);
}
