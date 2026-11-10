/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The move rescue of the text lineage (openspec change `collaboration-text-lineage`): a node
// that a peer moved into a subtree another peer removed without having seen it goes back to its
// old place, or to an earlier one when that place is gone too. When every old place is gone, it
// follows the children that moved out of the removed node (a join of a paragraph whose own
// join is in flight). After DOCX Editor
// `packages/pro/src/collaboration/document/materialize-orphans.ts` at 5ffe9674. The fields it
// reads are written by `markLineage` in `text-lineage.ts`. Everything read here comes from the
// room and is untrusted.

import * as Y from 'yjs';
import type { ReplicaElementSpec } from '@pptx-editor.dev/core/store';
import { F, type Room } from './schema.ts';
import type { Listings } from './removed-content.ts';
import { inSchemaOrder } from './singleton-children.ts';
import { removalSnapshot, removalsSawAnywhere, type Snapshot } from './removal-snapshots.ts';
import { undoneWithout } from './undo-unlisted.ts';
import {
  isCount,
  isRemoved,
  LINEAGE,
  MAX_EARLIER,
  MAX_PATH,
  MOVES_KEY,
  type PartLineage,
  type ReadRecord,
  type Spec,
} from './text-lineage.ts';

/** The most moved nodes one part read puts back; a hostile room may list more. */
const MAX_RESCUES = 1_024;

/**
 * The ancestor whose removal did not see a moved node, or null when the removal saw it: the
 * walk up its destination chain. A chain or a snapshot longer than any writer makes counts as
 * seen (no rescue). An ancestor that an undo took out without seeing the move (`move`, the
 * item of the move's `mt` field) counts as removed unseen. `snapshots` keeps each ancestor's
 * snapshot as a set, so a read builds each one once.
 */
function removedUnseen(
  room: Room,
  key: string,
  to: readonly unknown[],
  reachable: ReadonlyMap<string, number>,
  snapshots: Map<string, Snapshot>,
  move: Y.Item | undefined
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
    // An undo of the step that made the ancestor (a row insert) took it without a snapshot.
    if (undoneWithout(room, ancestor, move)) return ancestor;
  }
  return null;
}

/** The item of a record field: who wrote it, and when. */
function fieldItem(rec: Y.Map<unknown>, field: string): Y.Item | undefined {
  const item = rec._map.get(field);
  return item && !item.deleted ? item : undefined;
}

/**
 * Whether a removal of a parent in the chain `to` saw `key`: its snapshot holds it. A peer that
 * removed a row with the cell a node came from saw the node there, also when the node moved into
 * that row at the same time.
 */
function removalsSaw(
  room: Room,
  to: readonly unknown[],
  key: string,
  snapshots: Map<string, Snapshot>
): boolean {
  for (const ancestor of to.slice(0, MAX_PATH)) {
    if (typeof ancestor !== 'string') return true;
    let held = snapshots.get(ancestor);
    if (held === undefined) {
      held = removalSnapshot(room.nodes.get(ancestor));
      snapshots.set(ancestor, held);
    }
    if (held === 'all' || held?.has(key)) return true;
  }
  return false;
}

/** The longest chain of parents that `removalsAbove` walks up. */
const MAX_UP = 64;

/** A moved node that the rescue places, as `moveSibling` reads it. */
interface Placed {
  readonly key: string;
  readonly from: string;
  readonly clock: number;
  readonly to: readonly unknown[];
}

/**
 * The node that the write which moved `node` moved with it into the same place, and that the
 * rescue places: the last one before `node` in that place's list (`after`), or else the first
 * one after it. `at` is the index of `node` in that list. One write moves nodes with one clock
 * (`mq`) and one chain of parents (`mt`).
 */
function moveSibling(
  room: Room,
  node: { key: string; clock: number; to: readonly unknown[] },
  placed: readonly Placed[]
): { with: Placed; after: boolean; at: number } | null {
  const dest = node.to[0];
  if (typeof dest !== 'string') return null;
  const rec = room.nodes.get(dest);
  const list = rec instanceof Y.Map ? rec.get(F.children) : undefined;
  if (!(list instanceof Y.Array)) return null;
  const order = (list.toArray() as unknown[]).slice(0, MAX_RESCUES);
  const at = order.indexOf(node.key);
  if (at < 0) return null;
  const same = (c: Placed): boolean =>
    c.key !== node.key &&
    c.clock === node.clock &&
    c.to.length === node.to.length &&
    c.to.every((k, i) => k === node.to[i]);
  const byKeyOf = new Map(placed.filter(same).map((c) => [c.key, c]));
  if (byKeyOf.size === 0) return null;
  for (let i = at - 1; i >= 0; i -= 1) {
    const c = byKeyOf.get(order[i] as string);
    if (c) return { with: c, after: true, at };
  }
  for (let i = at + 1; i < order.length; i += 1) {
    const c = byKeyOf.get(order[i] as string);
    if (c) return { with: c, after: false, at };
  }
  return null;
}

/**
 * What the removals at or above `start` (by the records that list it, up to `MAX_UP` levels) saw
 * of `key`: `seen` when a snapshot holds it (or is too large to read), `unseen` when there are
 * snapshots and none holds it, and `none` when no record there was removed.
 */
function removalsAbove(
  room: Room,
  start: string,
  key: string,
  listings: Listings,
  snapshots: Map<string, Snapshot>
): 'seen' | 'unseen' | 'none' {
  let found = false;
  const visited = new Set<string>();
  let level = [start];
  for (let depth = 0; depth < MAX_UP && level.length > 0; depth += 1) {
    const up: string[] = [];
    for (const k of level) {
      if (visited.has(k)) continue;
      visited.add(k);
      let held = snapshots.get(k);
      if (held === undefined) {
        held = removalSnapshot(room.nodes.get(k));
        snapshots.set(k, held);
      }
      if (held === 'all' || held?.has(key)) return 'seen';
      if (held) found = true;
      for (const parent of listings().get(k) ?? []) up.push(parent);
    }
    level = up.slice(0, MAX_RESCUES);
  }
  return found ? 'unseen' : 'none';
}

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const isA = (s: Spec | undefined, local: string): s is ReplicaElementSpec =>
  !!s && !('text' in s) && s.ns === NS_A && s.local === local;

/** The keys a record lists, or none. */
function listedOf(room: Room, key: string): unknown[] {
  const rec = room.nodes.get(key);
  const list = rec instanceof Y.Map ? rec.get(F.children) : undefined;
  return list instanceof Y.Array ? (list.toArray() as unknown[]).slice(0, MAX_RESCUES) : [];
}

/** The DrawingML local name of a record, or undefined. */
function localOf(room: Room, key: unknown): string | undefined {
  const rec = typeof key === 'string' ? room.nodes.get(key) : undefined;
  const shell: unknown = rec instanceof Y.Map ? rec.get(F.shell) : undefined;
  if (typeof shell !== 'object' || shell === null) return undefined;
  const { ns, local } = shell as { ns?: unknown; local?: unknown };
  return ns === NS_A && typeof local === 'string' ? local : undefined;
}

/**
 * The place in a table cell that stays, for a paragraph or a text body whose chain of parents
 * `to` (its text body or its cell, then its row and its table) lost the cell or the row. A cell
 * must show its text, and a removal that did not see the node did not remove it on purpose.
 * From a removed row: the cell in the same column (its index in the removed row, else the first
 * column) of the first row of the table that can take it. From a removed cell of a row that stays: the first cell of
 * that row that can take it. A paragraph goes into the cell's text body, so a cell without one
 * cannot take it; a text body goes into the cell, beside any body it has (core joins the bodies
 * of a cell, `cellAsSaved`). Every replica reads the same room and spec, so all pick the same.
 */
function survivingCell(
  room: Room,
  node: string,
  to: readonly unknown[],
  specOf: ReadonlyMap<string, ReplicaElementSpec>
): string | undefined {
  const kind = localOf(room, node);
  const at = kind === 'p' ? 1 : kind === 'txBody' ? 0 : -1;
  if (at < 0 || localOf(room, to[at]) !== 'tc') return undefined;
  const [cell, row, table] = to.slice(at, at + 3) as string[];
  if (localOf(room, row) !== 'tr') return undefined;
  const into = (tc: Spec | undefined): string | undefined => {
    if (!isA(tc, 'tc')) return undefined;
    if (kind === 'txBody') return tc.key;
    return (tc.children ?? []).find((k) => isA(k, 'txBody'))?.key;
  };
  const shownRow = specOf.get(row!);
  if (!specOf.has(cell!) && isA(shownRow, 'tr'))
    for (const tc of shownRow.children ?? []) {
      const place = into(tc);
      if (place) return place;
    }
  const shownTable = specOf.get(table ?? '');
  if (shownRow || !isA(shownTable, 'tbl')) return undefined;
  // A column delete takes the cell out of its row too: its column is gone, so the first one.
  const column = Math.max(0, listedOf(room, row!).indexOf(cell));
  for (const tr of shownTable.children ?? []) {
    if (!isA(tr, 'tr')) continue;
    const cells = (tr.children ?? []).filter((k) => isA(k, 'tc'));
    const place = into(cells[Math.min(column, cells.length - 1)]);
    if (place) return place;
  }
  return undefined;
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
 * is added to `placed`. A node with no place left adds the removed ancestor that did not see it
 * to `stranded`: a removed paragraph then comes back with it (`withRemovedRuns`). With
 * `listings`, a node whose old parent a peer removed after the node left it goes with the nodes
 * the same write moved (`followMoveSiblings`).
 */
export function rescueMoved(
  room: Room,
  partName: string,
  spec: ReplicaElementSpec,
  read: ReadRecord,
  placed?: Set<string>,
  stranded?: Set<string>,
  listings?: Listings
): ReplicaElementSpec {
  const moves = room.doc.getMap(MOVES_KEY);
  if (moves.size === 0) return spec;
  const depthOf = new Map<string, number>();
  // Each shown key's parent and its place in document order, for a node that follows a join.
  const parentOf = new Map<string, string>();
  const orderOf = new Map<string, number>();
  const specOf = new Map<string, ReplicaElementSpec>();
  const collect = (s: Spec, depth: number, parent: string | null): void => {
    depthOf.set(s.key, depth);
    orderOf.set(s.key, orderOf.size);
    if (parent !== null) parentOf.set(s.key, parent);
    if (!('text' in s)) specOf.set(s.key, s);
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
    /** The key it goes right before (a node that follows the nodes moved with it). */
    before?: string;
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
    const move = fieldItem(rec as Y.Map<unknown>, LINEAGE.movedTo);
    const removed = removedUnseen(room, key, to, depthOf, snapshots, move);
    if (removed === null) continue;
    // A removal of the place that lists it now saw it there: a later write moved it on.
    if (listings && removalsAbove(room, key, key, listings, snapshots) === 'seen') continue;
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
  // An orphan that left its cell before a peer removed that cell goes with the nodes that the
  // same write moved with it (`sameMove`), as they go back: the cell's remover saw it leave, so
  // only the removal of its new place took it, and that removal did not see it. Each orphan
  // goes right after the last of them before it in the place they went to, or right before the
  // first one after it. Returns the orphans that have no such node.
  const followMoveSiblings = (rest: readonly Orphan[]): Orphan[] => {
    if (!listings) return [...rest];
    let left = [...rest];
    for (let round = 0; round < rest.length && left.length > 0; round += 1) {
      const next: Orphan[] = [];
      const placed: Candidate[] = [];
      for (const o of left) {
        const sibling = moveSibling(room, o, candidates);
        // No removal saw it: not one of the chain it moved into, not one above the place it
        // is listed now, and not one above its old parent, which a peer removed after it left.
        const free =
          !removalsSaw(room, o.to, o.key, snapshots) &&
          removalsAbove(room, o.key, o.key, listings, snapshots) !== 'seen' &&
          removalsAbove(room, o.last, o.key, listings, snapshots) === 'unseen';
        if (!sibling || !free) {
          next.push(o);
          continue;
        }
        const anchor = sibling.after ? { after: sibling.with.key } : { before: sibling.with.key };
        // `seq` holds its place in the list it went to, for the order below.
        const { key, clock, to } = o;
        placed.push({
          key,
          from: sibling.with.from,
          index: Infinity,
          clock,
          to,
          ...anchor,
          seq: sibling.at,
        });
      }
      if (placed.length === 0) break;
      placed.sort((x, y) => x.seq! - y.seq! || byKey(x.key, y.key));
      for (const c of placed) candidates.push({ ...c, seq: seq++ });
      left = next;
    }
    return left;
  };
  // A paragraph or a text body whose old cell and new cell both went, and whose removals did
  // not see it there, goes to a table cell that stays (`survivingCell`).
  let seenAnywhere: ReadonlySet<string> | 'all' | undefined;
  const cellOf = (o: Orphan): string | undefined => {
    if (!listings) return undefined;
    const free =
      !removalsSaw(room, o.to, o.key, snapshots) &&
      removalsAbove(room, o.last, o.key, listings, snapshots) !== 'seen';
    if (!free) return undefined;
    // A removal of a place where a read showed the node saw it there.
    seenAnywhere ??= removalsSawAnywhere(room, partName);
    if (seenAnywhere === 'all' || seenAnywhere.has(o.key)) return undefined;
    let into = survivingCell(room, o.key, o.to, specOf);
    if (into) return into;
    // Concurrent removals can leave no row, although none removed the moved text. Keep a
    // valid row as its container. Text that a removal saw is already deleted from its records.
    const rowKey = o.to.find((k) => localOf(room, k) === 'tr');
    const tableKey = o.to.find((k) => localOf(room, k) === 'tbl');
    const table = typeof tableKey === 'string' ? specOf.get(tableKey) : undefined;
    if (
      typeof rowKey !== 'string' ||
      !isA(table, 'tbl') ||
      table.children?.some((c) => isA(c, 'tr'))
    )
      return undefined;
    const row = read(rowKey, new Set(depthOf.keys()), depthOf.get(table.key)! + 1);
    if (!isA(row ?? undefined, 'tr')) return undefined;
    const append = (node: ReplicaElementSpec): ReplicaElementSpec => ({
      ...node,
      children:
        node.key === table.key
          ? [...(node.children ?? []), row!]
          : (node.children ?? []).map((c) => ('text' in c ? c : append(c))),
    });
    spec = append(spec);
    collect(spec, 0, null);
    placed?.add(row!.key);
    into = survivingCell(room, o.key, o.to, specOf);
    return into;
  };
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
    if (round.length === 0) {
      for (const o of followMoveSiblings(rest).sort((a, b) => byKey(a.key, b.key))) {
        const into = listings && cellOf(o);
        if (into)
          candidates.push({ key: o.key, from: into, index: Infinity, clock: o.clock, to: o.to });
        else stranded?.add(o.removed);
      }
      break;
    }
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
    before?: string;
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
      before: c.before,
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
      .filter((x) => x.after === undefined && x.before === undefined)
      .sort((x, y) => y.clock - x.clock || x.index - y.index || byKey(x.node.key, y.node.key));
    for (const { node, index } of order) placed.splice(Math.min(index, placed.length), 0, node);
    // A node that follows a join goes right after its anchor, after the nodes before it that
    // follow the same anchor, in the order the rounds above placed them.
    // A node that goes before its anchor goes right before it, so such nodes keep their order.
    const following = extra
      .filter((x) => x.after !== undefined || x.before !== undefined)
      .sort((x, y) => x.seq - y.seq);
    const lastAfter = new Map<string, string>();
    for (const { node, after, before } of following) {
      if (after === undefined) {
        const at = placed.findIndex((k) => k.key === before);
        placed.splice(at < 0 ? placed.length : at, 0, node);
        continue;
      }
      const anchor = lastAfter.get(after) ?? after;
      const at = placed.findIndex((k) => k.key === anchor);
      placed.splice(at < 0 ? placed.length : at + 1, 0, node);
      lastAfter.set(after, node.key);
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
