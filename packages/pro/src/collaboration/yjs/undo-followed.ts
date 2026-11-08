/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A peer's run that followed this replica's Enter (the "Follow an Enter" rule, `followEnter` in
// `text-repair.ts`) goes back with an undo of the Enter. The repair moved the run's list entry
// from the peer's paragraph (`from`) into the Enter's new paragraph (`to`); the session ties
// that move to the Enter's history step (`tieFollowedRuns`), so the undo deletes the new entry
// and puts the old one back. A step keeps the runs it moved in its meta (`FOLLOWED`), and the
// step that an undo or a redo pushes takes them over (`carryFollowed`).
//
// A peer can move the run on after the repair (an Enter inside it, a delete): the new paragraph
// no longer lists it. An undo then leaves the run where the peer put it (`checkFollowed`): it
// does not put the old entry back, or the run would show twice.

import * as Y from 'yjs';
import type { IdSet } from './yjs-items.ts';

/** The meta field of a history step: each run a repair moved as part of the step. */
export const FOLLOWED = 'followed';

/** The list entries one history step walks for one run; only a hostile room has more. */
const MAX_ENTRIES = 1 << 16;

/** A run that a repair moved after this replica's Enter. */
export interface FollowedRun {
  /** The first character of the Enter's copy: the step that inserted it made the Enter. */
  readonly anchor: Y.ID;
  /** The list entry the move deleted, in `from`. */
  readonly gone: Y.ID;
  /** The list entry the move inserted, in `to`. */
  readonly added: Y.ID;
  readonly run: string;
  readonly from: Y.Array<string>;
  readonly to: Y.Array<string>;
}

interface Lists {
  readonly from: Y.Array<string>;
  readonly to: Y.Array<string>;
}

/** The parts of a Yjs history step that this module reads and changes. */
interface Step {
  insertions: IdSet;
  deletions: IdSet;
  readonly meta: Map<string, unknown>;
}

const one = (id: Y.ID): IdSet => {
  const set = Y.createDeleteSet();
  set.clients.set(id.client, [{ clock: id.clock, len: 1 }] as never);
  return set;
};

const followedOf = (step: Pick<Step, 'meta'>): Map<string, Lists> | undefined => {
  const value = step.meta.get(FOLLOWED);
  return value instanceof Map ? (value as Map<string, Lists>) : undefined;
};

/** Ties each run's move to the history step that made the Enter it followed. */
export function tieFollowedRuns(steps: readonly Step[], runs: readonly FollowedRun[]): void {
  for (const r of runs)
    for (const step of steps) {
      if (!Y.isDeleted(step.insertions, r.anchor)) continue;
      step.insertions = Y.mergeDeleteSets([step.insertions, one(r.added)]);
      step.deletions = Y.mergeDeleteSets([step.deletions, one(r.gone)]);
      let followed = followedOf(step);
      if (!followed) step.meta.set(FOLLOWED, (followed = new Map()));
      followed.set(r.run, { from: r.from, to: r.to });
    }
}

/** The step that an undo or a redo of `from` pushes keeps the runs `from` moved. */
export function carryFollowed(from: Pick<Step, 'meta'>, to: Pick<Step, 'meta'>): void {
  const followed = followedOf(from);
  if (followed) to.meta.set(FOLLOWED, new Map(followed));
}

/**
 * Whether a list holds a live entry `key`. A list longer than `MAX_ENTRIES` counts as holding
 * it: the undo then puts the old entry back as usual, so the run is never listed nowhere.
 */
export const lists = (list: Y.Array<string>, key: string): boolean => {
  let walked = 0;
  for (let item = list._start; item !== null; item = item.right) {
    if ((walked += 1) > MAX_ENTRIES) return true;
    if (!item.deleted && item.countable && item.content.getContent().includes(key)) return true;
  }
  return false;
};

/** The ids of the deleted entries `key` of a list. */
function deletedEntries(list: Y.Array<string>, key: string): Y.ID[] {
  const out: Y.ID[] = [];
  let walked = 0;
  for (let item = list._start; item !== null && walked < MAX_ENTRIES; item = item.right) {
    walked += 1;
    if (!item.deleted || !item.countable) continue;
    const content = item.content.getContent() as unknown[];
    for (let i = 0; i < content.length; i += 1)
      if (content[i] === key) out.push(Y.createID(item.id.client, item.id.clock + i));
  }
  return out;
}

/** A delete set without the single ids `drop`. */
function without(set: IdSet, drop: readonly Y.ID[]): IdSet {
  if (drop.length === 0) return set;
  const out = Y.createDeleteSet();
  for (const [client, ranges] of set.clients) {
    const gone = drop.filter((id) => id.client === client).map((id) => id.clock);
    const next: { clock: number; len: number }[] = [];
    for (const { clock, len } of ranges as { clock: number; len: number }[]) {
      let at = clock;
      for (const c of gone.filter((c) => c >= clock && c < clock + len).sort((a, b) => a - b)) {
        if (c > at) next.push({ clock: at, len: c - at });
        at = c + 1;
      }
      if (at < clock + len) next.push({ clock: at, len: clock + len - at });
    }
    if (next.length > 0) out.clients.set(client, next as never);
  }
  return out;
}

/**
 * Before an undo of `step`: the runs it moved that the new paragraph still lists, which the
 * undo lists at their old place again (`noteUndoneListings`). For a run that a peer moved on,
 * the step no longer puts the old entry back.
 */
export function checkFollowed(step: Step): ReadonlySet<string> {
  const followed = followedOf(step);
  const out = new Set<string>();
  if (!followed) return out;
  for (const [run, { from, to }] of followed) {
    if (lists(to, run)) {
      out.add(run);
      continue;
    }
    step.deletions = without(step.deletions, deletedEntries(from, run));
    followed.delete(run);
  }
  return out;
}
