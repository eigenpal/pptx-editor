/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// What the removals of records saw (`rw.<client>`, written by `markLineage` in
// `text-lineage.ts`). Everything read here comes from the room and is untrusted: a value of a
// wrong type is ignored, and a snapshot larger than any writer makes counts as all keys.

import * as Y from 'yjs';
import { F, type Room } from './schema.ts';
import { isRemovalField, isRemoved, MAX_REMOVERS, MAX_SNAPSHOT } from './text-lineage.ts';

/** An ancestor's removal snapshot for a read: its keys, `all` when too large, null when none. */
export type Snapshot = ReadonlySet<string> | 'all' | null;

/**
 * The keys that any removal of a record saw: the union of every writer's snapshot. A snapshot
 * or a count of writers larger than any writer makes reads as `all`; a value of a wrong type
 * is ignored.
 */
export function removalSnapshot(rec: unknown, seen?: ReadonlyMap<number, number>): Snapshot {
  if (!(rec instanceof Y.Map)) return null;
  let out: Set<string> | null = null;
  let removers = 0;
  for (const [field, list] of (rec as Y.Map<unknown>).entries()) {
    if (!isRemovalField(field) || !Array.isArray(list)) continue;
    const item = rec._map.get(field);
    if (item && seen && item.id.clock < (seen.get(item.id.client) ?? 0)) continue;
    if (list.length > MAX_SNAPSHOT + 1 || (removers += 1) > MAX_REMOVERS) return 'all';
    out ??= new Set();
    for (const k of list as unknown[]) if (typeof k === 'string') out.add(k);
  }
  return out;
}

/** The most records `removalsSawAnywhere` looks at; only a hostile room holds more. */
const MAX_SCAN = 1 << 18;

/**
 * The keys that the removal snapshot of any record of the part holds, or `all` when a snapshot
 * is too large to read. A peer that removed a place where a read showed a node (a rescue, a
 * merged cell) saw the node there, though no list put it in that place.
 */
export function removalsSawAnywhere(
  room: Room,
  partName: string,
  seen?: ReadonlyMap<number, number>
): ReadonlySet<string> | 'all' {
  const out = new Set<string>();
  let looked = 0;
  for (const rec of room.nodes.values()) {
    if ((looked += 1) > MAX_SCAN) break;
    if (!(rec instanceof Y.Map) || rec.get(F.part) !== partName || !isRemoved(rec)) continue;
    const held = removalSnapshot(rec, seen);
    if (held === 'all') return 'all';
    for (const k of held ?? []) out.add(k);
  }
  return out;
}
