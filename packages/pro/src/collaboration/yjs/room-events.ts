/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import * as Y from 'yjs';
import { F, type Room } from './schema.ts';
import { PIECES_KEY, MOVES_KEY, SPLITS_KEY } from './text-lineage.ts';
import { UNDONE_KEY } from './undo-unlisted.ts';

/** Package parts whose contents or placement changed in a room transaction. */
export function touchedParts(
  room: Room,
  transaction: Y.Transaction,
  knownParts: ReadonlyMap<string, string>
): Set<string> {
  const out = new Set<string>();
  const parts: unknown = room.parts;
  const binaries: unknown = room.binaries;
  const nodes: unknown = room.nodes;
  const doc = room.doc;
  const lineage: unknown[] = [doc.getMap(PIECES_KEY), doc.getMap(MOVES_KEY)];
  const splits: unknown = doc.getMap(SPLITS_KEY);
  const undone: unknown = doc.getMap(UNDONE_KEY);
  for (const [changed, keys] of transaction.changed) {
    const type: unknown = changed;
    if (type === parts || type === binaries) {
      for (const k of keys) if (k !== null) out.add(k);
      continue;
    }
    // A lineage entry alone (`text-lineage.ts`) changes how its part reads: the piece and move
    // indexes hold the part name; a split index entry starts with the key of its run.
    if (lineage.includes(type)) {
      for (const k of keys) {
        const part = k === null ? undefined : (changed as unknown as Y.Map<unknown>).get(k);
        if (typeof part === 'string') out.add(part);
      }
      continue;
    }
    // An undone listing alone (`undo-unlisted.ts`) changes whether its node shows: its value
    // names the part, else the node's record does.
    if (type === undone) {
      for (const k of keys) {
        if (k === null) continue;
        const value: unknown = room.doc.getMap(UNDONE_KEY).get(k);
        const named =
          typeof value === 'object' && value !== null ? (value as { p?: unknown }).p : undefined;
        const rec = room.nodes.get(k);
        const part =
          typeof named === 'string' ? named : rec instanceof Y.Map ? rec.get(F.part) : undefined;
        if (typeof part === 'string') out.add(part);
      }
      continue;
    }
    if (type === splits) {
      for (const k of keys) {
        const rec = k === null ? undefined : room.nodes.get(k.split('\u0000')[0]!);
        const part = rec instanceof Y.Map ? rec.get(F.part) : undefined;
        if (typeof part === 'string') out.add(part);
      }
      continue;
    }
    if (type === nodes) {
      for (const k of keys) {
        if (k === null) continue;
        const rec = room.nodes.get(k);
        const part = rec instanceof Y.Map ? rec.get(F.part) : undefined;
        if (typeof part === 'string') {
          out.add(part);
          continue;
        }
        // A deleted record (an undo removed it): its part is the one this replica keeps it in.
        const known = knownParts.get(k);
        if (known !== undefined) out.add(known);
      }
      continue;
    }
    // An attribute map, a child list, or a text inside a record: climb to the record.
    let t: unknown = type;
    while (t instanceof Y.AbstractType && t.parent !== nodes) t = t.parent;
    const part = t instanceof Y.Map ? t.get(F.part) : undefined;
    if (typeof part === 'string') out.add(part);
  }
  return out;
}
