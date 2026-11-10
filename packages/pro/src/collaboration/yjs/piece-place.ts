/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The place of a peer's format pieces next to runs that this replica listed at the same place
// (openspec change `collaboration-text-lineage`, "Character identity").
//
// A format splits a run: the run keeps the text before the cut, and the new runs (its pieces,
// `sf` in `text-lineage.ts`) go right after it. A join or a move of this replica can list runs
// right after the same run at the same time: the join of a paragraph that an Enter cut lists the
// runs of the new paragraph again. Yjs orders two lists made at one place by client id, so when
// this replica's id is the smaller one, the pieces show after this replica's runs, and the
// formatted text moves to the end of the paragraph; when its id is the larger one, the pieces
// show before an Enter's copy of the text before them. The pieces hold the end of the split
// run's text, so they go right after it, and after the runs of this replica that copy text from
// before the pieces' first character (the part of an Enter's copy that the pieces did not take),
// but before the other runs this replica listed there.
//
// Only the replica whose runs took the place writes the move (a move is not idempotent), when
// the pieces arrive. Every other replica reads the moved list. A peer's undo of the format
// empties the pieces wherever they are listed, so the move needs no undo of its own.
//
// Everything read here comes from the room and is untrusted: each list walk is capped, and a
// move only takes keys that the list holds once.

import * as Y from 'yjs';
import { F, type Room } from './schema.ts';
import { LINEAGE } from './text-lineage.ts';
import { copyOf, rootOf } from './text-identity.ts';
import { placeOf } from './text-repair.ts';

/** The list items one call reads, over all changed lists and every walk. */
const MAX_LOOKS = 65_536;
/** The pieces one move takes. */
const MAX_PIECES = 64;

/** A move of a peer's pieces: each group goes right after the entry it names. */
interface PieceMove {
  readonly list: Y.Array<string>;
  readonly groups: readonly { readonly after: string; readonly keys: readonly string[] }[];
}

const sameId = (a: Y.ID | null, b: Y.ID | null): boolean =>
  a !== null && b !== null && a.client === b.client && a.clock === b.clock;

/** The key at `id` when `item` holds it, or undefined. */
function keyAt(item: Y.Item, id: Y.ID): string | undefined {
  const j = id.clock - item.id.clock;
  if (id.client !== item.id.client || j < 0 || j >= item.length) return undefined;
  const key: unknown = item.content.getContent()[j];
  return typeof key === 'string' ? key : undefined;
}

/** The record of `key` in the room, or undefined. */
function recordOf(room: Room, key: string): Y.Map<unknown> | undefined {
  const rec = room.nodes.get(key);
  return rec instanceof Y.Map ? (rec as Y.Map<unknown>) : undefined;
}

/** Where the first character that a run's text copies stands, or null for a run of new text. */
function startOf(room: Room, runKey: string): { text: unknown; at: number } | null {
  const run = recordOf(room, runKey);
  const runList = run?.get(F.children);
  if (!(runList instanceof Y.Array)) return null;
  for (const child of (runList as Y.Array<string>).toArray().slice(0, 8)) {
    const list = recordOf(room, child)?.get(F.children);
    if (!(list instanceof Y.Array)) continue;
    for (const key of (list as Y.Array<string>).toArray().slice(0, 8)) {
      const copy = copyOf(recordOf(room, key));
      if (copy) return placeOf(room.doc, rootOf(room.doc, copy.source(0)));
    }
  }
  return null;
}

/**
 * The move of the peer's pieces that start at `item`, or null. `item` is new in this change, and
 * its left neighbor when it was made is the entry of the run the pieces came from. Next to the
 * pieces, on one side, show only entries of this replica with the same left neighbor (Yjs
 * ordered the two by client id). The pieces go after those of the entries that copy text from
 * before the pieces' start, and before the others.
 */
function moveOf(
  room: Room,
  list: Y.Array<string>,
  item: Y.Item,
  me: number,
  step: () => boolean
): PieceMove | null {
  const origin = item.origin;
  if (!origin) return null;
  const isTied = (at: Y.Item) => at.id.client === me && sameId(at.origin, origin);
  // This replica's entries that Yjs listed before the pieces, in list order.
  const before: Y.Item[] = [];
  let left = item.left;
  for (; left !== null; left = left.left) {
    if (!step()) return null;
    if (left.deleted) continue;
    if (keyAt(left, origin) !== undefined) break;
    if (!isTied(left)) return null;
    before.unshift(left);
  }
  const from = left ? keyAt(left, origin) : undefined;
  if (!from) return null;
  // The peer's pieces of that run, in their order, from `item` on.
  const keys: string[] = [];
  const isPiece = (key: unknown): key is string =>
    typeof key === 'string' && recordOf(room, key)?.get(LINEAGE.splitFrom) === from;
  // The item that holds the last piece, and whether the pieces end with it.
  let last = item;
  let ends = true;
  for (let at: Y.Item | null = item; at !== null; at = at.right) {
    if (!step()) return null;
    if (at.deleted) continue;
    if (at.id.client !== item.id.client) break;
    const content = at.content.getContent();
    const end = content.findIndex((key) => !isPiece(key));
    if (end === 0) break;
    keys.push(...(content.slice(0, end < 0 ? content.length : end) as string[]));
    if (keys.length > MAX_PIECES) return null;
    last = at;
    ends = end < 0;
    if (!ends) break;
  }
  // This replica's entries that Yjs listed after the pieces.
  const after: Y.Item[] = [];
  for (let at = last.right; ends && at !== null && before.length === 0; at = at.right) {
    if (!step()) return null;
    if (at.deleted) continue;
    if (!isTied(at)) break;
    after.push(at);
  }
  const tied = [...before, ...after].flatMap((at) =>
    at.content.getContent().filter((key): key is string => typeof key === 'string')
  );
  // Two splits of one run at once are the split lineage's (`dropLosingSplits`).
  if (keys.length === 0 || tied.length === 0 || tied.some(isPiece)) return null;
  // Each piece goes after the entries that copy text from before its own start.
  const current = before.length > 0 ? tied[tied.length - 1]! : from;
  const starts = new Map<string, ReturnType<typeof startOf>>();
  const startOfTied = (key: string) => {
    if (!starts.has(key)) starts.set(key, startOf(room, key));
    return starts.get(key)!;
  };
  const groups: { after: string; keys: string[] }[] = [];
  for (const key of keys) {
    const start = startOf(room, key);
    if (!start) return null;
    let anchor = from;
    for (const t of tied) {
      if (!step()) return null;
      const other = startOfTied(t);
      if (!other || other.text !== start.text || other.at >= start.at) break;
      anchor = t;
    }
    const group = groups[groups.length - 1];
    if (group?.after === anchor) group.keys.push(key);
    else groups.push({ after: anchor, keys: [key] });
  }
  return groups.every((g) => g.after === current) ? null : { list, groups };
}

/**
 * Move the peer's pieces that `transaction` listed after this replica's runs at the same place
 * to right after the run they came from. Call after a peer's change; writes its own transaction
 * with `origin` when there is a move.
 */
export function placePeerPieces(room: Room, transaction: Y.Transaction, origin: unknown): void {
  const me = room.doc.clientID;
  const moves: PieceMove[] = [];
  let looks = MAX_LOOKS;
  const step = () => (looks -= 1) >= 0;
  for (const type of transaction.changed.keys()) {
    if (!(type instanceof Y.Array)) continue;
    const rec = type.parent;
    if (!(rec instanceof Y.Map) || rec.get(F.children) !== type) continue;
    const list = type as Y.Array<string>;
    for (let item = list._start; item !== null; item = item.right) {
      if (!step()) return;
      if (item.deleted || item.id.client === me) continue;
      if (item.id.clock < (transaction.beforeState.get(item.id.client) ?? 0)) continue;
      const move = moveOf(room, list, item, me, step);
      if (move) moves.push(move);
    }
  }
  if (moves.length === 0) return;
  room.doc.transact(() => {
    for (const { list, groups } of moves) {
      const keys = groups.flatMap((g) => g.keys);
      const listed = list.toArray();
      const at = keys.map((k) => listed.indexOf(k));
      const once = (k: string) =>
        listed.indexOf(k) >= 0 && listed.indexOf(k) === listed.lastIndexOf(k);
      if (!keys.every(once) || !groups.every((g) => once(g.after))) continue;
      for (const i of [...at].sort((x, y) => y - x)) list.delete(i, 1);
      for (const { after, keys: group } of groups)
        list.insert(list.toArray().indexOf(after) + 1, [...group]);
    }
  }, origin);
}
