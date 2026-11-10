/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Text that an undo or a redo moves back (openspec change `collaboration-text-lineage`, "Undo").
//
// A split (a format, an Enter) moves text: it deletes the characters after the cut from the run
// and inserts copies of them into new runs (`text-identity.ts`). An undo of the split deletes
// the copies and puts the old characters back as new characters. A redo does the reverse. Yjs
// sees a delete and an insert, so it puts back each character the step deleted. It does not
// know that a peer deleted the text after it moved:
//
// - A peer's delete before the step: the peer deleted the copies (or its delete reached them,
//   `text-repair.ts`). The step puts the text back in its old place, and the peer's delete is
//   lost. So the replica that undoes or redoes reads, before the step, which moved text still
//   shows anywhere (`liveRoots`). After the step, a character it put back for moved text that no
//   shown character held goes again (`dropDeadMoves`). That delete is not part of the step: a
//   later redo reads it as the peer's (`withoutTakenRedo` in `undo-splits.ts`).
// - A peer's delete at the same time as the step: the peer deleted the old characters or the
//   copies, and the characters the step put back are new. So the replica that deleted the text
//   deletes them too, when their restore marks arrive (`dropDeletedRestores`). Only that
//   replica knows its delete; a delete is idempotent, so every such replica may write it.
//
// A step moved a character when it also inserted a character of the same identity (`rootOf`):
// the copy of an undone split, or the character that the undo put back for a redo. A step that
// deleted a character without a copy (a delete, then its undo) is not a move, and its undo puts
// the character back as before.
//
// The reverse case: an undo of typing deletes text that a peer's split copies at the same time.
// The copy is new, so the undo's delete does not reach it. The undoing replica records the
// characters its step deleted (`deletedByStep`) with the deletes it typed, and its delete then
// reaches the copy when it arrives (`text-repair.ts`).
//
// Everything read here comes from the room and is untrusted: every walk is bounded.

import * as Y from 'yjs';
import { fromEnter } from './copy-order.ts';
import { F, type Room } from './schema.ts';
import {
  COPIES_KEY,
  copyOf,
  idKey,
  itemOf,
  MAX_RESTORED,
  recordOfItem,
  RESTORED_KEY,
  RESTORED_STEP_KEY,
  restoredOf,
  restoredPlaceOf,
  type RestoredPlace,
  rootOf,
  deletionRoot,
} from './text-identity.ts';
import { deleteChars, forEachItem, holderItem, type IdSet } from './yjs-items.ts';

/** The characters one read looks at; over the cap (only a hostile room), it changes nothing. */
const MAX_CHARS = 1 << 20;
/** The copy records one read looks at. */
const MAX_COPIES = 4_096;

/** Whether a text is the text of a node record in `nodes`. */
function isNodeText(room: Room, text: unknown): text is Y.Text {
  if (!(text instanceof Y.Text)) return false;
  const rec = text.parent;
  return rec instanceof Y.Map && holderItem(rec)?.parent === (room.nodes as unknown);
}

/**
 * The identities (`rootOf`, as `idKey`) of the text characters a step inserted, or null over
 * the cap. A step that put back a character of one of these identities moved it.
 */
export function movedRoots(room: Room, insertions: IdSet): Set<string> | null {
  const doc = room.doc;
  const out = new Set<string>();
  const memo = new Map<string, Y.ID>();
  let count = 0;
  let over = false;
  forEachItem(doc, insertions, (item) => {
    if (over || !item.countable || !isNodeText(room, item.parent)) return;
    for (let j = 0; j < item.length; j += 1) {
      if ((count += 1) > MAX_CHARS) {
        over = true;
        return;
      }
      const root = rootOf(doc, Y.createID(item.id.client, item.id.clock + j), memo);
      out.add(idKey(root.client, root.clock));
    }
  });
  return over ? null : out;
}

/** The original split boundary of each character that a step moved into a copy. */
export function movedStarts(room: Room, insertions: IdSet): Map<string, RestoredPlace> {
  const starts = new Map<string, RestoredPlace>();
  let remaining = MAX_CHARS;
  const memo = new Map<string, Y.ID>();
  forEachItem(room.doc, insertions, (item) => {
    if (remaining <= 0 || !item.countable || !isNodeText(room, item.parent)) return;
    const copy = copyOf(recordOfItem(item));
    for (let j = 0; j < item.length && --remaining >= 0; j += 1) {
      const id = Y.createID(item.id.client, item.id.clock + j);
      const start =
        copy &&
        id.client === copy.client &&
        id.clock >= copy.clock &&
        id.clock < copy.clock + copy.length
          ? {
              start: copy.source(0),
              enter: fromEnter(room.doc, copy),
              client: copy.client,
              key: String(recordOfItem(item)?.get(F.key) ?? ''),
            }
          : restoredPlaceOf(room.doc, id);
      if (!start) continue;
      const root = rootOf(room.doc, id, memo);
      starts.set(idKey(root.client, root.clock), start);
    }
  });
  return starts;
}

/** Whether the read shows the record that holds `text` (`shows(part, key)`). */
function shown(text: Y.Text, shows: (part: string, key: string) => boolean): boolean {
  const rec = text.parent as Y.Map<unknown>;
  const key = holderItem(rec)?.parentSub;
  const part = rec.get(F.part);
  return typeof key === 'string' && typeof part === 'string' && shows(part, key);
}

/**
 * Of `roots`, the identities that a live character holds now, in a text the read shows
 * (`shows`): the character first typed, a character of a copy (`COPIES_KEY`), or a character an
 * undo or a redo put back (`RESTORED_KEY`). A copy that the read drops (the pieces of a split
 * that lost to a peer's split, `text-lineage.ts`) holds nothing. Null over the caps.
 */
export function liveRoots(
  room: Room,
  roots: ReadonlySet<string>,
  shows: (part: string, key: string) => boolean
): Set<string> | null {
  const doc = room.doc;
  const out = new Set<string>();
  if (roots.size === 0) return out;
  const memo = new Map<string, Y.ID>();
  let looked = 0;
  const visit = (text: Y.Text, holds: (id: Y.ID) => boolean): boolean => {
    if (!shown(text, shows)) return true;
    for (let it = text._start; it !== null; it = it.right) {
      if (it.deleted || !it.countable) continue;
      for (let j = 0; j < it.length; j += 1) {
        if ((looked += 1) > MAX_CHARS) return false;
        const id = Y.createID(it.id.client, it.id.clock + j);
        if (!holds(id)) continue;
        const root = rootOf(doc, id, memo);
        const key = idKey(root.client, root.clock);
        if (roots.has(key)) out.add(key);
      }
    }
    return true;
  };
  for (const key of roots) {
    const [client, clock] = key.split(':').map(Number) as [number, number];
    const item = itemOf(doc, Y.createID(client, clock));
    if (item && !item.deleted && isNodeText(room, item.parent) && shown(item.parent, shows))
      out.add(key);
  }
  let copies = 0;
  for (const key of doc.getMap(COPIES_KEY).keys()) {
    if ((copies += 1) > MAX_COPIES) return null;
    const rec = room.nodes.get(key);
    const text = rec instanceof Y.Map ? rec.get(F.text) : undefined;
    if (!copyOf(rec) || !isNodeText(room, text)) continue;
    if (!visit(text, () => true)) return null;
  }
  const texts = new Set<Y.Text>();
  let marks = 0;
  for (const [key, value] of doc.getMap(RESTORED_KEY).entries()) {
    if ((marks += 1) > MAX_RESTORED) return null;
    const m = /^(\d{1,16}):(\d{1,16})$/.exec(key);
    if (!m || !Array.isArray(value)) continue;
    const item = itemOf(doc, Y.createID(Number(m[1]), Number(m[2])));
    if (item && isNodeText(room, item.parent)) texts.add(item.parent);
  }
  for (const text of texts) if (!visit(text, (id) => restoredOf(doc, id) !== null)) return null;
  return out;
}

/**
 * After an undo or a redo of a step that deleted `deletions`: each character it put back whose
 * identity is in `moved` (the step moved it) and not in `alive` (no character held it before
 * the step: a peer deleted it) goes again. Call inside a transaction. Returns whether it wrote.
 */
export function dropDeadMoves(
  room: Room,
  deletions: IdSet,
  moved: ReadonlySet<string>,
  alive: ReadonlySet<string>
): boolean {
  const doc = room.doc;
  const drop = new Map<Y.Text, Set<string>>();
  const memo = new Map<string, Y.ID>();
  let count = 0;
  forEachItem(doc, deletions, (old) => {
    if (count > MAX_CHARS || !old.countable || old.redone === null) return;
    const text = old.parent;
    if (!isNodeText(room, text)) return;
    const now = itemOf(doc, old.redone);
    if (!now || now.deleted || now.parent !== text) return;
    const length = Math.min(old.length, now.length);
    for (let j = 0; j < length; j += 1) {
      if ((count += 1) > MAX_CHARS) return;
      const root = rootOf(doc, Y.createID(old.id.client, old.id.clock + j), memo);
      const key = idKey(root.client, root.clock);
      if (!moved.has(key) || alive.has(key)) continue;
      const ids = drop.get(text) ?? new Set<string>();
      ids.add(idKey(now.id.client, now.id.clock + j));
      drop.set(text, ids);
    }
  });
  for (const [text, ids] of drop) deleteChars(text, ids);
  return drop.size > 0;
}

/**
 * After a peer's restore marks `keys` arrive (`RESTORED_KEY`): each live character they mark
 * that stands for a character the peer's own split moved (a source of the peer's copy, or a
 * character of that copy after the peer undid the split), whose identity this replica deleted
 * (`deleted`, roots through copies only), goes. The peer undid a
 * split of text that this replica deleted at the same time. Call inside a transaction. Returns
 * whether it wrote.
 */
export function dropDeletedRestores(
  room: Room,
  keys: Iterable<string | null>,
  deleted: ReadonlySet<string>,
  rejected?: (id: string, source: string) => void,
  deletedAt?: ReadonlyMap<string, number>
): boolean {
  if (deleted.size === 0) return false;
  const doc = room.doc;
  const marks = doc.getMap(RESTORED_KEY);
  // Per client: the characters its splits moved (the sources of its copies), each character
  // of its copies with the character it copies, and the characters its restore marks stand for.
  // Read once per call.
  const moves = new Map<number, Moves | null>();
  const movesOf = (client: number): Moves | null => {
    if (moves.has(client)) return moves.get(client)!;
    const result = readMoves(room, client);
    moves.set(client, result);
    return result;
  };
  const memo = new Map<string, Y.ID>();
  // The identities (`rootOf` through restores too) of what this replica deleted: it can have
  // deleted the characters that the peer's undo put back, and the redo puts back the copy.
  // Read once, and only when a mark needs it.
  let identities: Set<string> | null = null;
  const full = new Map<string, Y.ID>();
  const deletedIdentity = (target: Y.ID): boolean => {
    if (!identities) {
      identities = new Set();
      let n = 0;
      for (const key of deleted) {
        if ((n += 1) > MAX_CHARS) break;
        const [c, k] = key.split(':').map(Number) as [number, number];
        const r = rootOf(doc, Y.createID(c, k), full);
        identities.add(idKey(r.client, r.clock));
      }
    }
    const r = rootOf(doc, target, full);
    return identities.has(idKey(r.client, r.clock));
  };
  const drop = new Map<Y.Text, Set<string>>();
  for (const key of keys) {
    const m = key === null ? null : /^(\d{1,16}):(\d{1,16})$/.exec(key);
    const value: unknown = key === null ? undefined : marks.get(key);
    if (!m || !Array.isArray(value) || value.length !== 3) continue;
    const [tc, tk, n] = value as unknown[];
    if (!Number.isSafeInteger(tc) || !Number.isSafeInteger(tk) || !Number.isSafeInteger(n))
      continue;
    const client = Number(m[1]);
    const first = Number(m[2]);
    if (client === doc.clientID) continue;
    const by = movesOf(client);
    if (!by) continue;
    const length = Math.min(n as number, MAX_RESTORED);
    for (let k = 0; k < length; k += 1) {
      const target = Y.createID(tc as number, (tk as number) + k);
      const at = idKey(target.client, target.clock);
      // An undo of the split puts back a source; a redo after it puts back a copy's character.
      const source = by.copies.get(at);
      const step: unknown = doc.getMap(RESTORED_STEP_KEY).get(idKey(client, first + k));
      const precise =
        Array.isArray(step) &&
        (step.length === 3 || step.length === 4) &&
        Number.isSafeInteger(step[0]) &&
        step[0] >= 0 &&
        Number.isSafeInteger(step[1]) &&
        step[1] >= 0 &&
        typeof step[2] === 'boolean';
      let rejectedSource = '';
      if (precise) {
        const root = deletionRoot(doc, Y.createID(step[0], step[1]), memo);
        const rootKey = idKey(root.client, root.clock);
        if (!deleted.has(rootKey)) continue;
        if (!step[2]) {
          const vector = step[3];
          const deletionClock = deletedAt?.get(rootKey);
          if (!Array.isArray(vector) || vector.length > 8192 || deletionClock === undefined)
            continue;
          let sawDeletion = false;
          for (let i = 0; i + 1 < vector.length; i += 2)
            if (
              vector[i] === doc.clientID &&
              Number.isSafeInteger(vector[i + 1]) &&
              vector[i + 1] > deletionClock
            )
              sawDeletion = true;
          if (sawDeletion) continue;
        }
        rejectedSource = rootKey;
      } else {
        if (!by.sources.has(at) && !(source !== undefined && by.restored.has(source))) continue;
        const root = rootOf(doc, target, memo, false);
        if (!deleted.has(idKey(root.client, root.clock)) && !deletedIdentity(target)) continue;
      }
      rejected?.(idKey(client, first + k), rejectedSource);
      const item = itemOf(doc, Y.createID(client, first + k));
      if (!item || item.deleted || !isNodeText(room, item.parent)) continue;
      const ids = drop.get(item.parent) ?? new Set<string>();
      ids.add(idKey(client, first + k));
      drop.set(item.parent, ids);
    }
  }
  for (const [text, ids] of drop) deleteChars(text, ids);
  return drop.size > 0;
}

/** What one client's splits and restores wrote (`dropDeletedRestores`). */
interface Moves {
  /** The characters its copies copy. */
  readonly sources: Set<string>;
  /** Each character of its copies → the character it copies. */
  readonly copies: Map<string, string>;
  /** The characters its restore marks stand for. */
  readonly restored: Set<string>;
}

/** The moves of `client`, or null over the caps. */
function readMoves(room: Room, client: number): Moves | null {
  const doc = room.doc;
  const out: Moves = { sources: new Set(), copies: new Map(), restored: new Set() };
  let copies = 0;
  let chars = 0;
  for (const key of doc.getMap(COPIES_KEY).keys()) {
    if ((copies += 1) > MAX_COPIES) return null;
    const copy = copyOf(room.nodes.get(key));
    if (!copy || copy.client !== client) continue;
    for (let j = 0; j < copy.length; j += 1) {
      if ((chars += 1) > MAX_CHARS) return null;
      const s = copy.source(j);
      const source = idKey(s.client, s.clock);
      out.sources.add(source);
      out.copies.set(idKey(copy.client, copy.clock + j), source);
    }
  }
  const prefix = `${client}:`;
  let marks = 0;
  for (const [key, value] of doc.getMap(RESTORED_KEY).entries()) {
    if (!key.startsWith(prefix) || !Array.isArray(value) || value.length !== 3) continue;
    const [c, k, n] = value as unknown[];
    if (!Number.isSafeInteger(c) || !Number.isSafeInteger(k) || !Number.isSafeInteger(n)) continue;
    for (let j = 0; j < Math.min(n as number, MAX_RESTORED); j += 1) {
      if ((marks += 1) > MAX_CHARS) return null;
      out.restored.add(idKey(c as number, (k as number) + j));
    }
  }
  return out;
}

/** The live text characters of `insertions` (as `idKey`), or null over the cap. */
function latestItem(doc: Y.Doc, id: Y.ID): Y.Item | null {
  let at = id;
  for (let n = 0; n < 32; n += 1) {
    const item = itemOf(doc, at);
    if (!item || !item.redone) return item;
    at = Y.createID(item.redone.client, item.redone.clock + at.clock - item.id.clock);
  }
  return null;
}

export function liveInserted(room: Room, insertions: IdSet): Set<string> | null {
  const out = new Set<string>();
  let over = false;
  forEachItem(room.doc, insertions, (item) => {
    if (over || !item.countable || !isNodeText(room, item.parent)) return;
    for (let j = 0; j < item.length; j += 1) {
      if (out.size >= MAX_CHARS) {
        over = true;
        return;
      }
      const id = Y.createID(item.id.client, item.id.clock + j);
      const latest = latestItem(room.doc, id);
      if (latest && !latest.deleted) out.add(idKey(id.client, id.clock));
    }
  });
  return over ? null : out;
}

/**
 * After an undo or a redo of a step that inserted `insertions` and deleted `deletions`: the
 * roots (`rootOf`, through copies only and through restores too) of the characters the step
 * deleted that were live before it (`before`, from `liveInserted`). A redo that deletes a
 * character an undo put back deletes the character it stands for too. A character of a split's copy, and a character whose identity the step put back
 * (it moved the text back), do not count. This replica's delete then reaches a peer's copy of
 * them, as a delete that this replica typed does (`text-repair.ts`).
 */
export function deletedByStep(
  room: Room,
  ranges: { insertions: IdSet; deletions: IdSet },
  before: ReadonlySet<string>
): Set<string> {
  const doc = room.doc;
  const out = new Set<string>();
  if (before.size === 0) return out;
  const full = new Map<string, Y.ID>();
  const copies = new Map<string, Y.ID>();
  const putBack = new Set<string>();
  const typed = new Set<string>();
  let inspected = 0;
  forEachItem(doc, ranges.insertions, (item) => {
    if (!item.countable || !isNodeText(room, item.parent)) return;
    for (let j = 0; j < item.length && inspected++ < MAX_CHARS; j += 1) {
      const id = Y.createID(item.id.client, item.id.clock + j);
      if (Y.compareIDs(rootOf(doc, id, full), id)) typed.add(idKey(id.client, id.clock));
    }
  });
  let count = 0;
  forEachItem(doc, ranges.deletions, (old) => {
    if (count > MAX_CHARS || !old.countable || old.redone === null) return;
    if (!isNodeText(room, old.parent)) return;
    for (let j = 0; j < old.length; j += 1, count += 1) {
      const r = rootOf(doc, Y.createID(old.id.client, old.id.clock + j), full);
      putBack.add(idKey(r.client, r.clock));
    }
  });
  forEachItem(doc, ranges.insertions, (item) => {
    if (!item.deleted || !item.countable || !isNodeText(room, item.parent)) return;
    // The characters of a split's copy move text; the undo of the split moves it back.
    const copy = copyOf(recordOfItem(item));
    for (let j = 0; j < item.length; j += 1) {
      const id = Y.createID(item.id.client, item.id.clock + j);
      if (!before.has(idKey(id.client, id.clock))) continue;
      if (latestItem(doc, id)?.deleted !== true) continue;
      const r = rootOf(doc, id, full);
      if (copy && copy.client === id.client && id.clock >= copy.clock) {
        if (id.clock < copy.clock + copy.length && !typed.has(idKey(r.client, r.clock))) continue;
      }
      if (putBack.has(idKey(r.client, r.clock))) continue;
      // A peer's copy can copy the character itself or the character it stands for.
      const root = deletionRoot(doc, id, copies);
      out.add(idKey(root.client, root.clock));
      out.add(idKey(r.client, r.clock));
    }
  });
  return out;
}

/** An undo of typing also removes known copies of those typed characters. */
export function includeCopiedTyping(room: Room, ranges: { insertions: IdSet }): void {
  const roots = new Set<string>();
  const memo = new Map<string, Y.ID>();
  let remaining = MAX_CHARS;
  forEachItem(room.doc, ranges.insertions, (item) => {
    if (!item.countable || !isNodeText(room, item.parent)) return;
    for (let j = 0; j < item.length && remaining-- > 0; j += 1) {
      const id = Y.createID(item.id.client, item.id.clock + j);
      if (Y.compareIDs(rootOf(room.doc, id, memo), id)) roots.add(idKey(id.client, id.clock));
    }
  });
  if (remaining <= 0 || roots.size === 0) return;
  const added = Y.createDeleteSet();
  for (const record of room.nodes.values()) {
    if (--remaining < 0) return;
    const text = record instanceof Y.Map ? record.get(F.text) : null;
    if (!(text instanceof Y.Text)) continue;
    for (let item = text._start; item && --remaining >= 0; item = item.right) {
      if (item.deleted || !item.countable) continue;
      for (let j = 0; j < item.length && --remaining >= 0; j += 1) {
        const id = Y.createID(item.id.client, item.id.clock + j);
        const root = rootOf(room.doc, id, memo);
        if (!roots.has(idKey(root.client, root.clock))) continue;
        const entries = added.clients.get(id.client) ?? [];
        entries.push({ clock: id.clock, len: 1 } as never);
        added.clients.set(id.client, entries);
      }
    }
    if (remaining < 0) return;
  }
  ranges.insertions = Y.mergeDeleteSets([ranges.insertions, added]);
}
