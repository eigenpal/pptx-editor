/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The characters an undo puts back (openspec change `collaboration-text-lineage`, "Undo").
//
// A `Y.UndoManager` undo puts each character its step deleted back as a new character, right
// before the old one's tombstone. When the step was a split, that character had moved into the
// step's own pieces; after the undo, a peer's later split or Enter may hold a copy of the same
// character (`text-identity.ts` marks copies, and a copy of a copy leads to the character first
// typed). The undo then shows the character twice, or in the run it left while its neighbours
// stay where the peer moved them ("Helworldlo rld", ["ACDE", "B"]).
//
// After an undo, the undoing replica repairs what it put back, by each character's identity:
//
// - Once: a character put back whose copy shows (in a run or paragraph the read shows) goes
//   again; the copy holds it.
// - Follow: the others go right after the copy of the nearest character before them that a
//   peer moved (the walk left stops at a character that shows where it is), so the text keeps
//   its order in the run or paragraph that now holds its neighbours.
//
// DOCX Editor ties moved text to its source characters
// (`packages/pro/src/collaboration/document/split-text-sources.ts` at adf0e178), so its undo moves
// nothing back; this pass reaches the same text after the fact. Only the undoing replica knows
// which characters its undo put back, so only it writes the repair, as part of the undo's own step:
// a redo reverts it with the undo. Everything read is bounded.

import * as Y from 'yjs';
import { F, type Room } from './schema.ts';
import {
  COPIES_KEY,
  copyOf,
  idKey,
  itemOf,
  markRestored,
  MAX_RESTORED,
  RESTORED_KEY,
  restoredOf,
  rootOf,
} from './text-identity.ts';
import { deleteChars, forEachItem, holderItem, liveIndex, type IdSet } from './yjs-items.ts';

/** The copies one repair reads; over the cap (only a hostile room), it repairs nothing. */
const MAX_COPIES = 4_096;
/** The characters a walk to the left looks at for one run of put-back characters. */
const MAX_WALK = 4_096;
/** The restore marks one put-back character follows to the character it stands for. */
const MAX_BACK = 32;

/**
 * The live characters of every copy (`COPIES_KEY`) in `parts` that the read shows, by the root
 * of the character each one copies; the first copy in map order holds a root. Characters in
 * `skip` (`idKey`) do not count. Null over `MAX_COPIES` copies (only a hostile room).
 */
function shownCopies(
  room: Room,
  parts: ReadonlySet<string>,
  shows: (part: string, key: string) => boolean,
  memo: Map<string, Y.ID>,
  skip: ReadonlySet<string> = new Set()
): Map<string, Char> | null {
  const doc = room.doc;
  const shown = new Map<string, Char>();
  let read = 0;
  for (const [key, part] of doc.getMap(COPIES_KEY).entries()) {
    if (typeof part !== 'string' || !parts.has(part) || !shows(part, key)) continue;
    if ((read += 1) > MAX_COPIES) return null;
    const rec = room.nodes.get(key);
    const text = rec instanceof Y.Map ? rec.get(F.text) : undefined;
    if (!copyOf(rec) || !(text instanceof Y.Text)) continue;
    for (let it = text._start; it !== null; it = it.right) {
      if (it.deleted || !it.countable) continue;
      for (let j = 0; j < it.length; j += 1) {
        const own = idKey(it.id.client, it.id.clock + j);
        if (skip.has(own)) continue;
        const root = rootOf(doc, Y.createID(it.id.client, it.id.clock + j), memo);
        const id = idKey(root.client, root.clock);
        if (id !== own && !shown.has(id)) shown.set(id, { text, item: it, j });
      }
    }
  }
  return shown;
}

/** A text record's part, when the record sits in `nodes`. */
function partOfText(room: Room, text: Y.Text): string | null {
  const rec = text.parent;
  if (!(rec instanceof Y.Map) || holderItem(rec)?.parent !== (room.nodes as unknown)) return null;
  const part = rec.get(F.part);
  return typeof part === 'string' ? part : null;
}

/** The character at offset `j` of a text item, or '' for an embed. */
function charOf(item: Y.Item, j: number): string {
  const content: unknown[] = item.content.getContent();
  return content.length > 0 && typeof content[0] === 'string'
    ? String(content.join('')).charAt(j)
    : '';
}

interface Char {
  readonly text: Y.Text;
  readonly item: Y.Item;
  readonly j: number;
}
const idOf = (c: Char): Y.ID => Y.createID(c.item.id.client, c.item.id.clock + c.j);

/** A character put back: where it is, the root it stands for, and the character itself. */
interface Restored extends Char {
  readonly root: Y.ID;
  readonly char: string;
}

/**
 * After an undo of `deletions` (the step's deleted items): the characters the undo put back that
 * a peer's copy shows go again, and the others follow the copy of their nearest moved neighbour.
 * `shows(part, key)` tells whether the read shows a text record. Call inside a transaction.
 * Returns whether it wrote anything.
 */
export function repairRestored(
  room: Room,
  deletions: IdSet,
  shows: (part: string, key: string) => boolean
): boolean {
  const doc = room.doc;
  // Each character put back: the new item and offset, and the identity of the old character.
  const restored: Restored[] = [];
  const memo = new Map<string, Y.ID>();
  let over = false;
  forEachItem(doc, deletions, (old) => {
    if (over || !old.countable || !(old.parent instanceof Y.Text) || old.redone === null) return;
    const text = old.parent;
    const now = itemOf(doc, old.redone);
    if (!partOfText(room, text) || !now || now.deleted || now.parent !== text) return;
    const length = Math.min(old.length, now.length);
    for (let j = 0; j < length; j += 1) {
      if (restored.length >= MAX_RESTORED) {
        over = true;
        return;
      }
      const root = rootOf(doc, Y.createID(old.id.client, old.id.clock + j), memo);
      restored.push({ text, item: now, j, root, char: charOf(now, j) });
    }
  });
  return !over && repairChars(room, restored, shows, memo, () => true);
}

/**
 * Once and follow for `restored`: a character whose root a shown copy in another text holds
 * goes, and the others go right after the copy of the nearest moved character before them. A
 * move is not idempotent, so only the characters `mayMove` allows move; the moved characters
 * keep their restore marks. Call inside a transaction. Returns whether it wrote anything.
 */
function repairChars(
  room: Room,
  restored: readonly Restored[],
  shows: (part: string, key: string) => boolean,
  memo: Map<string, Y.ID>,
  mayMove: (r: Restored) => boolean
): boolean {
  const doc = room.doc;
  if (restored.length === 0) return false;
  const parts = new Set<string>();
  for (const r of restored) parts.add(partOfText(room, r.text)!);
  const key = (r: Char) => idKey(r.item.id.client, r.item.id.clock + r.j);
  const isRestored = new Set(restored.map(key));
  // Every shown copy's live characters, by the identity of the character each one copies. A
  // character a redo put back in a copy counts too (its restore mark leads to the copy).
  const shown = shownCopies(room, parts, shows, memo, isRestored);
  if (!shown || shown.size === 0) return false;
  // Where a kept character goes: after the copy of the nearest moved character to its left, or
  // nowhere when a character that shows in its own text comes first.
  const anchorOf = (r: Restored): Char | null => {
    let looked = 0;
    let item: Y.Item | null = r.item;
    let j = r.j - 1;
    while (item !== null && looked < MAX_WALK) {
      for (; j >= 0; j -= 1, looked += 1) {
        if (!item.countable) break;
        const id = Y.createID(item.id.client, item.id.clock + j);
        if (isRestored.has(idKey(id.client, id.clock))) continue;
        const root = rootOf(doc, id, memo);
        const copy = shown.get(idKey(root.client, root.clock));
        if (copy && copy.text !== r.text) return copy;
        if (!item.deleted) return null;
      }
      item = item.left;
      j = item ? item.length - 1 : -1;
    }
    return null;
  };
  // Decide every character before writing, so each decision reads the text as the undo left it.
  const drop: Restored[] = [];
  const moves: { to: Char; chars: Restored[] }[] = [];
  for (const r of restored) {
    const copy = shown.get(idKey(r.root.client, r.root.clock));
    if (copy && copy.text !== r.text) {
      drop.push(r);
      continue;
    }
    const to = mayMove(r) ? anchorOf(r) : null;
    if (!to) continue;
    const last = moves[moves.length - 1];
    if (last && last.to.item === to.item && last.to.j === to.j) last.chars.push(r);
    else moves.push({ to, chars: [r] });
  }
  if (drop.length === 0 && moves.length === 0) return false;
  // Write the moved text after its anchor, then delete what was put back, last first.
  const spans: number[] = [];
  for (const { to, chars } of moves) {
    const at = liveIndex(to.text, idOf(to));
    if (at < 0) continue;
    const value = chars.map((c) => c.char).join('');
    if (value.length === 0) continue;
    const clock = Y.getState(doc.store, doc.clientID);
    to.text.insert(at + 1, value);
    chars.forEach((c, k) => spans.push(clock + k, c.root.client, c.root.clock, 1));
    drop.push(...chars);
  }
  markRestored(doc, spans);
  const byText = new Map<Y.Text, Set<string>>();
  for (const r of drop) {
    const ids = byText.get(r.text);
    if (ids) ids.add(key(r));
    else byText.set(r.text, new Set([key(r)]));
  }
  for (const [text, ids] of byText) deleteChars(text, ids);
  return true;
}

/**
 * After an undo or a redo of a step that deleted `deletions`: mark each character it put back
 * with the character it stands for (`markRestored` in `text-identity.ts`), so a copy whose
 * characters came back keeps its identity. Only this replica knows what it put back, so only it
 * writes the mark. Call inside a transaction. Returns the root of each character put back, and
 * the texts that hold them.
 */
export function noteRestored(
  room: Room,
  deletions: IdSet
): { roots: Set<string>; texts: Set<Y.Text> } {
  const doc = room.doc;
  const spans: number[] = [];
  const roots = new Set<string>();
  const texts = new Set<Y.Text>();
  const copyMemo = new Map<string, Y.ID>();
  let count = 0;
  forEachItem(doc, deletions, (old) => {
    if (!old.countable || !(old.parent instanceof Y.Text) || old.redone === null) return;
    const text = old.parent;
    const redone = old.redone;
    if (redone.client !== doc.clientID || partOfText(room, text) === null) return;
    const now = itemOf(doc, redone);
    if (!now || now.parent !== text) return;
    texts.add(text);
    for (let j = 0; j < old.length; j += 1) {
      if ((count += 1) > MAX_RESTORED) return;
      // The first character it stands for, so undo after redo keeps the chain one step long.
      let target = Y.createID(old.id.client, old.id.clock + j);
      for (let d = 0; d < MAX_BACK; d += 1) {
        const back = restoredOf(doc, target);
        if (!back) break;
        target = back;
      }
      // The root a delete noted (`rootOf` through copies only).
      const root = rootOf(doc, target, copyMemo, false);
      roots.add(idKey(root.client, root.clock));
      const clock = redone.clock + j;
      const n = spans.length;
      const joins =
        n > 0 &&
        spans[n - 4]! + spans[n - 1]! === clock &&
        spans[n - 3] === target.client &&
        spans[n - 2]! + spans[n - 1]! === target.clock;
      if (joins) spans[n - 1]! += 1;
      else spans.push(clock, target.client, target.clock, 1);
    }
  });
  markRestored(doc, spans);
  return { roots, texts };
}

/** The texts that hold a live character of restore marks `keys` (`RESTORED_KEY`). */
export function textsOfMarks(room: Room, keys: Iterable<string | null>): Set<Y.Text> {
  const doc = room.doc;
  const marks = doc.getMap(RESTORED_KEY);
  const texts = new Set<Y.Text>();
  for (const key of keys) {
    const m = key === null ? null : /^(\d{1,16}):(\d{1,16})$/.exec(key);
    const value: unknown = key === null ? undefined : marks.get(key);
    if (!m || !Array.isArray(value)) continue;
    const length = Math.min(Number(value[2]) || 0, MAX_RESTORED);
    const client = Number(m[1]);
    // Any live character of the span: a repair may have taken its first ones.
    for (let clock = Number(m[2]), end = clock + length; clock < end;) {
      const item = itemOf(doc, Y.createID(client, clock));
      if (!item) break;
      if (!item.deleted && item.parent instanceof Y.Text && partOfText(room, item.parent)) {
        texts.add(item.parent);
        break;
      }
      clock = item.id.clock + item.length;
    }
  }
  return texts;
}

/** The texts of `parts` that hold a live character an undo or a redo put back. */
export function restoredTextsOf(room: Room, parts: ReadonlySet<string>): Set<Y.Text> {
  const texts = new Set<Y.Text>();
  // The keys in one order on every replica before the cap, so each looks at the same marks.
  const keys = [...room.doc.getMap(RESTORED_KEY).keys()].sort().slice(0, MAX_RESTORED);
  for (const key of keys) {
    for (const text of textsOfMarks(room, [key])) {
      const part = partOfText(room, text);
      if (part !== null && parts.has(part)) texts.add(text);
    }
  }
  return texts;
}

/** The key of a text's record in `nodes`, or null. */
function keyOfText(room: Room, text: Y.Text): string | null {
  const holder = holderItem(text.parent);
  return holder && holder.parent === (room.nodes as unknown) && typeof holder.parentSub === 'string'
    ? holder.parentSub
    : null;
}

/**
 * Whether twin `a` stays before twin `b`: the smaller client, then, of one client, the later
 * clock. A client puts one character back twice only when it moved it ("follow" writes a new
 * character, then deletes the old one); a replica can see both before the delete arrives, and
 * must keep the new one, or the delete takes both.
 */
const stays = (a: Y.ID, b: Y.ID): boolean =>
  a.client < b.client || (a.client === b.client && a.clock > b.clock);

/**
 * Characters put back twice. Two peers that each undid their own delete of one character put
 * it back as two new characters: in one text, or in two (one peer's undo of an Enter puts the
 * character back in the run the Enter cut, and another peer's undo of a delete puts it back in
 * the Enter's copy); and a character an undo put back while a peer's Enter or format copied the
 * same character elsewhere shows in both places. Of the characters put back that stand for one
 * root, the one with the smallest id stays: of all the texts of the parts of `texts` that the
 * read shows (`shows`), and of each text that it does not show alone. None stays where a shown
 * copy holds the root: the copy is where the peer moved the text. Every replica decides alike,
 * so the deletes agree. Call inside a transaction; returns whether it wrote anything.
 */
export function dropRestoredTwins(
  room: Room,
  texts: Iterable<Y.Text>,
  shows: (part: string, key: string) => boolean
): boolean {
  const doc = room.doc;
  const memo = new Map<string, Y.ID>();
  const given = [...texts];
  const parts = new Set<string>();
  for (const text of given) {
    const part = partOfText(room, text);
    if (part !== null) parts.add(part);
  }
  if (parts.size === 0) return false;
  // A twin can sit in another text than the marks that arrived, so the decision reads every
  // text of the parts that holds a character put back: every replica then reads the same set.
  const list = new Set(restoredTextsOf(room, parts));
  for (const text of given) list.add(text);
  // The character that stays for each root (of a hidden text: for each text and root).
  const keep = new Map<string, { id: Y.ID; text: Y.Text }>();
  const drop = new Map<Y.Text, Set<string>>();
  const dropOne = (text: Y.Text, id: Y.ID): void => {
    const ids = drop.get(text);
    if (ids) ids.add(idKey(id.client, id.clock));
    else drop.set(text, new Set([idKey(id.client, id.clock)]));
  };
  let looked = 0;
  for (const text of list) {
    const part = partOfText(room, text);
    const key = keyOfText(room, text);
    if (part === null || key === null) continue;
    const scope = shows(part, key) ? '' : `${key}\0`;
    for (let it = text._start; it !== null && looked < MAX_RESTORED; it = it.right) {
      if (it.deleted || !it.countable) continue;
      for (let j = 0; j < it.length; j += 1, looked += 1) {
        const id = Y.createID(it.id.client, it.id.clock + j);
        if (!restoredOf(doc, id)) continue;
        const root = rootOf(doc, id, memo);
        const r = scope + idKey(root.client, root.clock);
        const kept = keep.get(r);
        if (!kept) keep.set(r, { id, text });
        else if (stays(id, kept.id)) {
          dropOne(kept.text, kept.id);
          keep.set(r, { id, text });
        } else dropOne(text, id);
      }
    }
  }
  for (const [text, ids] of drop) deleteChars(text, ids);
  // The characters that stay: once beside a shown copy, and this replica's own follow it.
  const survivors: Restored[] = [];
  for (const { id, text } of keep.values()) {
    const item = itemOf(doc, id);
    if (!item || item.deleted || item.parent !== text) continue;
    const j = id.clock - item.id.clock;
    survivors.push({ text, item, j, root: rootOf(doc, id, memo), char: charOf(item, j) });
  }
  const client = doc.clientID;
  const moved = repairChars(room, survivors, shows, memo, (r) => r.item.id.client === client);
  return moved || drop.size > 0;
}
