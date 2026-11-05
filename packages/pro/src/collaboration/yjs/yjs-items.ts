/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Every read of Yjs structure that the public Yjs API does not offer, in one place, after DOCX
// Editor's `packages/pro/src/collaboration/document/yjs-items.ts` (commit 5993bd22). The undo
// rules read item ids, parents, and the ranges of a history step. Yjs declares these members in
// its own types, but outside its documented API, so an upgrade of Yjs is checked in this file.

import * as Y from 'yjs';

/** A set of Yjs id ranges: what a history step inserted or deleted. */
export type IdSet = ReturnType<typeof Y.createDeleteSet>;

/** The history step fields the undo rules read and replace. */
export interface StepRanges {
  insertions: IdSet;
  deletions: IdSet;
}

/** The item that holds a shared type in its parent, or null for a root type. */
export function holderItem(type: unknown): Y.Item | null {
  if (!(type instanceof Y.AbstractType)) return null;
  const item = type._item;
  return item instanceof Y.Item ? item : null;
}

/** Calls `visit` for each item of `doc` that `ranges` touches, once per item. */
export function forEachItem(doc: Y.Doc, ranges: IdSet, visit: (item: Y.Item) => void): void {
  for (const [client, list] of ranges.clients) {
    const structs = doc.store.clients.get(client);
    if (!structs || structs.length === 0) continue;
    for (const { clock, len } of list) {
      const last = structs[structs.length - 1]!;
      if (clock >= last.id.clock + last.length) continue;
      for (let i = Y.findIndexSS(structs, clock); i < structs.length; i += 1) {
        const struct = structs[i]!;
        if (struct.id.clock >= clock + len) break;
        if (struct instanceof Y.Item) visit(struct);
      }
    }
  }
}

/** `ranges` without the parts that cover an item `drop` selects. */
export function withoutItems(doc: Y.Doc, ranges: IdSet, drop: (item: Y.Item) => boolean): IdSet {
  const out = Y.createDeleteSet();
  for (const [client, list] of ranges.clients) {
    const structs = doc.store.clients.get(client);
    const kept: { clock: number; len: number }[] = [];
    for (const { clock, len } of list) {
      if (!structs || structs.length === 0) {
        kept.push({ clock, len });
        continue;
      }
      const end = clock + len;
      let at = clock;
      const last = structs[structs.length - 1]!;
      if (clock < last.id.clock + last.length)
        for (let i = Y.findIndexSS(structs, clock); i < structs.length; i += 1) {
          const struct = structs[i]!;
          if (struct.id.clock >= end) break;
          if (!(struct instanceof Y.Item) || !drop(struct)) continue;
          const from = Math.max(clock, struct.id.clock);
          if (from > at) kept.push({ clock: at, len: from - at });
          at = Math.min(end, struct.id.clock + struct.length);
        }
      if (end > at) kept.push({ clock: at, len: end - at });
    }
    if (kept.length > 0) out.clients.set(client, kept as never);
  }
  return out;
}

/** The live items of a type's sequence (a text's characters, an array's entries). */
// eslint-disable-next-line @typescript-eslint/no-explicit-any -- any shared type
export function* sequenceItems(type: Y.AbstractType<any>): Generator<Y.Item> {
  for (let item = type._start; item; item = item.right) if (!item.deleted) yield item;
}

/** The live items of a map, one per key. */
export function* mapItems(map: Y.Map<unknown>): Generator<Y.Item> {
  for (const item of map._map.values()) if (!item.deleted) yield item;
}

/** The live index of the character `id` in `text`, or -1 when it is deleted or elsewhere. */
export function liveIndex(text: Y.Text, id: Y.ID): number {
  let index = 0;
  for (let item = text._start; item !== null; item = item.right) {
    if (item.deleted || !item.countable) continue;
    if (
      item.id.client === id.client &&
      id.clock >= item.id.clock &&
      id.clock < item.id.clock + item.length
    )
      return index + (id.clock - item.id.clock);
    index += item.length;
  }
  return -1;
}

/**
 * Delete the live characters of `text` whose ids (`client:clock`) are in `ids`: one walk, then
 * one delete per range, from the end.
 */
export function deleteChars(text: Y.Text, ids: ReadonlySet<string>): void {
  const at: number[] = [];
  let index = 0;
  for (let item = text._start; item !== null && at.length < ids.size; item = item.right) {
    if (item.deleted || !item.countable) continue;
    for (let j = 0; j < item.length; j += 1)
      if (ids.has(`${item.id.client}:${item.id.clock + j}`)) at.push(index + j);
    index += item.length;
  }
  for (let end = at.length - 1; end >= 0;) {
    let start = end;
    while (start > 0 && at[start - 1] === at[start]! - 1) start -= 1;
    text.delete(at[start]!, at[end]! - at[start]! + 1);
    end = start - 1;
  }
}

/** Whether a map entry is the first value its key ever held (no earlier write it replaced). */
export const isFirstWrite = (item: Y.Item): boolean => item.left === null;
