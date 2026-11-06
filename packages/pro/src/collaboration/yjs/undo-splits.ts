/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// What an undo or a redo of a split must not put back (openspec change
// `collaboration-text-lineage`, "Undo of a split").
//
// A split moves the text after its cut out of the run's `Y.Text` into the pieces it mints. A
// `Y.UndoManager` undo of the split puts that text back as new characters, and a redo deletes
// them again and puts the pieces back. Yjs does not know who else took the same characters:
//
// - Two peers split one run at once (`text-lineage.ts`): both delete the same characters, and the
//   read shows only the winner's pieces. When the loser undoes, its undo puts the characters back
//   while the winner's pieces still show them: the text shows twice. When both undo, each puts
//   them back, because Yjs does not tell one replica what another restored. The winner's first
//   cut is the earliest, so its pieces hold every character a loser's split took: a loser's undo
//   puts none of the run's characters back (`withoutLoserRestores`). The winner's undo puts them
//   back, and the loser's pieces stay dropped, because a deleted winner still owns its group.
// - A split, its undo, then a peer's edit of the characters the undo put back (another split, or
//   a delete), then a redo of the first split: the redo would show its pieces beside the peer's.
//   When a peer took any of those characters, the redo leaves the run's characters and the
//   split's pieces alone (`withoutTakenRedo`); the peer's later edit wins.
//
// Both rules change the step before Yjs applies it; each reads only the step's own items and
// the records they name.

import * as Y from 'yjs';
import { F, partEntryOf, type Room } from './schema.ts';
import { lineageOf, LINEAGE, partLineage, winners } from './text-lineage.ts';
import { forEachItem, holderItem, withoutItems, type StepRanges } from './yjs-items.ts';

/** The `Y.Text` of each text node under a run's `a:t` children. */
function runTexts(room: Room, runKey: string, into: Set<Y.Text>): void {
  const childrenOf = (key: string): unknown[] => {
    const rec = room.nodes.get(key);
    const list = rec instanceof Y.Map ? rec.get(F.children) : undefined;
    return list instanceof Y.Array ? list.toArray() : [];
  };
  for (const t of childrenOf(runKey)) {
    if (typeof t !== 'string') continue;
    const shell = (room.nodes.get(t) as Y.Map<unknown> | undefined)?.get?.(F.shell) as
      { local?: unknown } | undefined;
    if (shell?.local !== 't') continue;
    for (const leaf of childrenOf(t)) {
      const rec = typeof leaf === 'string' ? room.nodes.get(leaf) : undefined;
      const text = rec instanceof Y.Map ? rec.get(F.text) : undefined;
      if (text instanceof Y.Text) into.add(text);
    }
  }
}

/** `type` when it is a node record (a map directly in `nodes`), else null. */
const recordOf = (room: Room, type: unknown): Y.Map<unknown> | null => {
  const holder = holderItem(type);
  return holder?.parent === (room.nodes as unknown) && type instanceof Y.Map ? type : null;
};

/**
 * Before an undo: `step`'s deletions without the characters of each run that this replica
 * split in the step, when a peer's concurrent split of the same run won (see the header).
 */
export function withoutLoserRestores(room: Room, step: StepRanges): StepRanges['deletions'] {
  const client = room.doc.clientID.toString(36);
  const winnerOf = winners(room);
  const origins = new Set<string>();
  // Per part, the groups whose winner's pieces all went with a subtree that a peer removed
  // without seeing them: there the loser's pieces hold the text (`partLineage`), so its undo
  // puts it back as usual.
  const unseen = new Map<string, ReadonlySet<string>>();
  const unseenIn = (part: unknown): ReadonlySet<string> => {
    if (typeof part !== 'string') return new Set();
    let groups = unseen.get(part);
    if (!groups) {
      const entry = partEntryOf(room.parts.get(part));
      groups =
        entry?.kind === 'xml' ? partLineage(room, part, entry.root).unseenWinners : new Set();
      unseen.set(part, groups);
    }
    return groups;
  };
  forEachItem(room.doc, step.insertions, (item) => {
    if (item.parent !== (room.nodes as unknown) || item.deleted) return;
    if (!(item.content instanceof Y.ContentType)) return;
    const rec = item.content.type;
    const p = lineageOf(rec);
    const winner = p ? winnerOf(p.group) : undefined;
    if (!p || winner === undefined || winner === client) return;
    if (!unseenIn((rec as Y.Map<unknown>).get(F.part)).has(p.group)) origins.add(p.from);
  });
  if (origins.size === 0) return step.deletions;
  const texts = new Set<Y.Text>();
  for (const origin of origins) runTexts(room, origin, texts);
  if (texts.size === 0) return step.deletions;
  return withoutItems(room.doc, step.deletions, (item) => texts.has(item.parent as Y.Text));
}

/**
 * Before a redo: `step` without the re-deletes of a run's characters and the pieces of the split
 * it redoes, when a peer took any of the characters the undo put back (see the header).
 */
export function withoutTakenRedo(room: Room, step: StepRanges): StepRanges {
  // The texts whose characters the undo put back, and whether a peer took any of them since.
  const restored = new Map<Y.Text, boolean>();
  forEachItem(room.doc, step.insertions, (item) => {
    const text = item.parent;
    if (!(text instanceof Y.Text) || !item.countable) return;
    if (!recordOf(room, text.parent)) return;
    restored.set(text, (restored.get(text) ?? false) || item.deleted);
  });
  if (![...restored.values()].some(Boolean)) return step;
  // The pieces the undo took out: their `sf` entry, deleted by the undo, names their run.
  const pieces = new Map<Y.Map<unknown>, string>();
  const takenTexts = new Set<Y.Text>();
  forEachItem(room.doc, step.deletions, (item) => {
    const rec = recordOf(room, item.parent);
    if (!rec || item.parentSub !== LINEAGE.splitFrom) return;
    if (!(item.content instanceof Y.ContentAny)) return;
    const origin = item.content.getContent()[0];
    if (typeof origin !== 'string') return;
    const texts = new Set<Y.Text>();
    runTexts(room, origin, texts);
    const taken = [...texts].filter((t) => restored.get(t) === true);
    if (taken.length === 0) return;
    for (const t of taken) takenTexts.add(t);
    const key = rec.get(F.key);
    if (typeof key === 'string') pieces.set(rec, key);
  });
  if (pieces.size === 0) return step;
  // The records under each piece, from the child entries the undo took: the redo leaves the
  // whole subtree alone, so no content of the redo lands in a node that no list shows.
  const childKeys = new Map<string, string[]>();
  forEachItem(room.doc, step.deletions, (item) => {
    const list = item.parent;
    const rec = list instanceof Y.Array ? recordOf(room, list.parent) : null;
    const parent = rec?.get(F.key);
    if (typeof parent !== 'string' || !(item.content instanceof Y.ContentAny)) return;
    let kids = childKeys.get(parent);
    if (!kids) childKeys.set(parent, (kids = []));
    for (const k of item.content.getContent()) if (typeof k === 'string') kids.push(k);
  });
  const keys = new Set(pieces.values());
  const subtree = new Set<unknown>();
  const stack = [...keys];
  while (stack.length > 0) {
    const key = stack.pop()!;
    const rec = room.nodes.get(key);
    if (subtree.has(rec) || !(rec instanceof Y.Map)) continue;
    subtree.add(rec);
    for (const k of childKeys.get(key) ?? []) stack.push(k);
  }
  /** The record an item belongs to: the record itself, or the record of its shared type. */
  const holderRecord = (item: Y.Item): unknown =>
    recordOf(room, item.parent) ?? recordOf(room, (item.parent as Y.AbstractType<unknown>)?.parent);
  const listsPiece = (item: Y.Item): boolean =>
    item.parent instanceof Y.Array &&
    recordOf(room, item.parent.parent) !== null &&
    item.content instanceof Y.ContentAny &&
    item.content.getContent().some((k) => typeof k === 'string' && keys.has(k));
  return {
    insertions: withoutItems(room.doc, step.insertions, (item) =>
      takenTexts.has(item.parent as Y.Text)
    ),
    deletions: withoutItems(
      room.doc,
      step.deletions,
      (item) => subtree.has(holderRecord(item)) || listsPiece(item)
    ),
  };
}
