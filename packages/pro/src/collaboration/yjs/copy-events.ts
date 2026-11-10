/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Copy metadata can arrive after its node when an update waits for dependencies.
import * as Y from 'yjs';
import { F, type Room } from './schema.ts';
import {
  COPIES_KEY,
  IDENTITY,
  copyOf,
  itemOf,
  recordOfItem,
  restoredOf,
  MAX_CHAIN,
} from './text-identity.ts';

/** Copies created or changed by a transaction, grouped by package part. */
export function changedCopies(
  room: Room,
  transaction: Y.Transaction,
  includeChangedText = true
): Map<string, string[]> {
  const out = new Map<string, string[]>();
  const keys = transaction.changed.get(
    room.nodes as unknown as Y.AbstractType<Y.YEvent<Y.Map<unknown>>>
  );
  for (const type of includeChangedText ? transaction.changed.keys() : []) {
    if (!(type instanceof Y.Text) || !(type.parent instanceof Y.Map)) continue;
    const rec = type.parent;
    const key = rec.get(F.key);
    const part = rec.get(F.part);
    if (typeof key !== 'string' || typeof part !== 'string' || !copyOf(rec)) continue;
    const list = out.get(part) ?? [];
    if (!list.includes(key)) list.push(key);
    out.set(part, list);
  }
  const index = room.doc.getMap(COPIES_KEY);
  const added = new Set(keys ?? []);
  for (const [type, fields] of transaction.changed) {
    if (!(type instanceof Y.Map) || type.parent !== room.nodes || !copyOf(type)) continue;
    if (!fields.has(IDENTITY.copyOf) && !fields.has(IDENTITY.newParagraph)) continue;
    const key = type.get(F.key);
    if (typeof key === 'string') added.add(key);
  }
  for (const key of transaction.changed.get(index as never) ?? [])
    if (typeof key === 'string') added.add(key);
  for (const k of added) {
    if (k === null) continue;
    const part = index.get(k);
    const rec = room.nodes.get(k);
    const copy = copyOf(rec);
    if (typeof part !== 'string' || !copy || copy.client === room.doc.clientID) continue;
    if ((rec as Y.Map<unknown>).get(F.part) !== part) continue;
    const list = out.get(part);
    if (list) list.push(k);
    else out.set(part, [k]);
  }
  return out;
}

/** A copy can name a character from another client's update that has not arrived yet. */
export function copySourcesReady(room: Room, partName: string): boolean {
  let remaining = 1 << 20;
  const checked = new Set<string>();
  const ready = (start: Y.ID): boolean => {
    let id = start;
    for (let depth = 0; depth < MAX_CHAIN; depth += 1) {
      if (--remaining < 0) return true;
      const key = `${id.client}:${id.clock}`;
      if (checked.has(key)) return true;
      const item = itemOf(room.doc, id);
      if (!item) return false;
      checked.add(key);
      const copy = copyOf(recordOfItem(item));
      const offset = copy ? id.clock - copy.clock : -1;
      if (copy && id.client === copy.client && offset >= 0 && offset < copy.length) {
        id = copy.source(offset);
        continue;
      }
      const restored = restoredOf(room.doc, id);
      if (!restored) return true;
      id = restored;
    }
    return true;
  };
  for (const [key, part] of room.doc.getMap(COPIES_KEY)) {
    if (--remaining < 0) return true;
    if (part !== partName) continue;
    const rec = room.nodes.get(key);
    const copy = copyOf(rec);
    if (!copy || !(rec instanceof Y.Map)) continue;
    const first = itemOf(room.doc, Y.createID(copy.client, copy.clock));
    if (!first || first.parent !== rec.get(F.text)) continue;
    for (let j = 0; j < copy.length; j += 1) if (!ready(copy.source(j))) return false;
  }
  return true;
}
