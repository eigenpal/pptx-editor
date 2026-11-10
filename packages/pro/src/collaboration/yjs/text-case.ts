/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Case choices change a character's presentation without replacing its Yjs identity.
// Each participant owns one choice per root character. Undo removes only that choice.
import * as Y from 'yjs';
import { idKey, itemOf, liveIds, recordOfItem, rootOf } from './text-identity.ts';

export const CASE_BASE_KEY = 'pptx-case-base-v1';
const PREFIX = 'cc.';
const MAX_FIELDS = 1 << 20;
interface Choice {
  order: number;
  author: number;
  value: number;
}
const cache = new WeakMap<Y.Map<unknown>, Map<string, Choice>>();
const watched = new WeakSet<Y.Map<unknown>>();

function choices(record: Y.Map<unknown>): Map<string, Choice> {
  const known = cache.get(record);
  if (known) return known;
  if (!watched.has(record)) {
    watched.add(record);
    record.observe(() => cache.delete(record));
  }
  const out = new Map<string, Choice>();
  if (record.size > MAX_FIELDS) return out;
  for (const [key, value] of record) {
    if (!key.startsWith(PREFIX)) continue;
    const match = /^cc\.(\d{1,16}:\d{1,16}):(\d{1,16})$/.exec(key);
    if (!match || !Array.isArray(value) || value.length !== 2) continue;
    const [order, text] = value;
    const author = Number(match[2]);
    if (
      !Number.isSafeInteger(order) ||
      order < 0 ||
      !Number.isSafeInteger(author) ||
      !Number.isInteger(text) ||
      text < 0 ||
      text > 0xffff
    )
      continue;
    const root = match[1]!;
    const old = out.get(root);
    if (!old || order > old.order || (order === old.order && author > old.author))
      out.set(root, { order, author, value: text });
  }
  cache.set(record, out);
  return out;
}

/** Materialize case choices through split, move, and restored character identities. */
export function readText(text: Y.Text): string {
  const raw = text.toString();
  const doc = text.doc;
  if (!doc) return raw;
  const bases = doc.getMap(CASE_BASE_KEY);
  if (bases.size === 0) return raw;
  const ids = liveIds(text, 0, text.length);
  const memo = new Map<string, Y.ID>();
  return ids
    .map((id, index) => {
      const root = rootOf(doc, id, memo);
      const key = idKey(root.client, root.clock);
      const base = bases.get(key);
      if (typeof base !== 'number' || !Number.isInteger(base) || base < 0 || base > 0xffff)
        return raw[index] ?? '';
      const item = itemOf(doc, root);
      const record = item && recordOfItem(item);
      return String.fromCharCode((record && choices(record).get(key)?.value) ?? base);
    })
    .join('');
}

/** Recognize only length-preserving case changes, including complete surrogate pairs. */
export function writeCase(text: Y.Text, old: string, value: string): boolean {
  const doc = text.doc;
  if (!doc || old.length !== value.length || old === value) return false;
  for (let i = 0; i < old.length;) {
    const before = String.fromCodePoint(old.codePointAt(i)!);
    const after = String.fromCodePoint(value.codePointAt(i)!);
    if (
      before.length !== after.length ||
      (before !== after && before.toUpperCase() !== after && before.toLowerCase() !== after)
    )
      return false;
    i += before.length;
  }
  const ids = liveIds(text, 0, text.length);
  const bases = doc.getMap(CASE_BASE_KEY);
  const growth = new Map<Y.Map<unknown>, Set<string>>();
  const writes: { record: Y.Map<unknown>; key: string; order: number; index: number }[] = [];
  for (let i = 0; i < old.length; i++) {
    if (old[i] === value[i]) continue;
    const root = rootOf(doc, ids[i]!);
    const item = itemOf(doc, root);
    const record = item && recordOfItem(item);
    if (!record || record.size >= MAX_FIELDS) return false;
    const key = idKey(root.client, root.clock);
    const order = (choices(record).get(key)?.order ?? 0) + 1;
    if (!Number.isSafeInteger(order)) return false;
    const field = `${PREFIX}${key}:${doc.clientID}`;
    if (!record.has(field)) {
      const keys = growth.get(record) ?? new Set<string>();
      keys.add(field);
      growth.set(record, keys);
      if (record.size + keys.size > MAX_FIELDS) return false;
    }
    writes.push({ record, key, order, index: i });
  }
  for (const { record, key, order, index } of writes) {
    // Outside the undo scope: copies need the original spelling after a case undo.
    if (!bases.has(key)) bases.set(key, old.charCodeAt(index));
    record.set(`${PREFIX}${key}:${doc.clientID}`, [order, value.charCodeAt(index)]);
    cache.delete(record);
  }
  return true;
}
