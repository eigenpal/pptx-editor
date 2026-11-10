/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import * as Y from 'yjs';
import type { ReplicaElementSpec } from '@pptx-editor.dev/core/store';
import { F, type Room } from './schema.ts';
import { IDENTITY, REMOVAL_CLOCK_PREFIX, liveIds, rootOf } from './text-identity.ts';
import { rehomeStranded } from './undo-stranded.ts';

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const MAX_KEYS = 1 << 16;

/** Keep concurrent typing in cells whose row or column was removed by another participant. */
export function withRemovedTableText(
  room: Room,
  spec: ReplicaElementSpec,
  read: (key: string, seen: Set<string>, depth: number) => ReplicaElementSpec | null,
  shownIn: Map<string, string>,
  placed: Set<string>
): ReplicaElementSpec {
  let budget = MAX_KEYS;
  const clocks = new Map<string, Map<number, number>[]>();
  const entries: { key: string; parent: string; table: string; row?: string }[] = [];
  const record = (key: string) => room.nodes.get(key);
  const local = (key: string): string | undefined => {
    const rec = record(key);
    const shell = rec instanceof Y.Map ? rec.get(F.shell) : undefined;
    return shell?.ns === NS_A && typeof shell.local === 'string' ? shell.local : undefined;
  };
  const children = (key: string): string[] => {
    const rec = record(key);
    const list = rec instanceof Y.Map ? rec.get(F.children) : undefined;
    if (!(list instanceof Y.Array) || (budget -= list.length) < 0) return [];
    return list.toArray().filter((k): k is string => typeof k === 'string');
  };
  const removed = (key: string, kind: string): string[] => {
    const rec = record(key);
    if (!(rec instanceof Y.Map)) return [];
    const out: string[] = [];
    for (const [field, value] of rec) {
      if (--budget < 0) return [];
      if (!field.startsWith(IDENTITY.removedRun) || typeof value !== 'string') continue;
      const k = field.slice(IDENTITY.removedRun.length);
      if (local(k) === kind) out.push(k);
    }
    return out.sort();
  };
  const removalClocks = (key: string): Map<number, number>[] => {
    const rec = record(key);
    const out: Map<number, number>[] = [];
    if (!(rec instanceof Y.Map)) return out;
    for (const [field, value] of rec) {
      if (--budget < 0) return [];
      if (!field.startsWith(REMOVAL_CLOCK_PREFIX)) continue;
      if (!(value instanceof Uint8Array) || (budget -= value.length) < 0) {
        budget = -1;
        return [];
      }
      try {
        out.push(Y.decodeStateVector(value));
      } catch {
        budget = -1;
        return [];
      }
    }
    return out;
  };
  const takeCell = (cell: string, table: string, removedKey: string, row?: string) => {
    const seen = removalClocks(removedKey);
    if (!seen.length) return;
    for (const body of children(cell).filter((k) => local(k) === 'txBody'))
      for (const key of [...children(body), ...removed(body, 'p')])
        if (local(key) === 'p') {
          entries.push({ key, parent: body, table, ...(row ? { row } : {}) });
          clocks.set(key, [...(clocks.get(key) ?? []), ...seen]);
        }
  };
  const walk = (node: ReplicaElementSpec, depth = 0): void => {
    if (--budget < 0 || depth > 64) {
      budget = -1;
      return;
    }
    if (node.ns === NS_A && node.local === 'tbl') {
      for (const row of removed(node.key, 'tr'))
        for (const cell of children(row).filter((k) => local(k) === 'tc'))
          takeCell(cell, node.key, row);
      for (const row of children(node.key).filter((k) => local(k) === 'tr'))
        for (const cell of removed(row, 'tc')) takeCell(cell, node.key, cell, row);
    }
    for (const child of node.children ?? []) if (!('text' in child)) walk(child, depth + 1);
  };
  walk(spec);
  if (budget < 0) return spec;
  const unique = [...new Map(entries.map((entry) => [entry.key, entry])).values()];
  const readUnseen = (key: string, seen: Set<string>, depth: number) => {
    const paragraph = read(key, seen, depth);
    const removals = clocks.get(key) ?? [];
    if (!paragraph || !removals.length) return null;
    const hasUnseen = (node: ReplicaElementSpec): boolean => {
      for (const child of node.children ?? []) {
        if (!('text' in child)) {
          if (hasUnseen(child)) return true;
          continue;
        }
        const rec = record(child.key);
        const text = rec instanceof Y.Map ? rec.get(F.text) : undefined;
        if (!(text instanceof Y.Text) || text.length > MAX_KEYS) continue;
        for (const id of liveIds(text, 0, text.length)) {
          if ((budget -= 1 + removals.length) < 0) return false;
          const original = rootOf(room.doc, id);
          if (removals.every((clock) => original.clock >= (clock.get(original.client) ?? 0)))
            return true;
        }
      }
      return false;
    };
    return hasUnseen(paragraph) ? paragraph : null;
  };
  const addedPlaces = new Map<string, string>();
  const addedKeys = new Set<string>();
  const result = rehomeStranded(room, spec, unique, readUnseen, addedPlaces, addedKeys, read);
  if (budget < 0) return spec;
  for (const [key, parent] of addedPlaces) shownIn.set(key, parent);
  for (const key of addedKeys) placed.add(key);
  return result;
}
