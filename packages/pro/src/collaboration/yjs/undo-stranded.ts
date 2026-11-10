/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import * as Y from 'yjs';
import type { ReplicaElementSpec, ReplicaTextSpec } from '@pptx-editor.dev/core/store';
import { F, type Room } from './schema.ts';

type Spec = ReplicaElementSpec | ReplicaTextSpec;
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const MAX_KEYS = 1 << 18;
const isA = (s: Spec, name: string): s is ReplicaElementSpec =>
  !('text' in s) && s.ns === NS_A && s.local === name;

/** Place an undone paragraph whose parent disappeared in a surviving cell of its own table. */
export function rehomeStranded(
  room: Room,
  spec: ReplicaElementSpec,
  entries: readonly { key: string; parent: string; table: string | null; row?: string }[],
  read: (key: string, seen: Set<string>, depth: number) => ReplicaElementSpec | null,
  shownIn: Map<string, string>,
  placed: Set<string>,
  readContainer = read
): ReplicaElementSpec {
  if (entries.length === 0) return spec;
  const shown = new Map<string, Spec>();
  const depths = new Map<string, number>();
  let budget = MAX_KEYS;
  const collect = (s: Spec, depth = 0): void => {
    if (--budget < 0 || depth > 64) {
      budget = -1;
      return;
    }
    depths.set(s.key, depth);
    shown.set(s.key, s);
    if (!('text' in s)) for (const c of s.children ?? []) collect(c, depth + 1);
  };
  collect(spec);
  if (budget < 0 || entries.length > MAX_KEYS) return spec;
  const pending = entries.filter((e) => !shown.has(e.key) && !shown.has(e.parent));
  if (pending.length === 0) return spec;
  const parents = new Map<string, string[]>();
  const root = room.nodes.get(spec.key);
  const part = root instanceof Y.Map ? root.get(F.part) : undefined;
  let looked = 0;
  for (const [key, rec] of room.nodes) {
    if ((looked += 1) > MAX_KEYS) return spec;
    const children =
      rec instanceof Y.Map && rec.get(F.part) === part ? rec.get(F.children) : undefined;
    if (!(children instanceof Y.Array)) continue;
    for (const child of children.toArray()) {
      if ((looked += 1) > MAX_KEYS) return spec;
      if (typeof child !== 'string') continue;
      const list = parents.get(child) ?? [];
      list.push(key);
      parents.set(child, list);
    }
  }
  const tableOf = (key: string): ReplicaElementSpec | undefined => {
    let level = [key];
    const seen = new Set<string>();
    for (let depth = 0; depth < 64 && level.length > 0; depth += 1) {
      const next: string[] = [];
      for (const k of level.sort()) {
        if (seen.has(k)) continue;
        seen.add(k);
        if (seen.size > MAX_KEYS) return undefined;
        const node = shown.get(k);
        if (node && isA(node, 'tbl')) return node;
        next.push(...(parents.get(k) ?? []));
      }
      level = next;
    }
    return undefined;
  };
  const additions = new Map<string, ReplicaElementSpec[]>();
  const givenBodies = new Map<string, ReplicaElementSpec>();
  const seen = new Set(shown.keys());
  for (const e of pending.sort((a, b) => (a.key < b.key ? -1 : a.key > b.key ? 1 : 0))) {
    const recorded = e.table ? shown.get(e.table) : undefined;
    const table = recorded && isA(recorded, 'tbl') ? recorded : tableOf(e.parent);
    if (!table) continue;
    const preferredRow = e.row ? shown.get(e.row) : undefined;
    const rows = preferredRow && isA(preferredRow, 'tr') ? [preferredRow] : (table.children ?? []);
    const cells = rows
      .filter((s) => isA(s, 'tr'))
      .flatMap((s) => s.children ?? [])
      .filter((s) => isA(s, 'tc'));
    let body = cells
      .flatMap((cell) => [
        ...(cell.children ?? []),
        ...(givenBodies.has(cell.key) ? [givenBodies.get(cell.key)!] : []),
      ])
      .find((s) => isA(s, 'txBody'));
    let recipient: ReplicaElementSpec | undefined;
    if (!body) {
      recipient = cells[0];
      if (!recipient || seen.has(e.parent)) continue;
      const original = readContainer(
        e.parent,
        new Set(seen),
        (depths.get(recipient.key) ?? 64) + 1
      );
      if (!original || !isA(original, 'txBody')) continue;
      body = {
        ...original,
        children: (original.children ?? []).filter((child) => !isA(child, 'p')),
      };
      depths.set(body.key, (depths.get(recipient.key) ?? 64) + 1);
    }
    const record = room.nodes.get(e.key);
    const shell =
      record instanceof Y.Map
        ? (record.get(F.shell) as { ns?: unknown; local?: unknown } | undefined)
        : undefined;
    if (shell?.ns !== NS_A || shell.local !== 'p') continue;
    const paragraph = read(e.key, seen, (depths.get(body.key) ?? 64) + 1);
    if (!paragraph || !isA(paragraph, 'p')) continue;
    const hasText = (s: Spec): boolean =>
      'text' in s ? s.text.length > 0 : (s.children ?? []).some(hasText);
    if (!hasText(paragraph)) continue;
    if (recipient) {
      givenBodies.set(recipient.key, body as ReplicaElementSpec);
      placed.add(body.key);
      shownIn.set(body.key, recipient.key);
      seen.add(body.key);
    }
    const list = additions.get(body.key) ?? [];
    list.push(paragraph);
    additions.set(body.key, list);
    placed.add(e.key);
    shownIn.set(e.key, body.key);
  }
  if (additions.size === 0) return spec;
  const insert = (s: ReplicaElementSpec): ReplicaElementSpec => ({
    ...s,
    children: [
      ...(givenBodies.has(s.key) ? [insert(givenBodies.get(s.key)!)] : []),
      ...(s.children ?? []).map((c) => ('text' in c ? c : insert(c))),
      ...(additions.get(s.key) ?? []),
    ],
  });
  return insert(spec);
}
