/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Keep known text placements while Yjs waits for missing update dependencies.
import * as Y from 'yjs';
import type { NodeId, PackagePart } from '@pptx-editor.dev/core/store';
import { F, type Room } from './schema.ts';
import { LINEAGE, MOVES_KEY, isRemoved } from './text-lineage.ts';
import { isTextBody } from './text-identity.ts';
import { indexTree } from './mirror-support.ts';

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const MAX_KEYS = 1 << 16;

/** Previous text children and moved nodes whose destination listing is still pending. */
export function pendingChildren(
  room: Room,
  partName: string,
  previous: PackagePart | undefined,
  ids: ReadonlyMap<NodeId, string> | undefined
): Map<string, string[]> {
  const pending = new Map<string, string[]>();
  let remaining = MAX_KEYS;
  if (room.doc.store.pendingStructs && previous?.kind === 'xml') {
    for (const node of indexTree(previous.xml.root).values()) {
      if (--remaining < 0) return new Map();
      if (
        node.type !== 'element' ||
        !(isTextBody(node) || (node.ns === NS_A && node.local === 'p'))
      )
        continue;
      const key = ids?.get(node.id);
      if (key)
        pending.set(
          key,
          node.children.flatMap((c) => {
            const k = ids?.get(c.id);
            return k ? [k] : [];
          })
        );
    }
  }
  if (room.doc.store.pendingStructs) {
    for (const [key, part] of room.doc.getMap(MOVES_KEY)) {
      if (--remaining < 0) return new Map();
      if (part !== partName) continue;
      const rec = room.nodes.get(key);
      if (!(rec instanceof Y.Map) || rec.get(F.part) !== partName || isRemoved(rec)) continue;
      const to = rec.get(LINEAGE.movedTo);
      const at = rec.get(LINEAGE.movedIndex);
      if (!Array.isArray(to) || typeof to[0] !== 'string' || !Number.isSafeInteger(at) || at < 0)
        continue;
      const parent = room.nodes.get(to[0]);
      const children = parent instanceof Y.Map ? parent.get(F.children) : undefined;
      if (
        !(parent instanceof Y.Map) ||
        !(children instanceof Y.Array) ||
        parent.get(F.part) !== partName
      )
        continue;
      remaining -= children.length;
      if (remaining < 0) return new Map();
      const keys = pending.get(to[0]) ?? children.toArray();
      if (!keys.includes(key)) keys.splice(Math.min(at, keys.length), 0, key);
      pending.set(to[0], keys);
    }
  }
  return pending;
}
