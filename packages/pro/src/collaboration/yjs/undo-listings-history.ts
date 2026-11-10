/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Retain listing keys before a peer deletion lets Yjs collect their contents.
import * as Y from 'yjs';
import { F, type Room } from './schema.ts';
import { forEachItem, type StepRanges } from './yjs-items.ts';

interface Listed {
  readonly keys: Set<string>;
  readonly siblings: readonly string[];
}
type Step = StepRanges & { readonly meta: Map<string, unknown> };
const KEY = 'listed-children';
const MAX_KEYS = 1 << 16;

export function savedListings(step: StepRanges): ReadonlyMap<Y.Array<unknown>, Listed> {
  const value = (step as Step).meta?.get(KEY);
  return value instanceof Map ? value : new Map();
}

/** Record each step's live child listings while their keys are still available. */
export function rememberListings(room: Room, step: Step): void {
  const found = new Map(savedListings(step));
  let remaining = MAX_KEYS;
  for (const entry of found.values()) remaining -= entry.keys.size + entry.siblings.length;
  if (remaining <= 0) return;
  forEachItem(room.doc, step.insertions, (item) => {
    if (
      --remaining < 0 ||
      item.deleted ||
      !(item.parent instanceof Y.Array) ||
      !(item.content instanceof Y.ContentAny)
    )
      return;
    const list = item.parent;
    const record = list.parent;
    if (
      !(record instanceof Y.Map) ||
      record.parent !== room.nodes ||
      record.get(F.children) !== list
    )
      return;
    const keys = new Set(found.get(list)?.keys);
    for (const key of item.content.getContent()) {
      if (--remaining < 0) return;
      if (typeof key === 'string') keys.add(key);
    }
    const siblings = list
      .slice(0, MAX_KEYS)
      .filter((key): key is string => typeof key === 'string');
    remaining -= siblings.length;
    if (remaining >= 0) found.set(list, { keys, siblings });
  });
  step.meta.set(KEY, found);
}
