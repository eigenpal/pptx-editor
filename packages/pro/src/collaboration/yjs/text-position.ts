/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import * as Y from 'yjs';
import { liveIds, rootOf } from './text-identity.ts';

/** Find a visible character through split, move, and restoration identities. */
export function locateCharacter(
  doc: Y.Doc,
  id: Y.ID,
  nodes: Iterable<readonly [string, string]>,
  textOf: (key: string) => Y.Text | undefined
): { node: string; offset: number; text: Y.Text } | null {
  const wanted = rootOf(doc, id);
  let budget = 1 << 20;
  for (const [key, node] of nodes) {
    if (--budget < 0) return null;
    const text = textOf(key);
    if (!text) continue;
    if ((budget -= text.length) < 0) return null;
    const ids = liveIds(text, 0, text.length);
    for (let offset = 0; offset < ids.length; offset++) {
      const root = rootOf(doc, ids[offset]!);
      if (root.client === wanted.client && root.clock === wanted.clock)
        return { node, offset, text };
    }
  }
  return null;
}
