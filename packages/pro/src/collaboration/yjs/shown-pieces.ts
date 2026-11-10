/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { clientOf } from './singleton-children.ts';
import type { Lineage } from './text-lineage.ts';

/** Index shown winning pieces once, rather than scanning all pieces for each loser. */
export function shownPieces(
  pieces: ReadonlyMap<string, Pick<Lineage, 'group' | 'at'>>,
  where: ReadonlyMap<string, { readonly order: number }>,
  winnerOf: (group: string) => string | undefined
): {
  last: ReadonlyMap<string, string>;
  preceding: (group: string, at: number) => string | undefined;
} {
  const groups = new Map<string, { key: string; at: number; order: number; rank: number }[]>();
  const last = new Map<string, string>();
  for (const [key, piece] of pieces) {
    if (clientOf(key) !== winnerOf(piece.group)) continue;
    const order = where.get(key)?.order;
    if (order === undefined) continue;
    const list = groups.get(piece.group) ?? [];
    list.push({ key, at: piece.at, order, rank: list.length });
    groups.set(piece.group, list);
    const previous = last.get(piece.group);
    if (previous === undefined || order > where.get(previous)!.order) last.set(piece.group, key);
  }
  // Equal cuts keep the last piece in document order, as the placement rule requires.
  for (const list of groups.values())
    list.sort((a, b) => a.at - b.at || a.order - b.order || b.rank - a.rank);
  return {
    last,
    preceding(group, at) {
      const list = groups.get(group);
      if (!list) return undefined;
      let lo = 0;
      let hi = list.length;
      while (lo < hi) {
        const mid = lo + Math.floor((hi - lo) / 2);
        if (list[mid]!.at <= at) lo = mid + 1;
        else hi = mid;
      }
      return list[lo - 1]?.key;
    },
  };
}
