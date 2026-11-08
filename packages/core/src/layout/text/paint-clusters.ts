// Bound each pathological grapheme's native shaping work without changing source text or
// semantic caret boundaries. Both painters consume the resulting layout glyph runs.
import { segmentGraphemes } from '@eigenpal/ooxml-core';

/** Largest grapheme piece passed to a painter, in UTF-16 units. */
export const MAX_PAINT_CLUSTER_UNITS = 256;

/** Whether a grapheme needs bounded paint pieces. Ordinary text keeps its shaping context. */
export function hasLongPaintCluster(text: string): boolean {
  return (
    text.length > MAX_PAINT_CLUSTER_UNITS &&
    segmentGraphemes(text).some((g) => g.utf16To - g.utf16From > MAX_PAINT_CLUSTER_UNITS)
  );
}

/** Contiguous paint ranges; only oversized graphemes split, and surrogate pairs stay whole. */
export function paintClusterRanges(text: string): readonly { start: number; end: number }[] {
  if (!hasLongPaintCluster(text)) return [{ start: 0, end: text.length }];
  const ranges: { start: number; end: number }[] = [];
  let start = 0;
  for (const cluster of segmentGraphemes(text)) {
    if (cluster.utf16To - cluster.utf16From <= MAX_PAINT_CLUSTER_UNITS) continue;
    if (start < cluster.utf16From) ranges.push({ start, end: cluster.utf16From });
    start = cluster.utf16From;
    while (start < cluster.utf16To) {
      let end = Math.min(start + MAX_PAINT_CLUSTER_UNITS, cluster.utf16To);
      const code = text.charCodeAt(end - 1);
      if (end < cluster.utf16To && code >= 0xd800 && code <= 0xdbff) end -= 1;
      ranges.push({ start, end });
      start = end;
    }
  }
  if (start < text.length) ranges.push({ start, end: text.length });
  return ranges;
}

/**
 * `runs` with each very long grapheme split into bounded pieces, so no painter draws one
 * oversized cluster. Unit `i` of a right-to-left run draws in cell `n - 1 - i`, so its piece
 * `[a, b)` takes the cell edges `n - b` to `n - a`.
 */
export function paintPieces<
  R extends {
    readonly text: string;
    readonly x: readonly number[];
    readonly rtl?: true;
    readonly start: number;
  },
>(runs: readonly R[]): R[] {
  return runs.flatMap((run) => {
    const pieces = paintClusterRanges(run.text);
    if (pieces.length === 1) return [run];
    const n = run.text.length;
    return pieces.map((piece) => ({
      ...run,
      text: run.text.slice(piece.start, piece.end),
      x: run.rtl
        ? run.x.slice(n - piece.end, n - piece.start + 1)
        : run.x.slice(piece.start, piece.end + 1),
      start: run.start + piece.start,
    }));
  });
}
