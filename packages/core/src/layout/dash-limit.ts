// Shared native dash expansion budget for shape outlines and table borders. Geometry and
// source properties remain intact; the SVG and PDF painters read the derived solid line.

import type { ResolvedLine } from '../cascade/index.ts';
import { dashArray, strokeWidth } from './line-strokes.ts';

/** The most dash segments one outline or border may ask a native painter to construct. */
const MAX_DASHES = 100_000;

/** A solid replacement above the dash budget; undefined when the source line can paint. */
export function boundedDashLine(
  line: ResolvedLine | undefined,
  length: number,
  starts = 1
): ResolvedLine | undefined {
  if (!line || line.fill.type === 'none') return undefined;
  const dashes = dashArray(line, strokeWidth(line));
  if (!dashes) return undefined;
  const period = dashes.reduce((sum, value) => sum + value, 0);
  if (Number.isFinite(length) && (length / period + starts) * (dashes.length / 2) <= MAX_DASHES)
    return undefined;
  return { ...line, dash: 'solid', dashes: undefined };
}
