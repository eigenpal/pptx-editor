// The outline presets that the shape and table chrome share: the dash list and the weights.
// A module of its own, so `chrome-table.ts` and `chrome-slot-commands.ts` share them without
// an import cycle.

import type { LineDash } from '../contracts/index.ts';

/** The preset dashes offered, solid first. @internal */
export const LINE_DASHES: readonly LineDash[] = [
  'solid',
  'sysDot',
  'sysDash',
  'dash',
  'dashDot',
  'lgDash',
  'lgDashDot',
];

/** Outline weights offered, in points. */
export const LINE_WIDTHS = [0.75, 1, 1.5, 2.25, 3, 4.5, 6] as const;
