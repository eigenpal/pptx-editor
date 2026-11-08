// The strikethrough of a run (`a:rPr/@strike`, ECMA-376 21.1.2.3.9): the bars that the SVG
// paint, the warped-text outline, and `pptx-to-pdf` all draw, so the three agree.

import type { FaceMetrics } from './measurer.ts';

/**
 * The strikethrough bars of a run, in the unit of the size passed to `strikeoutBars`.
 * @internal
 */
export interface StrikeoutBars {
  /** The top of each bar, from the run's baseline, y down: a bar above the baseline is negative. */
  readonly tops: readonly number[];
  /** The height of each bar. */
  readonly thickness: number;
}

/** The stroke center and thickness, in em, for a face without usable `OS/2` strikeout values. */
const FALLBACK_CENTER_EM = 0.3;
const FALLBACK_MIN_THICKNESS_EM = 1 / 20;

/**
 * The bars of a single (`sngStrike`) or double (`dblStrike`) strikethrough at `size`.
 *
 * The face's `OS/2` stroke decides (`FaceMetrics.strikeoutPosition` and
 * `strikeoutThickness`): the stroke is `yStrikeoutSize` thick, and its CENTER sits
 * `yStrikeoutPosition` above the baseline. Measured on the reference PDFs at 18 pt, whose
 * decorations snap to whole points: `font-effects` puts the single stroke 4.0 to 5.0 pt above
 * the baseline, and `text-runs` 3.9 to 4.9 pt. Carlito's stroke is 134 units at 512 over 2048
 * (1.18 pt at 4.5 pt): a stroke centered at 4.5 pt snaps to both. A stroke whose top is at
 * 4.5 pt snaps to 3.0 to 4.0 pt in `font-effects` and 2.9 to 4.9 pt in `text-runs`.
 *
 * A double strikethrough is two bars of the same thickness with one thickness between them,
 * centered on the same line: `font-effects` draws them 3.0 to 4.0 pt and 5.0 to 6.0 pt above
 * the baseline.
 *
 * A face without usable values (`sfntStrikeoutStrokeEm` refuses it) centers the stroke at
 * 0.3 em, with the underline thickness and at least 1/20 em.
 * @internal
 */
export function strikeoutBars(metrics: FaceMetrics, size: number, double: boolean): StrikeoutBars {
  const measured =
    metrics.strikeoutPosition !== undefined && metrics.strikeoutThickness !== undefined;
  const center = (measured ? metrics.strikeoutPosition! : FALLBACK_CENTER_EM) * size;
  const thickness = measured
    ? metrics.strikeoutThickness! * size
    : Math.max(metrics.underlineThickness * size, FALLBACK_MIN_THICKNESS_EM * size);
  const tops = double
    ? [-center - thickness * 1.5, -center + thickness * 0.5]
    : [-center - thickness / 2];
  return { tops, thickness };
}
