// The line model, in ONE module so calibration against the reference PDF output
// changes one place.
//
// Calibrated on 31 October 2026 against glyph origins in the reference PDFs: the
// fixtures (`line-model` holds the edge cases) and the 100-deck smoke set of eigenpal/pptx-eval
// (99 decks with references). The rules:
//
// - Percent spacing (`a:spcPct`) advances `pct × 1.2 ×` the largest font size on the line, for
//   the 11 Latin faces measured. Exact spacing (`a:spcPts`) advances the value rounded to a whole
//   point.
// - `lnSpcReduction` (`a:normAutofit`) is subtracted from the percentage: 110 % with a 10 %
//   reduction is 100 %. `fontScale` gives each run a whole-point size, rounded half up.
// - At the single line (1.2 × size), the baseline sits at the face's ascent share of the line.
//   Above it, for percent and exact spacing alike, the extra space goes 75 % above the
//   baseline, and the baseline sits at most at 75 % of the line box. Below it, the face's
//   descent stays below the baseline (a smaller line loses space above the text), but never
//   more than 25 % of the line box; a face with an ascent share under 75 % loses 75 % of the
//   space above the baseline, as a taller line gains it. A face with an ascent share of 75 % or more places every
//   line at 75 % of the box above the single line. The
//   face's ascent and descent are the `OS/2` values that fsSelection bit 7 selects
//   (`os2LineExtents` in `face-line-metrics.ts`), or the table values of a missing face.
// - Each baseline rounds to a whole point from the top of the text (its column top with
//   columns), before vertical anchoring (`placeBaseline`). ECMA-376 does not say this; the
//   reference PDFs show it on every fixture (`line-model`, `spacing`, `paragraph-format`,
//   `autofit`, `vertical-columns`): every laid-out line is then within 0.12 pt of the PDF glyph
//   origin, the half step of the 1/300 in grid that the PDF stores origins on. Table cells
//   round from their own text top too: 39 of 43 cell lines of `tables/1`, `tables/2`, and
//   `table-editing` are within 0.12 pt (0.35 pt at most without the rounding). The 4 others
//   are the first column of `tables/2` (the table style's major face, bold italic), which the
//   PDF puts one grid step above the other cells of its row; the rounding moves them from
//   0.04 pt to 0.36 pt off.
// - Diagram text (`dsp:txBody`, ECMA-376 Part 1, 21.4) has another single line: the face's
//   ascent plus descent (the values above), not 1.2 × the size. Percent spacing and the baseline
//   rules above then take that line (`LineBasis`). Measured on 6 November 2026 on the reference
//   PDF of `diagrams` and on the same target diagram laid out at three other frame sizes, at
//   15 to 34 pt with 90 % spacing: all 45 lines are within 0.12 pt, against up to 0.91 pt with
//   1.2 × the size. Plain shapes with the same text body keep 1.2 × the size.

import type { FaceMetrics } from './measurer.ts';

/** Line pitch multiplier for 100% spacing. */
export const SINGLE_LINE_FACTOR = 1.2;

/**
 * Share of the line box above the baseline when the box is taller than the single line, for
 * percent and exact spacing alike.
 */
const LOOSE_BASELINE_SHARE = 0.75;

const EMU_PER_PT = 12_700;

/** The effective percentage: `lnSpcReduction` is subtracted from the stored percentage. */
function effectivePct(pct: number, reduction: number): number {
  return Math.max(0, pct - reduction);
}

/**
 * A run size (points) under `a:normAutofit/@fontScale`: whole points, half up.
 * @internal
 */
export function autofitSizePt(sizePt: number, fontScale: number): number {
  if (fontScale >= 1) return sizePt;
  // The epsilon keeps products such as 20 × 0.475 = 9.4999… on the half-up side.
  return Math.max(1, Math.floor(sizePt * fontScale + 0.5 + 1e-9));
}

/**
 * Line box height (points) for a line whose largest font is `sizePt`. Exact spacing rounds to a
 * whole point; `reduction` (`lnSpcReduction`) is subtracted from a percentage.
 */
export function lineHeightPt(
  sizePt: number,
  spacing: { pct: number } | { pts: number },
  reduction: number
): number {
  return lineHeightInPt(SINGLE_LINE_FACTOR * sizePt, spacing, reduction);
}

/**
 * What a text body's single line is: 1.2 × the size (`size`), or the face's ascent plus
 * descent (`face`, diagram text).
 * @internal
 */
export type LineBasis = 'size' | 'face';

/**
 * The single line (points) of a line whose largest font is `sizePt` in a face with `metrics`.
 * @internal
 */
export function singleLinePt(sizePt: number, metrics: FaceMetrics, basis: LineBasis): number {
  return basis === 'face'
    ? sizePt * Math.max(0.01, metrics.ascent + metrics.descent)
    : SINGLE_LINE_FACTOR * sizePt;
}

/**
 * `lineHeightPt` for a single line of `singlePt` points (`singleLinePt`).
 * @internal
 */
export function lineHeightInPt(
  singlePt: number,
  spacing: { pct: number } | { pts: number },
  reduction: number
): number {
  if ('pts' in spacing) return Math.round(spacing.pts / 100);
  return singlePt * effectivePct(spacing.pct, reduction);
}

/** Distance from the line box top to the baseline (points), before rounding. */
export function baselineOffsetPt(sizePt: number, lineHeight: number, metrics: FaceMetrics): number;
/**
 * Distance from the line box top to the baseline (points), before rounding.
 * @deprecated The spacing is ignored: the line height decides. Call without it.
 */
export function baselineOffsetPt(
  sizePt: number,
  lineHeight: number,
  metrics: FaceMetrics,
  spacing: { pct: number } | { pts: number }
): number;
export function baselineOffsetPt(
  sizePt: number,
  lineHeight: number,
  metrics: FaceMetrics,
  _spacing?: { pct: number } | { pts: number }
): number {
  return baselineInLinePt(SINGLE_LINE_FACTOR * sizePt, lineHeight, metrics);
}

/**
 * `baselineOffsetPt` for a single line of `single` points (`singleLinePt`).
 * @internal
 */
export function baselineInLinePt(single: number, lineHeight: number, metrics: FaceMetrics): number {
  const loose = lineHeight * LOOSE_BASELINE_SHARE;
  const share = metrics.ascent / Math.max(0.01, metrics.ascent + metrics.descent);
  // Taller than the single line: the extra space goes 75 % above the baseline, but the
  // baseline never sits below 75 % of the box. A face whose ascent share is under 75 % keeps
  // its own share of the single line (Cordia New, win share 0.648, and Gabriola, typo share
  // 0.684, at 1.5 lines: 0 of 51 lines missed, against 51 of 51 with 75 % of the box).
  if (lineHeight > single + 1e-9)
    return Math.min(loose, LOOSE_BASELINE_SHARE * (lineHeight - single) + single * share);
  const fit = lineHeight - single * (1 - share);
  // The single line puts the baseline at the face's ascent share, also below 75 % (Cordia New
  // and Gabriola at single spacing: 0 of 178 lines missed, against 178 of 178 with the 75 %
  // floor). A tighter line keeps the face's descent below the baseline, but at most 25 % of the
  // box, as measured on Latin faces. A face under 75 % loses 75 % of the space from above the
  // baseline, as a taller line gains it (Noto Sans Arabic, typo share 0.651, at 90 %, 20 and
  // 36 pt: the 26 lines of `rtl-text` are within 0.12 pt, against up to 4 pt with the 25 %
  // limit). Both rules give 75 % of the box at a share of 75 %.
  if (lineHeight < single - 1e-9)
    return Math.max(
      fit,
      Math.min(loose, LOOSE_BASELINE_SHARE * (lineHeight - single) + single * share)
    );
  return fit;
}

/**
 * A baseline, in EMU from the top of the text, rounded to a whole point (12,700 EMU), as
 * measured on reference PDFs. Returns EMU.
 * @internal
 */
export function roundToWholePoint(emu: number): number {
  return Math.round(emu / EMU_PER_PT) * EMU_PER_PT;
}

/**
 * Whether layout rounds baselines to whole points, as measured on reference PDFs. On: every
 * line-checked fixture line is then within 0.12 pt of the PDF glyph origin; off, lines are up
 * to 0.61 pt off. Table cells: see the file header. Kept as a constant so `bun run fidelity:calibrate` can report the model it measures.
 * @internal
 */
export const WHOLE_POINT_BASELINES = true;

/** A baseline in EMU from the top of the text, placed as `WHOLE_POINT_BASELINES` says. @internal */
export function placeBaseline(emu: number): number {
  return WHOLE_POINT_BASELINES ? roundToWholePoint(emu) : emu;
}

/**
 * Paragraph spacing (`spcBef`/`spcAft`) in points. Points round to a whole point, as exact line
 * spacing does. Percent is of 1.2 × the first run size and does not round.
 */
export function paragraphSpacingPt(
  spacing: { pct: number } | { pts: number },
  sizePt: number
): number {
  if ('pts' in spacing) return Math.round(spacing.pts / 100);
  return spacing.pct * SINGLE_LINE_FACTOR * sizePt;
}

/** Small caps render lowercase letters at this fraction of the size. */
export const SMALL_CAPS_SCALE = 0.8;

/** Superscript/subscript (`baseline` ≠ 0) shrink runs to this fraction (measured). */
export const SCRIPT_SIZE_SCALE = 2 / 3;
