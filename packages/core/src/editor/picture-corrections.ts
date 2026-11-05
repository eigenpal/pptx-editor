// The Corrections gallery of the `image` group: the Brightness/Contrast presets as a grid with a
// row per contrast and a column per brightness, each cell a small sample picture with that
// luminance effect. Both adapters draw the grid from these rules, so they cannot drift.

import { lumTransfer } from '../output/index.ts';
import type { ChromeControl } from './chrome-controls.ts';

type CorrectionOption = NonNullable<ChromeControl['options']>[number];

/** One cell of the Corrections grid. @internal */
export interface CorrectionGridCell {
  readonly option: CorrectionOption;
  /** Brightness from −1 to 1, as `setImageAdjustments` takes it. */
  readonly brightness: number;
  /** Contrast from −1 to 1, as `setImageAdjustments` takes it. */
  readonly contrast: number;
}

/** One row of the Corrections grid: one contrast, every brightness. @internal */
export interface CorrectionGridRow {
  readonly contrast: number;
  readonly cells: readonly CorrectionGridCell[];
}

/**
 * The Corrections grid from the `image.corrections` options (`brightness,contrast` in percent):
 * a row per contrast in the options' order. An option that is not a pair is left out. @internal
 */
export function correctionGridRows(
  options: readonly CorrectionOption[]
): readonly CorrectionGridRow[] {
  const rows = new Map<number, CorrectionGridCell[]>();
  for (const option of options) {
    const [b, c] = option.value.split(',').map(Number);
    if (!Number.isFinite(b) || !Number.isFinite(c)) continue;
    const row = rows.get(c!) ?? [];
    row.push({ option, brightness: b! / 100, contrast: c! / 100 });
    rows.set(c!, row);
  }
  return [...rows].map(([contrast, cells]) => ({ contrast: contrast / 100, cells }));
}

/**
 * A grid value as its axis label shows it, in percent with its sign: `+20`, `0`, `-40`, as the
 * cell names (`pictureFormat.correction`) give it. @internal
 */
export function correctionPercent(fraction: number): string {
  const v = Math.round(fraction * 100);
  return v > 0 ? `+${v}` : String(v);
}

/** The sample picture of a Corrections cell, in a 40 × 28 view box: sky, sun, hills, ground. */
const SAMPLE: readonly { readonly d: string; readonly rgb: readonly [number, number, number] }[] = [
  { d: 'M0 0H40V28H0Z', rgb: [142, 197, 232] },
  { d: 'M24 9a5 5 0 1 0 10 0a5 5 0 1 0 -10 0Z', rgb: [246, 196, 69] },
  { d: 'M0 19Q10 10 22 17T40 15V28H0Z', rgb: [94, 158, 74] },
  { d: 'M0 24Q18 18 40 23V28H0Z', rgb: [47, 107, 58] },
];

const hex = (v: number): string =>
  Math.round(Math.min(255, Math.max(0, v)))
    .toString(16)
    .padStart(2, '0');

/**
 * The paths of the sample picture with a luminance effect, as the slide paints a picture with
 * `a:lum` (the same transfer as the SVG paint). Each path has its `d` and its fill. @internal
 */
export function correctionSwatch(
  brightness: number,
  contrast: number
): readonly { readonly d: string; readonly fill: string }[] {
  const { slope, intercept } = lumTransfer(brightness, contrast);
  const channel = (v: number): string => hex((slope * (v / 255) + intercept) * 255);
  return SAMPLE.map(({ d, rgb }) => ({ d, fill: `#${rgb.map(channel).join('')}` }));
}
