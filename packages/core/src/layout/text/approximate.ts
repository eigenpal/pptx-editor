// The measurers whose heights model no font: the fixed measurer (`createFixedMeasurer`), and a
// shaping measurer with no face (`createHarfBuzzMeasurer`), whose families all take stand-in
// metrics. A read that reports measured text (`readTextFit`) refuses with one. Not on the
// `layout` subpath: only the measurers mark, and only that read asks.

import type { TextMeasurer } from './measurer.ts';

const approximate = new WeakSet<TextMeasurer>();

/** Record that `measurer` models no font. */
export function markApproximateMeasurer(measurer: TextMeasurer): void {
  approximate.add(measurer);
}

/** Whether `measurer` models no font. */
export function isApproximateMeasurer(measurer: TextMeasurer): boolean {
  return approximate.has(measurer);
}
