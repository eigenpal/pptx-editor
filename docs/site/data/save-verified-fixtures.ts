// The save evidence of the feature matrix (`pptx-features.ts`, "Evidence rule"). It stays out of
// the matrix file, which the docs site syncs as site data.

/**
 * Fixtures whose saved copies open with no repair prompt. Both copies
 * from `bun run fidelity:roundtrip -- --saved <dir>` count: `.saved.pptx` (open, save) and
 * `.edited.pptx` (a scripted edit session, then save). Add a name only after you do that check (CONTRIBUTING.md, "Fidelity harness"). Last checked on 31 October 2026: all
 * 23 fixtures, 46 copies, no prompt (21 fixtures on copies written at commit 76a46c2,
 * `slide-size` at commit eb683ac, `slide-size-rounding` from its deck as committed in 787540c).
 * `table-editing`: 2 copies written at commit 0073112, no prompt on
 * 31 October 2026. `fill-fidelity` and `line-model`: 4 copies written at commit 1558bb1, no
 * prompt on 1 November 2026. That gives 26 fixtures.
 */
export const SAVE_VERIFIED_FIXTURES: readonly string[] = [
  'alignment-anchor',
  'autofit',
  'bullets-numbering',
  'colors',
  'custom-geometry',
  'effects',
  'fill-fidelity',
  'fills-lines',
  'font-substitution',
  'gradients',
  'groups-transforms',
  'line-model',
  'lines-arrows',
  'pictures',
  'placeholders-background',
  'preserve-unsupported',
  'preset-shapes',
  'slide-size',
  'slide-size-rounding',
  'slide-structure',
  'spacing',
  // A saved copy, plain and after a table edit session, opened with no repair prompt
  // (31 October 2026).
  'table-editing',
  'tables',
  'text-runs',
  'vertical-columns',
  'vertical-east-asian',
  // A saved copy opened with no repair prompt (31 October 2026, ui/pictures-engine).
  'picture-editing',
  // A saved copy opened with no repair prompt (31 October 2026, ui/shapes-engine).
  'arrange-flip-hidden',
  'shape-effects-presets',
  'shape-styles',
];
