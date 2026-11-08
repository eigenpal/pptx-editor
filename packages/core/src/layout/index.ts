// The `layout` lane: geometry evaluation, text measurement and layout, and positioned slide
// records. DOM-free; the text measurer is injected. Emits geometry, never paints.

export * from './slide-layout.ts';
export * from './table-layout.ts';
// Internal: the renderers a capability package installs (`@pptx-editor.dev/pro/drawing`).
export * from './drawing-renderers.ts';
export * from './geometry/geometry.ts';
// Internal: the dashes and line ends that the SVG paint and `pptx-to-pdf` both draw.
export * from './line-strokes.ts';
// Internal: the strikethrough bars that the SVG paint and `pptx-to-pdf` both draw.
export * from './text/strikeout-bars.ts';
export type { PresetShapeDefinition } from '@eigenpal/ooxml-core';
export * from './text/measurer.ts';
export * from './text/harfbuzz-measurer.ts';
export { setHarfBuzzWasmUrl } from './harfbuzz-wasm-binary.ts';
export * from './text/text-layout.ts';
export type { LineBidi } from './text/line-visual-order.ts';
export * from './text/line-metrics.ts';
export {
  autonumberText,
  isExactScheme,
  isSymbolFont,
  symbolBulletText,
} from './text/autonumber.ts';
export { breakOpportunities, isIdeographic, type KinsokuLists } from './text/line-break.ts';
// Font configuration, copied from DOCX Editor's layout lane (each file header names its source).
// Named exports: `font-resource.ts` also names a `FontRequest` and a `FontSubstitution` that
// differ from the measurer's.
export {
  HARD_MAX_AGGREGATE_FONT_BYTES,
  HARD_MAX_FONT_BYTES,
  HARD_MAX_FONT_SOURCES,
  FontResolutionError,
  boundedStructuralFontValidator,
  sha256FontBytes,
  type FontByteValidator,
  type FontResolutionErrorCode,
  type FontValidationResult,
} from './font-resource.ts';
export * from './font-composition.ts';
export * from './font-resolver.ts';
export { deferFonts } from './defer-fonts.ts';
export type { WarpedRun } from './text/text-warp.ts';
export { runGlyphs, type RunGlyph } from './text/bidi-line.ts';
