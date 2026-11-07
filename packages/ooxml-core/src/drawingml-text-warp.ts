// DrawingML preset text warps (ECMA-376 Part 1, 20.1.9.19 `prstTxWarp`, 20.1.10.76
// `ST_TextShapeType`), format-neutral. Each warp is a geometry definition: adjust values, guide
// formulas, and paths, generated from the spec's `presetTextWarpDefinitions.xml`. The geometry
// evaluator draws them as it draws a preset shape; the text layout of each editor maps glyphs
// between the paths.
//
// This is PPTX Editor's module, not a copy of a DOCX Editor file.

import {
  definitionAdjustments,
  evaluateGeometry,
  type GeometryPath,
  type PresetAdjustment,
} from './drawingml-geometry.ts';
import { PRESET_TEXT_WARPS } from './drawingml-preset-text-warp.generated.ts';

const own = (name: string): boolean =>
  Object.prototype.hasOwnProperty.call(PRESET_TEXT_WARPS, name);

/**
 * The warps that `ST_TextShapeType` names, in the order of the schema. `textNoShape` draws the
 * text flat; every other name has paths.
 */
export const TEXT_WARP_NAMES: readonly string[] = Object.freeze([
  'textNoShape',
  'textPlain',
  'textStop',
  'textTriangle',
  'textTriangleInverted',
  'textChevron',
  'textChevronInverted',
  'textRingInside',
  'textRingOutside',
  'textArchUp',
  'textArchDown',
  'textCircle',
  'textButton',
  'textArchUpPour',
  'textArchDownPour',
  'textCirclePour',
  'textButtonPour',
  'textCurveUp',
  'textCurveDown',
  'textCanUp',
  'textCanDown',
  'textWave1',
  'textWave2',
  'textDoubleWave1',
  'textWave4',
  'textInflate',
  'textDeflate',
  'textInflateBottom',
  'textDeflateBottom',
  'textInflateTop',
  'textDeflateTop',
  'textDeflateInflate',
  'textDeflateInflateDeflate',
  'textFadeRight',
  'textFadeLeft',
  'textFadeUp',
  'textFadeDown',
  'textSlantUp',
  'textSlantDown',
  'textCascadeUp',
  'textCascadeDown',
]);

/** Whether a name is an `ST_TextShapeType` value. */
export function isTextWarpName(name: string): boolean {
  return TEXT_WARP_NAMES.includes(name);
}

/**
 * The paths of a preset text warp at a size (EMU), with its adjust values, in shape-local EMU.
 * Undefined for `textNoShape` and for a name the table does not hold.
 */
export function evaluatePresetTextWarp(
  name: string,
  w: number,
  h: number,
  adjust: ReadonlyMap<string, number>
): readonly GeometryPath[] | undefined {
  if (!own(name)) return undefined;
  return evaluateGeometry(PRESET_TEXT_WARPS[name]!, w, h, adjust).paths;
}

/**
 * The adjust values of a preset text warp in `a:avLst` order, with their handle ranges at the
 * size `w` by `h` (EMU). Undefined for a warp that has no paths.
 */
export function presetTextWarpAdjustments(
  name: string,
  w: number,
  h: number,
  adjust: ReadonlyMap<string, number>
): PresetAdjustment[] | undefined {
  if (!own(name)) return undefined;
  return definitionAdjustments(PRESET_TEXT_WARPS[name]!, w, h, adjust);
}
