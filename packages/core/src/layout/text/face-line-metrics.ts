// The ascent and descent that place a face's lines, and a table of them for common deck faces
// whose font is unavailable.
//
// Which `OS/2` values place a line (OpenType, "OS/2 table"), measured on 2,231 lines in 10 fonts
// exported to reference PDFs: when fsSelection bit 7 (USE_TYPO_METRICS) is set,
// sTypoAscender and sTypoDescender; otherwise usWinAscent and usWinDescent, never `hhea`. The
// rule predicts every line within a rounding step: Aptos (typo 1923 and 577; win misses 68 of
// 178 lines), Consolas (win; `hhea` misses 109), Yu Gothic (win 2017 and 619; `hhea` misses 64
// of 240), Cambria Math (typo 1593 and 455, with bit 7 set in an `OS/2` version 3 table; win
// misses 90), Century Gothic, Cordia New, and Gabriola. Yu Mincho misses 4 lines, all half-point
// ties at 18 pt (101.4998 pt predicted, 102 pt in the PDF). Most faces store the same values in
// every set, so the rule changes nothing for them.
//
// No open face stands in for the East Asian families here, so a missing one measures with the
// fallback face, and its lines sit where the fallback's ascent and descent put them. The
// packaged stand-ins for the Latin families here match their advances, not their ascent and
// descent (TeX Gyre Adventor for Century Gothic, Carlito for Aptos). The
// line model needs only two facts of the face it replaces: the ascent and descent that the rule
// above selects. Those are numbers,
// not outlines, so a table keeps the lines of a missing face in place without shipping it. The
// advances still come from the face that draws the text.
//
// This follows DOCX Editor's `packages/core/src/layout/symbol-face-metrics.ts`. The values are
// read with the rule above from the faces that the reference PDFs use (the regular face of each
// family): typo values for Aptos and Aptos Narrow (bit 7 set), win values for every other entry. Format
// neutral: no PresentationML names.

import type { FaceMetrics } from './measurer.ts';

interface FaceLineMetrics {
  /** unitsPerEm. */
  readonly upem: number;
  /** The ascent that `os2LineExtents` selects, font units. */
  readonly ascent: number;
  /** The descent that `os2LineExtents` selects, font units (positive). */
  readonly descent: number;
}

/** fsSelection bit 7, USE_TYPO_METRICS. */
const USE_TYPO_METRICS = 0x80;

/**
 * The ascent and descent (font units, both positive) that place a face's lines, from its `OS/2`
 * table: sTypoAscender and sTypoDescender when fsSelection bit 7 is set, else usWinAscent and
 * usWinDescent. The bit is read in every table version: Cambria Math sets it in version 3, and
 * its lines follow the typo values. Undefined when the table is missing or too short, or when
 * the selected values are both 0; the caller then takes the `hhea` extents.
 * @internal
 */
export function os2LineExtents(os2: Uint8Array | null | undefined): [number, number] | undefined {
  if (!os2 || os2.length < 78) return undefined;
  const view = new DataView(os2.buffer, os2.byteOffset, os2.byteLength);
  const typo = (view.getUint16(62) & USE_TYPO_METRICS) !== 0;
  const ascent = typo ? view.getInt16(68) : view.getUint16(74);
  const descent = typo ? -view.getInt16(70) : view.getUint16(76);
  if (ascent + descent <= 0) return undefined;
  return [Math.max(0, ascent), Math.max(0, descent)];
}

const FACE_LINE_METRICS: ReadonlyMap<string, FaceLineMetrics> = new Map([
  ['aptos', { upem: 2048, ascent: 1923, descent: 577 }],
  ['aptos narrow', { upem: 2048, ascent: 1923, descent: 577 }],
  ['cambria', { upem: 2048, ascent: 1946, descent: 455 }],
  ['century gothic', { upem: 2048, ascent: 1989, descent: 451 }],
  ['yu gothic', { upem: 2048, ascent: 2017, descent: 619 }],
  ['yu gothic medium', { upem: 2048, ascent: 2028, descent: 608 }],
  ['yu gothic light', { upem: 2048, ascent: 2026, descent: 610 }],
  ['yu gothic ui', { upem: 2048, ascent: 2210, descent: 514 }],
  ['yu mincho', { upem: 2048, ascent: 2038, descent: 598 }],
  ['meiryo', { upem: 2048, ascent: 2171, descent: 901 }],
  ['meiryo ui', { upem: 2048, ascent: 2171, descent: 430 }],
  ['ms gothic', { upem: 256, ascent: 220, descent: 36 }],
  ['ms pgothic', { upem: 256, ascent: 220, descent: 36 }],
  ['ms mincho', { upem: 256, ascent: 220, descent: 36 }],
  ['ms pmincho', { upem: 256, ascent: 220, descent: 36 }],
  ['simsun', { upem: 256, ascent: 220, descent: 36 }],
  ['dengxian', { upem: 2048, ascent: 1659, descent: 475 }],
  ['microsoft yahei', { upem: 2048, ascent: 2167, descent: 536 }],
  ['microsoft jhenghei', { upem: 2048, ascent: 2203, descent: 521 }],
  ['mingliu', { upem: 1024, ascent: 820, descent: 204 }],
  ['pmingliu', { upem: 1024, ascent: 820, descent: 204 }],
  ['malgun gothic', { upem: 2048, ascent: 2229, descent: 495 }],
]);

/**
 * `metrics` of the face that draws a run, with the ascent and descent of `family` when the
 * table knows it. For a run whose own face is missing.
 * @internal
 */
export function withFaceLineMetrics(family: string, metrics: FaceMetrics): FaceMetrics {
  const known = FACE_LINE_METRICS.get(family.trim().toLowerCase());
  if (!known) return metrics;
  return { ...metrics, ascent: known.ascent / known.upem, descent: known.descent / known.upem };
}
