// The vertical extent that an equation claims on its line.
//
// After DOCX Editor (eigenpal/docx-editor, packages/core/src/layout/equation-line-extent.ts
// at adf0e178c, Apache-2.0), `fitEquationLineExtent`: the line keeps the band of the equation's
// base text, and every other part claims for itself at its own position. A glyph piece claims
// its ink, plus the math leading above the ink. A glyph piece with no outline claims the font
// box of the face that draws it. Text that draws nothing claims no height. The line then takes
// the band and the claims together as its single line, and line spacing scales that line
// (`text-layout.ts`).
//
// These rules differ from DOCX Editor. The reference PDFs contradict the first four:
// - The math leading is 0.16 em of the base size for math text, and 0.14 em for normal text
//   (`m:nor`), not 300/2048 em (see `MATH_LEADING`).
// - A glyph piece claims its ink and the leading also above its own font box: a superscript
//   digit claims more than its face's ascent.
// - The band has an ascent share of 0.81 whatever face draws the equation (`MATH_LINE_BAND`).
// - A structure claims the ink of its marks (rules, radical signs, delimiters, and braces), not
//   its box, also a fraction of a display equation. The rule of a radical also claims the math
//   leading above it (`MathMark.claimAbove`, set in `math-layout.ts`).
// These are not measured, because no fixture holds them (the reference draws n-ary operators
// with a math font that is not open), so they keep the ink rule:
// - The operator of an inline n-ary structure claims its ink, not its whole box, and an inline
//   integral claims the ink of its parts, not the box of the whole structure.
// - The band is the line's own: the size of the zone's base runs. DOCX Editor takes the font
//   box of the first text atom, or of the structure size when a structure opens the equation. In
//   the fixtures, both are the same size.

import { mathTextInk, isInvisibleMathText } from './math-ink.ts';
import type { MathGlyphs, MathMark } from './math-layout.ts';
import type { TextMeasurer } from './measurer.ts';

const EMU_PER_PT = 12_700;

/**
 * The space that a glyph piece of math text claims above its ink, in em of the equation's base
 * size, at every script level. Measured on reference PDFs of the open STIX Two Math face at
 * 18 pt, in percent spacing from 80 % to 200 % and in exact spacing: stacked display fractions,
 * superscripts, and nested fractions put each line within the 1 pt baseline rounding of the
 * reference with 0.16 em, and up to 2 pt off with 300/2048 em (DOCX Editor's `MathLeading`) or
 * with no leading.
 */
export const MATH_LEADING = 0.16;

/**
 * The space that a glyph piece of normal text (`m:nor`) claims above its ink, in em of the
 * equation's base size. The superscript of `preserved-content`, slide 1, and of
 * `equation-lines`, slide 2, in Liberation Serif and STIX Two Math, put their lines 1 pt low
 * with `MATH_LEADING`.
 */
export const NORMAL_TEXT_LEADING = 0.14;

/**
 * The band of an equation line, in em: the ascent and descent that place the baseline in the
 * 1.2 × size line, whatever face draws the equation. A line of `α + β` in STIX Two Math (ascent
 * share 0.762) is within 0.12 pt of the reference at 80 % to 200 % spacing with this share of
 * 0.81, and up to 1 pt high with the face's own share.
 */
export const MATH_LINE_BAND = { ascent: 0.81, descent: 0.19 } as const;

/** The extent, in EMU above and below the math baseline, that an equation claims on its line. */
export interface MathLineExtent {
  readonly ascent: number;
  readonly descent: number;
}

/**
 * The extent that the glyph pieces and marks of an equation claim, in EMU from its baseline.
 * `leadingSizePt` is the size of the equation's base text, which sizes the leading of every
 * part.
 */
export function mathLineExtent(
  glyphs: readonly MathGlyphs[],
  marks: readonly MathMark[],
  leadingSizePt: number,
  measurer: TextMeasurer
): MathLineExtent {
  let ascent = 0;
  let descent = 0;
  for (const g of glyphs) {
    if (isInvisibleMathText(g.text)) continue;
    const em = g.sizePt * EMU_PER_PT;
    const ink = mathTextInk(measurer, g.face, g.text);
    let above: number;
    let below: number;
    if (ink.fromFontBox) {
      // A glyph with no outline: the box of its own face, or of the fallback face that draws
      // it, which can be taller.
      above = g.y + Math.max(g.face.metrics.ascent, ink.hi) * em;
      below = Math.max(g.face.metrics.descent, -ink.lo) * em - g.y;
    } else {
      const leading = g.normal ? NORMAL_TEXT_LEADING : MATH_LEADING;
      above = g.y + ink.hi * em + leading * leadingSizePt * EMU_PER_PT;
      below = -ink.lo * em - g.y;
    }
    ascent = Math.max(ascent, above);
    descent = Math.max(descent, below);
  }
  for (const m of marks)
    for (const c of m.commands)
      if (c.op !== 'Z') {
        ascent = Math.max(ascent, c.y + (m.claimAbove ?? 0));
        descent = Math.max(descent, -c.y);
      }
  return { ascent, descent };
}
