// The vertical extent that an equation claims on its line.
//
// Adapted from DOCX Editor (eigenpal/docx-editor, packages/core/src/layout/equation-line-extent.ts
// at adf0e178c, Apache-2.0), `fitEquationLineExtent`: the line keeps the font box of the
// equation's base text (the band that the line model already gives it), and every other part
// claims for itself at its own position. A glyph piece claims its ink, plus the math leading
// above the ink, and never more than its own font box. A glyph piece with no outline claims the
// font box of the face that draws it. Text that draws nothing claims no height.
//
// Two rules differ from DOCX Editor, because the reference PDFs contradict them:
// - The math leading is 0, not 300/2048 em. With the leading, the lines of the `equations`,
//   `equation-lines`, and `preserved-content` fixtures move down by up to 3 pt (see
//   `MATH_LEADING`).
// - A structure claims the ink of its marks (rules, radical signs, delimiters, and braces), not
//   its box, also a fraction of a display equation. The reference PDF of `equation-lines`, slide
//   1, stacks two display fractions closer than their boxes: the ink of the second starts 1.25
//   pt above the end of the ink of the first.

import { mathTextInk, isInvisibleMathText } from './math-ink.ts';
import type { MathGlyphs, MathMark } from './math-layout.ts';
import type { TextMeasurer } from './measurer.ts';

const EMU_PER_PT = 12_700;

/**
 * The space kept above the ink of a part, in em of the leading size (`MathLeading` in DOCX
 * Editor: 300/2048 em). The reference PDFs leave none. With 300/2048 em, the second line of
 * the stacked superscripts of `equation-lines`, slide 1, is 2.95 pt below the reference, and
 * the equation line of `preserved-content`, slide 1, 3 pt; with 0, they are 0.05 pt and 0.09
 * pt from it.
 */
export const MATH_LEADING = 0;

/** The extent, in EMU above and below the math baseline, that an equation claims on its line. */
export interface MathLineExtent {
  readonly ascent: number;
  readonly descent: number;
}

/**
 * The extent that the glyph pieces and marks of an equation claim, in EMU from its baseline.
 * `leadingSizePt` is the size of the equation's base text: a part takes the smaller of it and
 * its own size for the leading, and a script of a script (level 2) takes it.
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
    let above = g.y + g.face.metrics.ascent * em;
    let below = g.face.metrics.descent * em - g.y;
    const ink = mathTextInk(measurer, g.face, g.text);
    if (ink.fromFontBox) {
      // A glyph with no outline: the box of the face that draws it, which can be a fallback
      // face taller than the piece's own face.
      above = Math.max(above, g.y + ink.hi * em);
      below = Math.max(below, -ink.lo * em - g.y);
    } else {
      const sizePt = g.level >= 2 ? leadingSizePt : Math.min(leadingSizePt, g.sizePt);
      above = Math.min(above, g.y + ink.hi * em + MATH_LEADING * sizePt * EMU_PER_PT);
      below = Math.min(below, -ink.lo * em - g.y);
    }
    ascent = Math.max(ascent, above);
    descent = Math.max(descent, below);
  }
  for (const m of marks)
    for (const c of m.commands)
      if (c.op !== 'Z') {
        ascent = Math.max(ascent, c.y);
        descent = Math.max(descent, -c.y);
      }
  return { ascent, descent };
}
