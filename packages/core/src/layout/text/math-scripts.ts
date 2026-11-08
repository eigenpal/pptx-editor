// The shifts of a subscript and a superscript beside a base.
//
// Adapted from DOCX Editor (eigenpal/docx-editor, packages/core/src/layout/equation-scripts.ts at
// adf0e178c, Apache-2.0), `attachScripts`: offsets in em of the construct size, from the math
// constants of the default math face (2048 font units to the em). The parts are ink boxes
// (`math-layout.ts`), so the clearances are between drawn glyphs. Where the reference PDFs
// contradict a constant or a rule, the comment of the constant gives the measurement.

/**
 * The least superscript baseline rise. DOCX Editor uses 750/2048 em. The reference PDFs give
 * 0.36 em: with 750/2048 em, the superscript of slide 1 of `preserved-content` and its line move
 * 1 pt down (0.99 pt from the reference, against 0.09 pt with 0.36 em).
 */
const SUPERSCRIPT_SHIFT_UP = 0.36;
/**
 * The least superscript baseline rise in a cramped style: under a radical, or in a subscript.
 * DOCX Editor has no cramped style. Measured on the reference PDF of `equations`.
 */
const SUPERSCRIPT_SHIFT_UP_CRAMPED = 0.293;
/**
 * How far a superscript baseline may sit below the top of the base's ink. DOCX Editor takes the
 * top of the base's lowercase text (976/2048 em) and 460/2048 em, so an ascender keeps the same
 * rise. The reference PDFs raise a superscript over a taller base: with the lowercase rule, the
 * lines of slide 2 of `equations` (`b` squared under a radical) come 1.15 pt from the reference,
 * against 0.17 pt with the ink top and 0.3 em.
 */
const SUPERSCRIPT_BASELINE_DROP_MAX = 0.3;
/** The least ink gap between a superscript and a subscript of one base. */
const SUB_SUPERSCRIPT_GAP_MIN = 300 / 2048;
/** How high the top of a subscript's ink may reach above the base baseline. */
const SUBSCRIPT_TOP_MAX = 760 / 2048;
/** How high a superscript bottom may rise to open the gap above a subscript. */
const SUPERSCRIPT_BOTTOM_MAX_WITH_SUBSCRIPT = 765 / 2048;
/** The least subscript baseline drop. */
const SUBSCRIPT_SHIFT_DOWN = 0.2;

/** A part of a script construct: its ink extents in EMU above and below its baseline. */
interface ScriptPart {
  readonly asc: number;
  readonly desc: number;
  /** Some glyph has no outline, so the box of its face stands in for its ink. */
  readonly fromFontBox?: true;
}

/**
 * The rise of the superscript baseline (`u`) and the drop of the subscript baseline (`v`), in
 * EMU, for the scripts of a construct whose size is `em` EMU, in a `cramped` style or not.
 */
export function scriptShifts(
  base: ScriptPart,
  sup: ScriptPart | undefined,
  sub: ScriptPart | undefined,
  em: number,
  cramped: boolean
): { u: number; v: number } {
  // The superscript baseline rises at least a fixed share of the construct size, and never
  // sits lower than a fixed drop below the top of the base's ink.
  let u = sup
    ? Math.max(
        (cramped ? SUPERSCRIPT_SHIFT_UP_CRAMPED : SUPERSCRIPT_SHIFT_UP) * em,
        base.asc - SUPERSCRIPT_BASELINE_DROP_MAX * em
      )
    : 0;
  // The subscript baseline drops a fifth of an em, and lower when the top of its ink would
  // rise past a fixed height, as a tall subscript does. A font box that stands in for missing
  // glyph ink reaches far above the glyphs, so it says nothing about the limit.
  const inkTop = sub && !sub.fromFontBox ? sub.asc : 0;
  let v = sub ? Math.max(SUBSCRIPT_SHIFT_DOWN * em, inkTop - SUBSCRIPT_TOP_MAX * em) : 0;
  if (sup && sub) {
    // Keep an ink gap between the scripts. The superscript rises first, up to a limit, and the
    // subscript drops for the rest.
    const missing = SUB_SUPERSCRIPT_GAP_MIN * em - (u - sup.desc - (sub.asc - v));
    if (missing > 0) {
      const raised = Math.min(
        missing,
        Math.max(0, SUPERSCRIPT_BOTTOM_MAX_WITH_SUBSCRIPT * em - (u - sup.desc))
      );
      u += raised;
      v += missing - raised;
    }
  }
  return { u, v };
}
