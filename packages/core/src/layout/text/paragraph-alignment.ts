// After DOCX Editor `packages/core/src/layout/paragraph-alignment.ts` at commit adf0e178:
// `paragraphAlignment`, the physical alignment of a paragraph. The source reads `w:jc`, whose
// `left` and `right` are the leading and trailing edges of a bidi paragraph. DrawingML
// `a:pPr/@algn` names physical edges instead: in the `rtl-text` reference PDF (slides 1 and
// 4), `algn="l"` in an `rtl="1"` paragraph ends at the left inset and `algn="r"` at the right
// inset. Only that rule differs.

import type { ResolvedParagraph } from '../../cascade/index.ts';

/** Horizontal alignment of a paragraph on the page. */
export type Alignment = 'left' | 'center' | 'right' | 'both';

/** The PHYSICAL alignment of a paragraph (`a:pPr/@algn`, ECMA-376 Part 1, 21.1.2.2.7). */
export function paragraphAlignment(algn: ResolvedParagraph['align']): Alignment {
  switch (algn) {
    case 'ctr':
      return 'center';
    case 'r':
      return 'right';
    // `justLow` justifies as `just` does, and `thaiDist` distributes as `dist` does.
    case 'just':
    case 'justLow':
    case 'dist':
    case 'thaiDist':
      return 'both';
    default:
      return 'left';
  }
}

/**
 * Where a line starts, measured from the paragraph's leading edge (the left edge, or the right
 * edge of a right-to-left paragraph), given the `room` the line leaves. A justified line, and
 * the last line of a justified paragraph, start at the leading edge (`rtl-text` slide 4: the
 * last line of a justified right-to-left paragraph ends at the right inset).
 */
export function leadingOffset(alignment: Alignment, rtl: boolean, room: number): number {
  if (alignment === 'center') return room / 2;
  return alignment === (rtl ? 'left' : 'right') ? room : 0;
}

const EMU_PER_PT = 12_700;

/**
 * The leading edge of a right-to-left paragraph, measured from the left edge of its column, for
 * a column `width` wide (EMU): the column width in whole points, rounded down. This is the line
 * model's whole-point grid (`placeBaseline`, `placeColumn` in `line-metrics.ts`) on the mirrored
 * axis; ECMA-376 gives no grid. In the `rtl-text` reference PDF, boxes with text areas of 273.6,
 * 417.6, 410.4, 266.4, and 871.2 pt end their right-to-left lines 273, 417, 410, 266, and 871 pt
 * from the left inset, also bullets and lines with a right indent. Line breaks, the room of an
 * aligned line, and justification still use the whole width, so every line of the paragraph
 * moves left by the cut.
 */
export function rtlLeadingEdge(width: number): number {
  return Math.max(0, Math.floor(width / EMU_PER_PT + 1e-6)) * EMU_PER_PT;
}
