// Adapted from DOCX Editor `packages/core/src/layout/paragraph-alignment.ts` at commit adf0e178:
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
