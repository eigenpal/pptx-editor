// Adapted from DOCX Editor `packages/core/src/layout/run-kerning.ts` at commit adf0e178.
// `unkernedTextRanges` is unchanged. `isRunKerningEnabled` and `kernsText` are not ported:
// slide text layout applies the `a:rPr/@kern` threshold itself (ECMA-376 21.1.2.3.9), and
// paint places every glyph at its measured position, so no renderer kerns a whole run.

/**
 * Full-width East Asian text that pair kerning never moves: kana, ideographs, CJK symbols and
 * punctuation, and full-width and half-width forms. Each keeps its own advance when a run
 * turns kerning on. Latin text in the same faces still kerns. Hangul syllables (U+AC00 to
 * U+D7A3) are not in the list and keep the font's pairs; common Korean faces carry no Hangul
 * kerning pairs.
 */
const UNKERNED_CLASS =
  '[\\u2e80-\\u9fff\\uf900-\\ufaff\\ufe30-\\ufe4f\\uff00-\\uffef\\u{20000}-\\u{3ffff}]';
const UNKERNED_EAST_ASIAN = new RegExp(UNKERNED_CLASS, 'u');
const UNKERNED_EAST_ASIAN_RUNS = new RegExp(`${UNKERNED_CLASS}+`, 'gu');

/**
 * The UTF-16 ranges of `text` that kerning leaves alone, as `[start, end)` pairs, or an
 * empty list when none does. A shaper turns its `kern` feature off over each range.
 */
export function unkernedTextRanges(text: string): readonly (readonly [number, number])[] {
  // Every unkerned character has a code unit at U+2E80 or above: most text leaves here.
  let wide = false;
  for (let index = 0; index < text.length && !wide; index++)
    wide = text.charCodeAt(index) >= 0x2e80;
  if (!wide || !UNKERNED_EAST_ASIAN.test(text)) return [];
  const ranges: [number, number][] = [];
  for (const match of text.matchAll(UNKERNED_EAST_ASIAN_RUNS))
    ranges.push([match.index, match.index + match[0].length]);
  return ranges;
}
