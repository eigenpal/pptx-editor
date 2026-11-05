// Line-break opportunities for slide text. Slide text never hyphenates. It breaks after
// spaces and tabs, after hyphens and dashes between letters, and between East Asian
// ideographs (with the usual kinsoku: no break before closing punctuation or after opening
// punctuation). A word wider than the line breaks between characters.
//
// Opportunities are reported as UTF-16 offsets: a break at `i` means a line may end before
// unit `i`.

const CLOSE = new Set(
  '!),.:;?]}¢°’”‰′″℃、。〉》」』】〕〗〙〛！），．：；？］｝｡｣､ーぁぃぅぇぉっゃゅょゎゕゖァィゥェォッャュョヮヵヶ々〻ヽヾゝゞ'.split(
    ''
  )
);
const OPEN = new Set('([{£¥‘“〈《「『【〔〖〘〚（［｛｢＄'.split(''));
const DASHES = new Set(['-', '‐', '‒', '–', '—', '/']);

/** Whether a code point is an East Asian character that breaks per character. */
export function isIdeographic(cp: number): boolean {
  return (
    (cp >= 0x2e80 && cp <= 0x9fff) || // CJK radicals … unified ideographs, kana, bopomofo
    (cp >= 0xac00 && cp <= 0xd7af) || // Hangul syllables
    (cp >= 0xf900 && cp <= 0xfaff) || // CJK compatibility ideographs
    (cp >= 0xff00 && cp <= 0xffef) || // half/full width forms
    (cp >= 0x20000 && cp <= 0x3134f) // supplementary ideographs
  );
}

const isSpace = (c: string): boolean => c === ' ' || c === '\t' || c === '　';
const isLetter = (c: string): boolean => /[\p{L}\p{N}]/u.test(c);

/**
 * Break opportunities in `text`. `eaLineBreak` false disables per-character breaking of
 * ideographs (`a:pPr/@eaLnBrk="0"`).
 */
export function breakOpportunities(text: string, eaLineBreak = true): Uint8Array {
  const out = new Uint8Array(text.length + 1);
  out[text.length] = 1;
  for (let i = 1; i < text.length; i += 1) {
    const prev = text[i - 1];
    const cur = text[i];
    // Never inside a surrogate pair.
    const unit = text.charCodeAt(i);
    if (unit >= 0xdc00 && unit <= 0xdfff) continue;
    if (isSpace(prev) && !isSpace(cur)) {
      out[i] = 1;
      continue;
    }
    if (isSpace(cur)) continue;
    if (DASHES.has(prev) && i >= 2 && isLetter(text[i - 2]) && isLetter(cur)) {
      out[i] = 1;
      continue;
    }
    if (!eaLineBreak) continue;
    const prevCp =
      text.codePointAt(
        i -
          1 -
          (i >= 2 && text.charCodeAt(i - 1) >= 0xdc00 && text.charCodeAt(i - 1) <= 0xdfff ? 1 : 0)
      ) ?? 0;
    const curCp = text.codePointAt(i) ?? 0;
    if ((isIdeographic(prevCp) || isIdeographic(curCp)) && !CLOSE.has(cur) && !OPEN.has(prev))
      out[i] = 1;
  }
  return out;
}

/** Whether a UTF-16 unit is whitespace that may hang past the right edge at a line end. */
export function isTrailingSpace(code: number): boolean {
  return code === 32 || code === 9 || code === 0x3000;
}
