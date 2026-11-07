// Line-break opportunities for slide text. Slide text never hyphenates. It breaks after
// spaces and tabs, after hyphens and dashes between letters, and next to Chinese and Japanese
// characters, which break per character. Korean breaks at spaces. A word wider than the line
// breaks between characters.
//
// Two paragraph properties change the rules (ECMA-376 Part 1, 21.1.2.2.7 `pPr`; measured on
// the reference PDFs of the `east-asian-line-breaking` fixture):
// - `eaLnBrk` applies the East Asian line breaking rules (kinsoku): no break before closing
//   punctuation, small kana, and the prolonged sound mark, and none after opening punctuation.
//   With `eaLnBrk="0"`, East Asian text still breaks per character, with no such rule.
// - `latinLnBrk="1"` lets a line break between any two characters, also inside a Latin word.
//
// The presentation's `p:kinsoku` (19.2.1.17) replaces the built-in lists for its language: its
// `invalStChars` may not start a line, and its `invalEndChars` may not end one (reference PDF of
// the `east-asian-typography` fixture: with custom lists, `」` may start a line).
//
// Opportunities are reported as UTF-16 offsets: a break at `i` means a line may end before
// unit `i`.
//
// No break, and no forced cut inside a word wider than the line, falls inside a grapheme cluster
// (UAX #29): an emoji sequence, a flag, a skin-tone modifier, a combining mark, or a Hangul
// syllable in jamo stay on one line. The clusters come from the `ooxml-core` segmenter, which
// DOCX Editor also uses for its East Asian and oversized-word breaks.

import { segmentGraphemes } from '@eigenpal/ooxml-core';

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

/** Whether a code point is Hangul: syllables, jamo, and compatibility jamo. */
export function isHangul(cp: number): boolean {
  return (
    (cp >= 0xac00 && cp <= 0xd7af) ||
    (cp >= 0x1100 && cp <= 0x11ff) ||
    (cp >= 0x3130 && cp <= 0x318f) ||
    (cp >= 0xa960 && cp <= 0xa97f) ||
    (cp >= 0xd7b0 && cp <= 0xd7ff)
  );
}

/**
 * Whether a line may break on either side of a character: East Asian characters, except
 * Hangul. Korean text breaks at spaces, as words (reference PDF of `east-asian-theme-fonts`);
 * a Korean word wider than the line breaks between characters.
 */
const breaksPerCharacter = (cp: number): boolean => isIdeographic(cp) && !isHangul(cp);

const isSpace = (c: string): boolean => c === ' ' || c === '\t' || c === '　';
const isLetter = (c: string): boolean => /[\p{L}\p{N}]/u.test(c);

/**
 * The first UTF-16 unit that can join the unit before it, or the unit after it, into one grapheme
 * cluster: combining marks start at U+0300, and every other cluster extender, prepended
 * character, joiner, Hangul jamo, emoji modifier, and surrogate lies above it. Between two units
 * below it a cluster always ends (UAX #29), except CR LF: a `&#13;` in `a:t` reaches line
 * breaking, so a carriage return takes the segmented path. Latin text needs no segmentation.
 */
const FIRST_CLUSTER_UNIT = 0x0300;

const CARRIAGE_RETURN = 0x0d;

/** Whether a cluster ends between two units without segmentation: both below U+0300, no CR. */
const plainSeam = (unit: number, before: number): boolean =>
  unit < FIRST_CLUSTER_UNIT && before < FIRST_CLUSTER_UNIT && before !== CARRIAGE_RETURN;

/**
 * Whether a line may not end before a unit of one text. Segments the text once, on the first
 * question that needs it.
 *
 * The clusters come from the runtime's `Intl.Segmenter`. Indic conjuncts (UAX #29 GB9c, such as
 * a Devanagari consonant, virama, and consonant) form one cluster only with Unicode 15.1 in the
 * runtime ICU; an older ICU ends a cluster after the virama.
 */
class ClusterTest {
  private starts: Uint8Array | null = null;

  constructor(private readonly text: string) {}

  /**
   * Whether a break before unit `i` would split a character: inside a grapheme cluster, before
   * a combining mark or variation selector (also after a space or a tab, which the mark then
   * belongs to), and on either side of a zero-width joiner (UAX #14 LB8a).
   */
  inside(i: number): boolean {
    const text = this.text;
    const unit = text.charCodeAt(i);
    const before = text.charCodeAt(i - 1);
    if (plainSeam(unit, before)) return false;
    if (unit === 0x200d || before === 0x200d) return true;
    if (unit >= FIRST_CLUSTER_UNIT && MARK.test(text[i])) return true;
    return this.insideCluster(i);
  }

  /**
   * Whether unit `i` continues a grapheme cluster, with no line breaking rule on top: the test
   * for a forced cut, which may fall after a joiner or before a mark where a cluster ends.
   */
  insideCluster(i: number): boolean {
    const text = this.text;
    if (plainSeam(text.charCodeAt(i), text.charCodeAt(i - 1))) return false;
    if (!this.starts) {
      const starts = new Uint8Array(text.length + 1);
      for (const segment of segmentGraphemes(text)) starts[segment.utf16From] = 1;
      starts[text.length] = 1;
      this.starts = starts;
    }
    return this.starts[i] === 0;
  }
}

/** The cluster test of each `breakOpportunities` result, for `forcedLineEnd`. */
const clusterTests = new WeakMap<Uint8Array, ClusterTest>();

const MARK = /^[\p{M}\uFE00-\uFE0F]$/u;

/**
 * The end of a line that starts at `start` when unit `at` passes the line end and no break
 * opportunity is left: the start of the grapheme cluster that holds `at`, so a word wider than
 * the line breaks between characters. When the first cluster alone is wider than the line, the
 * line ends after it. `breaks` is the `breakOpportunities` result for `text`.
 */
export function forcedLineEnd(text: string, breaks: Uint8Array, start: number, at: number): number {
  const test = clusterTests.get(breaks) ?? new ClusterTest(text);
  let end = Math.max(at, start + 1);
  while (end > start + 1 && end < text.length && test.insideCluster(end)) end -= 1;
  if (end === start + 1) while (end < text.length && test.insideCluster(end)) end += 1;
  return end;
}

/** Kinsoku lists: characters that may not start a line, and characters that may not end one. */
export interface KinsokuLists {
  /** Characters that may not start a line. */
  readonly notStart: string;
  /** Characters that may not end a line. */
  readonly notEnd: string;
}

const listCache = new WeakMap<KinsokuLists, { start: Set<string>; end: Set<string> }>();

function listSets(lists: KinsokuLists): { start: Set<string>; end: Set<string> } {
  let sets = listCache.get(lists);
  if (!sets) {
    sets = { start: new Set(lists.notStart), end: new Set(lists.notEnd) };
    listCache.set(lists, sets);
  }
  return sets;
}

/**
 * Break opportunities in `text`. `eaLineBreak` applies the East Asian line breaking rules
 * (`a:pPr/@eaLnBrk`, on by default); `latinLineBreak` allows a break between any two
 * characters (`a:pPr/@latinLnBrk`). `kinsoku` replaces the built-in lists of those rules.
 */
export function breakOpportunities(
  text: string,
  eaLineBreak = true,
  latinLineBreak = false,
  kinsoku?: KinsokuLists
): Uint8Array {
  const custom = kinsoku ? listSets(kinsoku) : undefined;
  const notStart = custom?.start ?? CLOSE;
  const notEnd = custom?.end ?? OPEN;
  const out = new Uint8Array(text.length + 1);
  out[text.length] = 1;
  const clusters = new ClusterTest(text);
  clusterTests.set(out, clusters);
  for (let i = 1; i < text.length; i += 1) {
    const prev = text[i - 1];
    const cur = text[i];
    // Never inside a grapheme cluster, before a combining mark, or at a joiner.
    if (clusters.inside(i)) continue;
    if (isSpace(prev) && !isSpace(cur)) {
      out[i] = 1;
      continue;
    }
    if (isSpace(cur)) continue;
    const held = eaLineBreak && (notStart.has(cur) || notEnd.has(prev));
    if (latinLineBreak) {
      if (!held) out[i] = 1;
      continue;
    }
    if (DASHES.has(prev) && i >= 2 && isLetter(text[i - 2]) && isLetter(cur)) {
      out[i] = 1;
      continue;
    }
    const prevCp =
      text.codePointAt(
        i -
          1 -
          (i >= 2 && text.charCodeAt(i - 1) >= 0xdc00 && text.charCodeAt(i - 1) <= 0xdfff ? 1 : 0)
      ) ?? 0;
    const curCp = text.codePointAt(i) ?? 0;
    if ((breaksPerCharacter(prevCp) || breaksPerCharacter(curCp)) && !held) out[i] = 1;
  }
  return out;
}

/**
 * East Asian commas and full stops: with `hangingPunct`, one of them that does not fit at the
 * line end hangs past it instead of taking the character before it to the next line
 * (reference PDF of `east-asian-theme-fonts`).
 */
export function isHangingPunctuation(code: number): boolean {
  return (
    code === 0x3001 ||
    code === 0x3002 ||
    code === 0xff0c ||
    code === 0xff0e ||
    code === 0xff61 ||
    code === 0xff64
  );
}

/** Whether a UTF-16 unit is whitespace that may hang past the right edge at a line end. */
export function isTrailingSpace(code: number): boolean {
  return code === 32 || code === 9 || code === 0x3000;
}
