// Adapted from DOCX Editor `packages/core/src/layout/rtl-paragraph.ts` at commit adf0e178, with
// `rtl-letter.ts`, `arabic-script-language.ts`, and `line-end-whitespace.ts` of the same commit.
// The source reads run directions from `w:rtl`, resolves each run-direction group on its own,
// and works on styled spans. Slide runs have no direction, so a paragraph resolves as one
// UAX #9 paragraph, a run that reads against the paragraph resolves on its own like an
// isolate (`isolateRunLevels`), and this module works on the UTF-16 units of a paragraph.

import type { ResolvedParagraph } from '../../cascade/index.ts';
import { bidiAlgorithm } from './bidi.ts';

/**
 * Hebrew, Arabic, Syriac, Thaana, N'Ko, Samaritan, Mandaic and the Arabic extensions, with
 * their presentation forms, and the right-to-left scripts outside the Basic Multilingual Plane
 * (Adlam, Hanifi Rohingya, and others): the letters that start right-to-left text.
 */
export const RIGHT_TO_LEFT_LETTER =
  /[\u0590-\u08ff\ufb1d-\ufdff\ufe70-\ufeff\u{10800}-\u{10fff}\u{1e800}-\u{1efff}]/u;

/** Languages written in Arabic script, by primary subtag. */
const ARABIC_SCRIPT_LANGUAGES = new Set(['ar', 'fa', 'ur', 'ps', 'sd', 'ug', 'ckb', 'ks']);

/** Whether a run's language (`a:rPr/@lang`, BCP 47) is written in Arabic script. */
export function arabicScriptLanguage(lang: string | undefined): boolean {
  return lang !== undefined && ARABIC_SCRIPT_LANGUAGES.has(lang.toLowerCase().split('-')[0]!);
}

const HAS_BIDI_CONTROLS = /[\u061c\u200e\u200f\u202a-\u202e\u2066-\u2069]/u;
/**
 * Embedding, override, and isolate controls: authored structure that the per-run isolates
 * (`isolateRunLevels`) must not change. The marks (ALM, LRM, RLM) are strong characters only.
 */
const HAS_BIDI_STRUCTURE = /[\u202a-\u202e\u2066-\u2069]/u;

/**
 * A slash between ASCII digits. In a right-to-left run of an Arabic-script language it separates
 * the numbers instead of joining them into one, so `17/08/96` reads `96/08/17`. Resolution sees a
 * neutral in its place (UAX #9 class ON instead of CS); the source text does not change. A Hebrew
 * run keeps `852/20` one left-to-right number, as UAX #9 does. Unit tests cover this rule; no
 * reference PDF of slide text measures it.
 */
const RTL_RUN_DIGIT_SLASH = /(?<=[0-9])\/(?=[0-9])/gu;
const NEUTRAL_SLASH = '!';

/** The text of one run, as bidi resolution reads it. */
export interface BidiRunText {
  readonly text: string;
  /** The run's language, for the Arabic slash rule. */
  readonly lang?: string;
}

/** The bidi levels of a paragraph, one entry per UTF-16 unit of its source text. */
export interface ParagraphBidi {
  /** The paragraph's base direction (`a:pPr/@rtl`). */
  readonly rtl: boolean;
  /** UAX #9 embedding levels, resolved before line breaking. */
  readonly levels: Uint8Array;
}

/** The strong direction of a character: `true` right to left, `false` left to right. */
function strongRtl(ch: string): boolean | undefined {
  const type = bidiAlgorithm.getBidiCharTypeName(ch);
  return type === 'R' || type === 'AL' ? true : type === 'L' ? false : undefined;
}

/**
 * The direction a run reads in. Slide runs have no direction of their own. A run reads against
 * the paragraph only when every strong character in it has that other direction; any other run
 * reads in the paragraph's direction.
 */
export function runIsRtl(text: string, paragraphRtl: boolean): boolean {
  let against = false;
  for (const ch of text) {
    const strong = strongRtl(ch);
    if (strong === undefined) continue;
    if (strong === paragraphRtl) return paragraphRtl;
    against = true;
  }
  return against ? !paragraphRtl : paragraphRtl;
}

/**
 * Each slash between digits in a run of an Arabic-script language that reads right to left,
 * read as a neutral. The language of the run that holds the slash decides, so a date split over
 * two runs reads the same way as in one.
 */
function withArabicSlashes(text: string, runs: readonly BidiRunText[], rtl: boolean): string {
  const arabic = runs.map((run) => arabicScriptLanguage(run.lang) && runIsRtl(run.text, rtl));
  if (!arabic.includes(true)) return text;
  return text.replace(RTL_RUN_DIGIT_SLASH, (slash, at: number) => {
    let end = 0;
    for (let index = 0; index < runs.length; index++) {
      end += runs[index]!.text.length;
      if (at < end) return arabic[index] ? NEUTRAL_SLASH : slash;
    }
    return slash;
  });
}

/**
 * The partner of each paired bracket of `text` (UAX #9 BD16), or -1. Each paragraph (up to a
 * paragraph separator) pairs on its own, with a stack of 63 openings; a paragraph that
 * overflows the stack pairs none of its brackets, as BD16 stops there.
 */
function bracketPartners(text: string): Int32Array {
  const partner = new Int32Array(text.length).fill(-1);
  let stack: { at: number; closing: string }[] = [];
  let pairs: [number, number][] = [];
  let overflow = false;
  const endParagraph = (): void => {
    if (!overflow)
      for (const [open, close] of pairs) [partner[open], partner[close]] = [close, open];
    stack = [];
    pairs = [];
    overflow = false;
  };
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]!;
    if (bidiAlgorithm.getBidiCharTypeName(ch) === 'B') {
      endParagraph();
      continue;
    }
    if (overflow) continue;
    const closing = bidiAlgorithm.openingToClosingBracket(ch);
    if (closing) {
      if (stack.length === 63) overflow = true;
      else stack.push({ at: i, closing: bidiAlgorithm.getCanonicalBracket(closing) ?? closing });
      continue;
    }
    if (!bidiAlgorithm.closingToOpeningBracket(ch)) continue;
    const canonical = bidiAlgorithm.getCanonicalBracket(ch) ?? ch;
    for (let k = stack.length - 1; k >= 0; k--) {
      if (stack[k]!.closing !== canonical) continue;
      pairs.push([stack[k]!.at, i]);
      stack.length = k;
      break;
    }
  }
  endParagraph();
  return partner;
}

/**
 * Resolve each run that reads against the paragraph on its own, as an isolate does (UAX #9
 * X5a-X5c), one level above the paragraph: a right-to-left run of a left-to-right paragraph
 * from level 1, a left-to-right run of a right-to-left paragraph from level 2. Its letters,
 * digits, and punctuation keep their order inside it: in the `rtl-text` reference PDF, a
 * left-to-right run of a right-to-left paragraph keeps its colon and its quotation marks
 * (`An English phrase in a right-to-left paragraph: `, `"Hello, world!"`). Three kinds of
 * character keep their paragraph-wide level: the whitespace at the run's edges, so it still
 * separates the run from the text beside it; a bracket whose pair lies outside the run, so the
 * pair resolves as one (UAX #9 N0), with the whitespace between it and the run's text; and
 * everything when the paragraph has authored embedding, override, or isolate controls. Text outside the run sees the run's characters, so right-to-left words in two
 * runs of a left-to-right paragraph still read as one stretch.
 */
function isolateRunLevels(
  levels: Uint8Array,
  text: string,
  runs: readonly BidiRunText[],
  rtl: boolean
): void {
  let partners: Int32Array | undefined;
  let offset = 0;
  for (const run of runs) {
    const start = offset;
    offset += run.text.length;
    const runRtl = runIsRtl(run.text, rtl);
    if (runRtl === rtl) continue;
    const own = bidiAlgorithm.getEmbeddingLevels(
      runRtl && arabicScriptLanguage(run.lang)
        ? run.text.replace(RTL_RUN_DIGIT_SLASH, NEUTRAL_SLASH)
        : run.text,
      runRtl ? 'rtl' : 'ltr'
    ).levels;
    // The run's own base is 1 (right to left) or 0; the isolate starts one above the paragraph.
    const raise = rtl ? 2 : 0;
    partners ??= bracketPartners(text);
    const pairs = partners;
    // Whitespace and brackets paired outside the run, at either edge, stay with the paragraph.
    const edge = (i: number): boolean => {
      const pair = pairs[start + i]!;
      return /\s/u.test(run.text[i]!) || (pair >= 0 && (pair < start || pair >= offset));
    };
    let first = 0;
    while (first < run.text.length && edge(first)) first++;
    let last = run.text.length;
    while (last > first && edge(last - 1)) last--;
    for (let i = first; i < last; i++) {
      const pair = pairs[start + i]!;
      if (pair >= 0 && (pair < start || pair >= offset)) continue;
      levels[start + i] = own[i]! + raise;
    }
  }
}

/**
 * Resolve paragraph-wide levels before wrapping; source offsets remain logical. Undefined for a
 * left-to-right paragraph without right-to-left letters or bidi controls, which lays out
 * unchanged.
 */
export function paragraphBidi(
  runs: readonly BidiRunText[],
  rtl: boolean
): ParagraphBidi | undefined {
  const text = runs.map((run) => run.text).join('');
  const controls = HAS_BIDI_CONTROLS.test(text);
  if (!rtl && !controls && !RIGHT_TO_LEFT_LETTER.test(text)) return undefined;
  const levels = bidiAlgorithm.getEmbeddingLevels(
    withArabicSlashes(text, runs, rtl),
    rtl ? 'rtl' : 'ltr'
  ).levels;
  // Authored embedding, override, and isolate controls decide on their own.
  if (!HAS_BIDI_STRUCTURE.test(text)) isolateRunLevels(levels, text, runs, rtl);
  return { rtl, levels };
}

/** The bidi levels of a paragraph, kept for the `normAutofit` search, like its breaks. */
const bidiOf = new WeakMap<ResolvedParagraph, ParagraphBidi | null>();
export function paragraphBidiOf(p: ResolvedParagraph): ParagraphBidi | undefined {
  let kept = bidiOf.get(p);
  if (kept === undefined) {
    const runs = p.runs.map((r) => ({
      text: r.kind === 'break' ? '\n' : r.text,
      lang: r.props.lang,
    }));
    kept = paragraphBidi(runs, p.rtl) ?? null;
    bidiOf.set(p, kept);
  }
  return kept ?? undefined;
}

/** UAX #9 classes that rule L1 resets at a line end: whitespace, breaks, and tabs. */
function isLineEndClass(code: number): boolean {
  const type = bidiAlgorithm.getBidiCharTypeName(String.fromCharCode(code));
  return type === 'WS' || type === 'B' || type === 'S';
}

/** One line in visual order: line-relative unit indices, and each unit's level after L1. */
export interface BidiLineOrder {
  readonly order: readonly number[];
  readonly levels: readonly number[];
}

/**
 * UAX #9 L1/L2 for the units `[start, end)` of a paragraph (`reorderBidiSpans` and `bidiOrder`
 * in the source). `codes` holds the UTF-16 units of the line. A tab, the whitespace before it,
 * and the whitespace, tabs, and breaks that end the line take the paragraph level.
 */
export function lineVisualOrder(
  bidi: ParagraphBidi,
  codes: ArrayLike<number>,
  start: number,
  end: number
): BidiLineOrder {
  const levels = Array.from(bidi.levels.subarray(start, end));
  const count = end - start;
  const base = bidi.rtl ? 1 : 0;
  for (let index = 0; index < count; index++) {
    if (codes[index] !== 9) continue;
    levels[index] = base;
    for (let ws = index - 1; ws >= 0 && isLineEndClass(codes[ws]!) && codes[ws] !== 9; ws--)
      levels[ws] = base;
  }
  for (let index = count; index > 0 && isLineEndClass(codes[index - 1]!); index--)
    levels[index - 1] = base;
  return { order: visualOrderOfLevels(levels), levels };
}

/** UAX #9 L2 over `levels[start, end)`: the source indices in visual (left-to-right) order. */
export function visualOrderOfLevels(
  levels: readonly number[],
  start = 0,
  end = levels.length
): number[] {
  const order = Array.from({ length: end - start }, (_, index) => start + index);
  let maximum = 0;
  let lowestOdd = Infinity;
  for (let index = start; index < end; index++) {
    maximum = Math.max(maximum, levels[index]!);
    if (levels[index]! % 2) lowestOdd = Math.min(lowestOdd, levels[index]!);
  }
  for (let level = maximum; level >= lowestOdd; level--) {
    for (let i = 0; i < order.length;) {
      if (levels[order[i]!]! < level) {
        i++;
        continue;
      }
      let end = i + 1;
      while (end < order.length && levels[order[end]!]! >= level) end++;
      for (let left = i, right = end - 1; left < right; left++, right--)
        [order[left], order[right]] = [order[right]!, order[left]!];
      i = end;
    }
  }
  return order;
}

/**
 * Split pieces where the bidi direction changes, so each piece shapes in one direction: the
 * measurer guesses the direction from the piece's script, and digits in right-to-left text then
 * shape left to right. `offset` is the paragraph offset of the first piece.
 */
export function bidiPieces<T extends { readonly text: string }>(
  pieces: T[],
  levels: Uint8Array | undefined,
  offset: number
): T[] {
  if (!levels) return pieces;
  const out: T[] = [];
  let at = offset;
  for (const piece of pieces) {
    let from = 0;
    for (let i = 1; i <= piece.text.length; i += 1)
      if (i === piece.text.length || levels[at + i]! % 2 !== levels[at + i - 1]! % 2) {
        out.push(
          from === 0 && i === piece.text.length
            ? piece
            : { ...piece, text: piece.text.slice(from, i) }
        );
        from = i;
      }
    if (piece.text.length === 0) out.push(piece);
    at += piece.text.length;
  }
  return out;
}
