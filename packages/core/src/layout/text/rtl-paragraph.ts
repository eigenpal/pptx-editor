// Adapted from DOCX Editor `packages/core/src/layout/rtl-paragraph.ts` at commit adf0e178, with
// `rtl-letter.ts`, `arabic-script-language.ts`, and `line-end-whitespace.ts` of the same commit.
// The source reads run directions from `w:rtl`, resolves each run-direction group on its own,
// and works on styled spans. Slide runs have no direction, so a paragraph resolves as one
// UAX #9 paragraph (`runEdgeLevels` keeps the neutrals at the edges of a run that reads
// against the paragraph inside that run), and this module works on the UTF-16 units of a
// paragraph.

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
 * Keep the punctuation at the edges of a run that reads against the paragraph inside that run:
 * the characters before its first strong character and after its last one, other than
 * whitespace, take the level of that strong character. In the `rtl-text` reference PDF, a
 * left-to-right run of a right-to-left paragraph keeps its colon and its quotation marks inside
 * it (`An English phrase in a right-to-left paragraph: `, `"Hello, world!"`). Whitespace keeps
 * its UAX #9 level, so it still separates the run from the text beside it.
 */
function runEdgeLevels(levels: Uint8Array, runs: readonly BidiRunText[], rtl: boolean): void {
  let offset = 0;
  for (const run of runs) {
    const start = offset;
    offset += run.text.length;
    if (runIsRtl(run.text, rtl) === rtl) continue;
    const strong: number[] = [];
    for (let i = 0; i < run.text.length;) {
      const cp = run.text.codePointAt(i)!;
      if (strongRtl(String.fromCodePoint(cp)) !== undefined) strong.push(i);
      i += cp > 0xffff ? 2 : 1;
    }
    const first = strong[0];
    const last = strong[strong.length - 1];
    if (first === undefined || last === undefined) continue;
    const take = (i: number, from: number): void => {
      if (!/\s/u.test(run.text[i]!)) levels[start + i] = levels[start + from]!;
    };
    for (let i = 0; i < first; i++) take(i, first);
    const lastEnd = last + (run.text.codePointAt(last)! > 0xffff ? 2 : 1);
    for (let i = lastEnd; i < run.text.length; i++) take(i, last);
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
  // Authored Unicode controls decide on their own.
  if (!controls) runEdgeLevels(levels, runs, rtl);
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
