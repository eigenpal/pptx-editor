// Adapted from DOCX Editor `packages/core/src/layout/rtl-paragraph.ts` at commit adf0e178, with
// `rtl-letter.ts`, `arabic-script-language.ts`, and `line-end-whitespace.ts` of the same commit.
// The source reads run directions from `w:rtl` and works on styled spans; slide text has no
// run direction, so a run takes the direction of its first strong character, and this module
// works on the UTF-16 units of a paragraph.

import type { ResolvedParagraph } from '../../cascade/index.ts';
import { bidiAlgorithm } from './bidi.ts';

/**
 * Hebrew, Arabic, Syriac, Thaana, N'Ko, Samaritan, Mandaic and the Arabic extensions, with
 * their presentation forms: the letters that start right-to-left text.
 */
export const RIGHT_TO_LEFT_LETTER = /[֐-ࣿיִ-﷿ﹰ-﻿]/u;

/** Languages written in Arabic script, by primary subtag. */
const ARABIC_SCRIPT_LANGUAGES = new Set(['ar', 'fa', 'ur', 'ps', 'sd', 'ug', 'ckb', 'ks']);

/** Whether a run's language (`a:rPr/@lang`, BCP 47) is written in Arabic script. */
export function arabicScriptLanguage(lang: string | undefined): boolean {
  return lang !== undefined && ARABIC_SCRIPT_LANGUAGES.has(lang.toLowerCase().split('-')[0]!);
}

const HAS_BIDI_CONTROLS = /[؜‎‏‪-‮⁦-⁩]/u;

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
  /**
   * The run-direction group of each unit: 1 for a right-to-left run. Undefined when the text
   * has Unicode bidi controls, which keep the full-paragraph UAX #9 path.
   */
  readonly groups?: Uint8Array;
}

/** The direction of the first strong character (UAX #9 P2 and P3), or undefined with none. */
function firstStrongRtl(text: string): boolean | undefined {
  for (const ch of text) {
    const type = bidiAlgorithm.getBidiCharTypeName(ch);
    if (type === 'L') return false;
    if (type === 'R' || type === 'AL') return true;
  }
  return undefined;
}

/**
 * The direction a run reads in. Slide runs have no direction of their own, so a run reads in
 * the direction of its first strong character, and a run with none in the paragraph's. In the
 * `rtl-text` reference PDF, a left-to-right run keeps its colon and its quotation marks inside
 * it (`An English phrase in a right-to-left paragraph: `, `"Hello, world!"`), and a period run
 * after a left-to-right run stands on the paragraph's side (`ECMA-376` then `.`).
 */
export function runIsRtl(text: string, paragraphRtl: boolean): boolean {
  return firstStrongRtl(text) ?? paragraphRtl;
}

/**
 * Resolve paragraph-wide levels before wrapping; source offsets remain logical. Undefined for a
 * left-to-right paragraph without right-to-left letters, which lays out unchanged.
 *
 * Contiguous runs with the same direction share one directional context, as in the source.
 */
export function paragraphBidi(
  runs: readonly BidiRunText[],
  rtl: boolean
): ParagraphBidi | undefined {
  const text = runs.map((run) => run.text).join('');
  if (!rtl && !RIGHT_TO_LEFT_LETTER.test(text)) return undefined;
  const direction = rtl ? 'rtl' : 'ltr';
  // Keep authored Unicode controls on the established full-paragraph UAX #9 path.
  if (HAS_BIDI_CONTROLS.test(text))
    return { rtl, levels: bidiAlgorithm.getEmbeddingLevels(text, direction).levels };
  const levels = new Uint8Array(text.length);
  const groups = new Uint8Array(text.length);
  let offset = 0;
  for (let start = 0; start < runs.length;) {
    const groupRtl = runIsRtl(runs[start]!.text, rtl);
    let end = start;
    let length = 0;
    while (end < runs.length && runIsRtl(runs[end]!.text, rtl) === groupRtl) {
      length += runs[end]!.text.length;
      end++;
    }
    const group = text.slice(offset, offset + length);
    const resolved = bidiAlgorithm.getEmbeddingLevels(
      groupRtl && arabicScriptLanguage(runs[start]!.lang)
        ? group.replace(RTL_RUN_DIGIT_SLASH, NEUTRAL_SLASH)
        : group,
      groupRtl ? 'rtl' : 'ltr'
    );
    levels.set(resolved.levels, offset);
    if (groupRtl) groups.fill(1, offset, offset + length);
    offset += length;
    start = end;
  }
  return { rtl, levels, groups };
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

/** Spaces that hang at a line end (DOCX Editor `isCollapsibleLineEndWhitespace`). */
export function isCollapsibleSpace(code: number): boolean {
  return code === 32 || code === 0x3000 || code === 0x200b;
}

/** UAX #9 classes that rule L1 resets at a line end: whitespace, breaks, and tabs. */
function isLineEndClass(code: number): boolean {
  const type = bidiAlgorithm.getBidiCharTypeName(String.fromCharCode(code));
  return type === 'WS' || type === 'B' || type === 'S';
}

/** Where the line-end whitespace that UAX #9 L1 resets starts in `codes[start, end)`. */
function lineEndWhitespaceStart(codes: ArrayLike<number>, start: number, end: number): number {
  let trailing = end;
  while (trailing > start && isLineEndClass(codes[trailing - 1]!)) trailing--;
  return trailing;
}

/** Whether `codes[index]` is a line break after line-end spaces. */
function breakAfterSpaces(codes: ArrayLike<number>, index: number): boolean {
  let before = index;
  while (before > 0 && codes[before] === 10) before--;
  return before < index && isCollapsibleSpace(codes[before]!);
}

function bidiOrder(
  codes: ArrayLike<number>,
  levels: number[],
  start: number,
  end: number,
  rtl: boolean,
  spacesKeepLevel = false
): number[] {
  for (let index = lineEndWhitespaceStart(codes, start, end); index < end; index++) {
    if (spacesKeepLevel && isCollapsibleSpace(codes[index]!)) continue;
    levels[index] =
      spacesKeepLevel && index > start && breakAfterSpaces(codes, index)
        ? levels[index - 1]!
        : rtl
          ? 1
          : 0;
  }
  return visualOrderOfLevels(levels, start, end);
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

/** One line in visual order: line-relative unit indices, and each unit's level after L1. */
export interface BidiLineOrder {
  readonly order: readonly number[];
  readonly levels: readonly number[];
}

/**
 * UAX #9 L1/L2 for the units `[start, end)` of a paragraph (`reorderBidiSpans` in the source).
 * `codes` holds the UTF-16 units of the line.
 *
 * L1 puts line-ending tabs and breaks back in the paragraph context, outside the run-direction
 * group whose visible text precedes them. Line-ending spaces keep their run's group and their
 * resolved level, so they stay beside the text they follow. Groups stand in paragraph order,
 * and each group resolves its own order, so a numeric left-to-right group stays intact beside
 * a right-to-left run.
 */
export function lineVisualOrder(
  bidi: ParagraphBidi,
  codes: ArrayLike<number>,
  start: number,
  end: number
): BidiLineOrder {
  const levels = Array.from(bidi.levels.subarray(start, end));
  const count = end - start;
  const rtl = bidi.rtl;
  // UAX #9 L1: a tab, and the whitespace before it, take the paragraph level.
  for (let index = 0; index < count; index++) {
    if (codes[index] !== 9) continue;
    levels[index] = rtl ? 1 : 0;
    for (let ws = index - 1; ws >= 0 && isLineEndClass(codes[ws]!) && codes[ws] !== 9; ws--)
      levels[ws] = rtl ? 1 : 0;
  }
  const groupsOf = bidi.groups;
  if (!groupsOf) return { order: bidiOrder(codes, levels, 0, count, rtl), levels };
  const trailing = lineEndWhitespaceStart(codes, 0, count);
  const directionAt = (index: number): boolean => {
    // A break after the spaces stays beside them, so the caret before it does too.
    if (index >= trailing && breakAfterSpaces(codes, index)) return directionAt(index - 1);
    return index >= trailing && !isCollapsibleSpace(codes[index]!)
      ? rtl
      : groupsOf[start + index] === 1;
  };
  const groups: { start: number; end: number; rtl: boolean }[] = [];
  for (let from = 0; from < count;) {
    const direction = directionAt(from);
    let to = from + 1;
    while (to < count && directionAt(to) === direction) to++;
    groups.push({ start: from, end: to, rtl: direction });
    from = to;
  }
  if (rtl) groups.reverse();
  const order: number[] = [];
  for (const group of groups)
    for (const index of bidiOrder(codes, levels, group.start, group.end, group.rtl, true))
      order.push(index);
  return { order, levels };
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
