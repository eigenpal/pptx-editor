// Text units for editing: clean input text, and the caret stops of a paragraph. A caret steps
// and deletes by user-perceived character (grapheme cluster), so it never stops inside an
// emoji or between a letter and its combining mark. Fields and line breaks are one unit each.
// The clusters and the words (UAX #29, with dictionary words in Thai, Lao, Khmer, Myanmar, and
// CJK text) come from the segmentation that DOCX Editor shares through `@eigenpal/ooxml-core`.

import { segmentGraphemes, segmentWords } from '@eigenpal/ooxml-core';
import type { ResolvedParagraph } from '../../cascade/index.ts';

/**
 * Text as it can go into a slide: CRLF and CR become LF, and characters XML 1.0 cannot hold
 * (control characters, U+FFFE, U+FFFF, unpaired surrogates) are removed. U+000B, the line
 * break that other editors put on the clipboard, is kept for the caller to map.
 */
export function cleanText(raw: string): string {
  const text = raw.replace(/\r\n?/g, '\n');
  let out = '';
  for (let i = 0; i < text.length; i += 1) {
    const c = text.charCodeAt(i);
    if (c >= 0xd800 && c <= 0xdbff) {
      const next = text.charCodeAt(i + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        out += text[i] + text[i + 1];
        i += 1;
      }
      continue;
    }
    if (c >= 0xdc00 && c <= 0xdfff) continue;
    if (c < 0x20 && c !== 0x9 && c !== 0xa && c !== 0xb) continue;
    if (c === 0xfffe || c === 0xffff) continue;
    out += text[i];
  }
  return out;
}

/** Starts of the grapheme clusters of `text` (UAX #29, through the shared boundary). */
function clusterStarts(text: string): number[] {
  return segmentGraphemes(text).map((s) => s.utf16From);
}

/** One unit of a paragraph for word steps: a word, a field, or anything between words. */
export interface WordSegment {
  readonly start: number;
  readonly end: number;
  readonly word: boolean;
  /** White space only: a word step forward passes it with the word before it. */
  readonly space?: boolean;
}

function wordSegments(text: string, base: number): WordSegment[] {
  return segmentWords(text).map((s) => ({
    start: base + s.utf16From,
    end: base + s.utf16To,
    word: s.wordLike,
    ...(/^[^\S\n]+$/u.test(text.slice(s.utf16From, s.utf16To)) ? { space: true } : {}),
  }));
}

/** The caret stops and word segments of one paragraph, in source offsets. */
export interface ParagraphUnits {
  /** Every offset a caret may sit at, ascending, from 0 to the paragraph length. */
  readonly stops: readonly number[];
  readonly words: readonly WordSegment[];
  readonly length: number;
}

/**
 * Units of a laid-out paragraph: fields and breaks whole, and each stretch of text runs by
 * cluster and word over the stretch's whole text, so a word or a cluster that crosses a run
 * boundary (a bold first letter, a combining mark in its own run) stays one unit.
 */
export function paragraphUnits(p: ResolvedParagraph | undefined): ParagraphUnits {
  const stops = new Set<number>([0]);
  const segs: WordSegment[] = [];
  let at = 0;
  let stretch = '';
  const flush = (): void => {
    const base = at - stretch.length;
    for (const s of clusterStarts(stretch)) stops.add(base + s);
    for (const w of wordSegments(stretch, base)) segs.push(w);
    stretch = '';
    stops.add(at);
  };
  for (const run of p?.runs ?? []) {
    if (run.kind === 'text') {
      stretch += run.text;
      at += run.text.length;
      continue;
    }
    flush();
    const len = run.kind === 'break' ? 1 : run.text.length;
    segs.push({ start: at, end: at + len, word: run.kind === 'field' });
    at += len;
    stops.add(at);
  }
  flush();
  return { stops: [...stops].sort((a, b) => a - b), words: segs, length: at };
}

/** The stop before `offset`, or 0. */
export function stopBefore(u: ParagraphUnits, offset: number): number {
  let best = 0;
  for (const s of u.stops) if (s < offset) best = s;
  return best;
}

/** The stop after `offset`, or the paragraph length. */
export function stopAfter(u: ParagraphUnits, offset: number): number {
  for (const s of u.stops) if (s > offset) return s;
  return u.length;
}

/** `offset` moved to the nearest stop at or before it (out of a cluster or a field). */
export function snapToStop(u: ParagraphUnits, offset: number): number {
  const clamped = Math.max(0, Math.min(offset, u.length));
  let best = 0;
  for (const s of u.stops) if (s <= clamped) best = s;
  return best;
}

/** Where a word deletion backward from `offset` stops: past the gap, then one word. */
export function wordStartBefore(u: ParagraphUnits, offset: number): number {
  let pos = offset;
  const segs = u.words.filter((s) => s.start < offset);
  let i = segs.length - 1;
  while (i >= 0 && !segs[i].word) pos = segs[i--].start;
  if (i >= 0) pos = segs[i].start;
  return pos;
}

/** Where a word deletion forward from `offset` stops: past the gap, then one word. */
export function wordEndAfter(u: ParagraphUnits, offset: number): number {
  let pos = offset;
  const segs = u.words.filter((s) => s.end > offset);
  let i = 0;
  while (i < segs.length && !segs[i].word) pos = segs[i++].end;
  if (i < segs.length) pos = segs[i].end;
  return pos;
}

/**
 * Where a word step forward stops: past the gap, one word, and the spaces after it, so the
 * caret lands at the start of the next word. Ctrl or Alt with Right, Shift for a selection,
 * and a word deletion forward take this step.
 */
export function nextWordStart(u: ParagraphUnits, offset: number): number {
  // At a word end, the spaces after it are the step: the caret lands on the next word.
  const gap = u.words.find((s) => s.start === offset && s.space);
  if (gap) return gap.end;
  let pos = wordEndAfter(u, offset);
  for (const s of u.words) if (s.start === pos && s.space) pos = s.end;
  return pos;
}

/** The word segments (UAX #29) of plain text, such as `wordTextOf` gives for a paragraph. */
export function textWordSegments(text: string): WordSegment[] {
  return wordSegments(text, 0);
}

/**
 * The segment a double-click at `offset` takes: the word segment (UAX #29) on the right of the
 * offset, or the word on its left when the right is not a word, so a click at a word's end
 * takes that word. Between words it is the run of spaces or punctuation there.
 */
export function wordSegmentAt(
  words: readonly WordSegment[],
  offset: number
): WordSegment | undefined {
  const right = words.find((s) => s.start <= offset && offset < s.end);
  const left = words.find((s) => s.start < offset && offset <= s.end);
  return right && (right.word || !left?.word) ? right : left;
}

/** The range a double-click at `offset` selects, as `[start, end)` (`wordSegmentAt`). */
export function wordSelectionAt(u: ParagraphUnits, offset: number): [number, number] {
  const seg = wordSegmentAt(u.words, offset);
  return seg ? [seg.start, seg.end] : [offset, offset];
}
