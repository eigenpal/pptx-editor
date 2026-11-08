// The text range of a threaded comment. A text moniker list (`ac:txMkLst`) names the slide, the
// shape (`ac:spMk`), and then the range as a text moniker, `ac:txMk cp len`: the first character
// and the number of characters in the shape's text. The text of a shape counts as the editor
// counts it (`text-model.ts`), with one character between two paragraphs. The editor shows a
// range as two `TextPoint`s, as a text selection does.
//
// `ac:txMk` can hold `ac:context len hash`, a check of the text that the range was made on. A
// new range is written without it. `store/store/comment-anchors.ts` moves a range with its text
// and removes the check, which no longer describes the text.

import {
  child,
  presentationInfo,
  type Element,
  type ParagraphText,
  type PptxPackage,
} from '../../store/index.ts';
import type { CommentTextRange, TextPoint } from '../../contracts/neutral.ts';
import { paragraphLength, paragraphText, paragraphsOf } from './text-model.ts';

/** The paragraph lengths of the text of `shape` (`p:sp`), or null when it has no text body. */
export function shapeParagraphLengths(
  shape: Element | undefined,
  slideNumber?: number
): number[] | null {
  if (!shape || shape.local !== 'sp') return null;
  const body = child(shape, 'p:txBody');
  if (!body) return null;
  return paragraphsOf(body).map((p) => paragraphLength(p, slideNumber));
}

/** The character of `point` in the text whose paragraphs are `lengths` long. */
function charOf(lengths: readonly number[], point: TextPoint): number {
  let c = 0;
  const last = Math.max(0, Math.min(point.paragraph, lengths.length - 1));
  for (let i = 0; i < last; i += 1) c += lengths[i]! + 1;
  return c + Math.max(0, Math.min(point.offset, lengths[last] ?? 0));
}

/**
 * The point of character `c` in the text whose paragraphs are `lengths` long. A `start` at the
 * end of a paragraph is the start of the next one, so a range never starts on a line it does
 * not cover.
 */
function pointOf(lengths: readonly number[], c: number, start: boolean): TextPoint {
  let rest = c;
  for (let i = 0; i < lengths.length; i += 1) {
    const atEnd = rest === lengths[i]! && start && i < lengths.length - 1;
    if (rest <= lengths[i]! && !atEnd) return { paragraph: i, offset: rest };
    rest -= lengths[i]! + 1;
  }
  const last = Math.max(0, lengths.length - 1);
  return { paragraph: last, offset: lengths[last] ?? 0 };
}

/** The whole text's length, with one character between two paragraphs. */
export const textLength = (lengths: readonly number[]): number =>
  lengths.reduce((n, l) => n + l, 0) + Math.max(0, lengths.length - 1);

/**
 * `range` as a text moniker's first character and length, in the bounds of the text. Null when
 * the range is empty in the text.
 */
export function rangeChars(
  lengths: readonly number[],
  range: CommentTextRange
): { readonly cp: number; readonly len: number } | null {
  const a = charOf(lengths, range.start);
  const b = charOf(lengths, range.end);
  const cp = Math.min(a, b);
  const len = Math.abs(b - a);
  return len > 0 ? { cp, len } : null;
}

/**
 * The range of a text moniker (`cp`, `len`, as the file gave them) in the text, cut to its
 * bounds. Null for values that are not counts and for a range with no text left in the bounds.
 */
export function rangeOfChars(
  lengths: readonly number[],
  cp: string | undefined,
  len: string | undefined
): CommentTextRange | null {
  const count = (v: string | undefined): number | undefined =>
    v !== undefined && /^\s*\d{1,9}\s*$/.test(v) ? Number(v.trim()) : undefined;
  const first = count(cp);
  const n = count(len);
  if (first === undefined || n === undefined || lengths.length === 0) return null;
  const total = textLength(lengths);
  const start = Math.min(first, total);
  const end = Math.min(first + n, total);
  if (end <= start) return null;
  return { start: pointOf(lengths, start, true), end: pointOf(lengths, end, false) };
}

/** The slide number of each slide part of a package, found once per package. */
const slideNumbers = new WeakMap<PptxPackage, ReadonlyMap<string, number>>();

/**
 * A slide paragraph's text as the editor counts it, for the store (`paragraphText` of
 * `DocumentStoreOptions`): a `slidenum` field shows its slide's number, a math zone counts one
 * character, and a line break counts one character.
 */
export const editorParagraphText: ParagraphText = (p, pkg, part) => {
  let byPart = slideNumbers.get(pkg);
  if (!byPart) {
    const found = new Map<string, number>();
    presentationInfo(pkg).slides.forEach(
      (s, i) => found.has(s.partName) || found.set(s.partName, i + 1)
    );
    slideNumbers.set(pkg, (byPart = found));
  }
  return paragraphText(p, byPart.get(part)).replaceAll('\n', '\u000b');
};
