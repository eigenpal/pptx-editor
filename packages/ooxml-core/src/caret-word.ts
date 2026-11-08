// Adapted from docx-editor/packages/core/src/layout/caret-word.ts at 5993bd223 (EigenPal).
// The format-neutral part only: `CaretWordRange` and `caretWordRange`, with the helpers they
// use, unchanged. The layout and revision readers of that module (`hiddenRevisionRanges`,
// `wordTextFromLayout`, `caretWordAt`, `caretWordSelection`) stay in DOCX Editor.
// Keep these lines equal to that file.
// The word a caret format or a format painter click reaches.
//
// A format picked at a caret strictly inside a word formats the whole word; at a word's edge
// it only arms the next typed text. The word here is letters, digits, combining marks,
// apostrophes, optional hyphens, and zero-width joiners: a hyphen, an underscore, a period, or
// a comma ends it, and so does any object a run holds, such as a picture or a note reference.

import { segmentWords } from './word-segment.ts';

// U+001F is the model's optional hyphen; U+00AD, ZWNJ and ZWJ also stay inside a word.
const CARET_WORD_CHARACTER = new RegExp(
  "^[\\p{L}\\p{M}\\p{N}'\\u2019\\u001F\\u00AD\\u200C\\u200D]$",
  'u'
);
/** Scripts written without spaces between words, segmented by dictionary. */
const DICTIONARY_SCRIPT = /[\p{Script=Thai}\p{Script=Lao}\p{Script=Khmer}\p{Script=Myanmar}]/u;

/** A range of canonical offsets: `start` inclusive, `end` exclusive. */
export interface CaretWordRange {
  readonly start: number;
  readonly end: number;
}

/** Whether the UTF-16 unit at `index` belongs to a word character, surrogate pairs whole. */
function isWordAt(text: string, index: number): boolean {
  const unit = text.charCodeAt(index);
  if (Number.isNaN(unit)) return false;
  // A low surrogate belongs to the code point that starts one unit earlier.
  const start = unit >= 0xdc00 && unit <= 0xdfff && index > 0 ? index - 1 : index;
  const codePoint = text.codePointAt(start);
  return codePoint !== undefined && CARET_WORD_CHARACTER.test(String.fromCodePoint(codePoint));
}

/**
 * The word a format at a collapsed caret reaches, or null when the caret is not inside one.
 *
 * Inside means a word character on both sides of the caret, with no stop at the caret. A
 * stop (an inline control edge) also ends the word on either side.
 */
export function caretWordRange(
  text: string,
  offset: number,
  stops: ReadonlySet<number>
): CaretWordRange | null {
  const isWord = (index: number) => isWordAt(text, index);
  if (offset <= 0 || offset >= text.length || stops.has(offset)) return null;
  if (!isWord(offset - 1) || !isWord(offset)) return null;
  let start = offset - 1;
  while (start > 0 && isWord(start - 1) && !stops.has(start)) start -= 1;
  let end = offset + 1;
  while (end < text.length && isWord(end) && !stops.has(end)) end += 1;
  const narrowed = dictionaryWord(text, start, end, offset);
  return narrowed === undefined ? { start, end } : narrowed;
}

/**
 * Narrow a run of dictionary-script text to the dictionary word holding the caret.
 *
 * Undefined when the run has no such script; null when the caret falls on a word boundary
 * inside the run, which is an edge rather than the inside of a word.
 */
function dictionaryWord(
  text: string,
  start: number,
  end: number,
  offset: number
): CaretWordRange | null | undefined {
  const run = text.slice(start, end);
  if (!DICTIONARY_SCRIPT.test(run)) return undefined;
  for (const segment of segmentWords(run)) {
    const segmentStart = start + segment.utf16From;
    const segmentEnd = start + segment.utf16To;
    if (offset === segmentStart) return null;
    if (offset > segmentStart && offset < segmentEnd) {
      return { start: segmentStart, end: segmentEnd };
    }
  }
  return null;
}
