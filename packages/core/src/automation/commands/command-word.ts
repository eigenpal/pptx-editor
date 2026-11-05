// The word a command acts on at a collapsed caret: Change Case and Insert link. PPTX Editor
// behavior beside the shared caret word rule (`caretWordRange` in `@eigenpal/ooxml-core`,
// copied from DOCX Editor and kept equal to it). That rule takes a word only with the caret
// strictly inside it, because a format at a word edge arms the text typed next. A command that
// changes existing text has nothing to arm: at the end of a word it takes the word before the
// caret, and at the start of a word the word after it.

import { caretWordRange, type CaretWordRange } from '@eigenpal/ooxml-core';
import type { TextPoint } from '../../contracts/neutral.ts';
import type { OpBatch } from './op-batch.ts';
import { NO_WORD_STOPS, paragraphsOf, slideNumberOf, wordTextOf } from './text-model.ts';

/** The characters of a caret word: the class of `caretWordRange`. */
const WORD_CHARACTER = /^[\p{L}\p{M}\p{N}'\u2019\u001F\u00AD\u200C\u200D]$/u;

/** The character (one code point) that ends at `offset`, or that starts at it (`after`). */
function characterAt(text: string, offset: number, after: boolean): string {
  if (after) return String.fromCodePoint(text.codePointAt(offset) ?? 0);
  const low = text.charCodeAt(offset - 1);
  const pair = low >= 0xdc00 && low <= 0xdfff && offset > 1;
  return text.slice(pair ? offset - 2 : offset - 1, offset);
}

function isWord(text: string, offset: number, after: boolean): boolean {
  if (after ? offset >= text.length : offset <= 0) return false;
  return WORD_CHARACTER.test(characterAt(text, offset, after));
}

/**
 * The word a command reaches at a caret: the word around it (`caretWordRange`), else the word
 * that ends at the caret, else the word that starts at it. Null between words. @internal
 */
export function commandWordRange(
  text: string,
  offset: number,
  stops: ReadonlySet<number>
): CaretWordRange | null {
  const inside = caretWordRange(text, offset, stops);
  if (inside) return inside;
  if (isWord(text, offset, false)) {
    const width = characterAt(text, offset, false).length;
    // One step back is inside the word when it has more than one character; the shared rule
    // then finds its start (and a dictionary script's word). Else the word is one character.
    const word = caretWordRange(text, offset - width, stops);
    return { start: word?.start ?? offset - width, end: offset };
  }
  if (isWord(text, offset, true)) {
    const width = characterAt(text, offset, true).length;
    const word = caretWordRange(text, offset + width, stops);
    return { start: offset, end: word?.end ?? offset + width };
  }
  return null;
}

/**
 * The range of the command word at caret `at` in a text body, or null. @internal
 */
export function commandWordAt(
  batch: OpBatch,
  part: string,
  txBodyId: string,
  at: TextPoint
): { from: TextPoint; to: TextPoint } | null {
  const p = paragraphsOf(batch.element(part, txBodyId))[at.paragraph];
  if (!p) return null;
  const word = commandWordRange(
    wordTextOf(p, slideNumberOf(batch, part)),
    at.offset,
    NO_WORD_STOPS
  );
  if (!word) return null;
  return { from: { ...at, offset: word.start }, to: { ...at, offset: word.end } };
}
