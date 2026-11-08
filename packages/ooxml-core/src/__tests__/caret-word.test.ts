// Tests for the caret word rule, adapted from DOCX Editor
// (`packages/core/src/layout/caret-word.ts` at 5993bd223): the word a format at a collapsed
// caret reaches.

import { describe, expect, test } from 'bun:test';
import { caretWordRange } from '../caret-word.ts';

const NONE: ReadonlySet<number> = new Set();

describe('caretWordRange', () => {
  test('a caret strictly inside a word reaches the whole word', () => {
    expect(caretWordRange('alpha beta', 2, NONE)).toEqual({ start: 0, end: 5 });
    expect(caretWordRange('alpha beta', 8, NONE)).toEqual({ start: 6, end: 10 });
  });

  test('a caret at a word edge, between words, or at the text ends reaches nothing', () => {
    expect(caretWordRange('alpha beta', 0, NONE)).toBeNull();
    expect(caretWordRange('alpha beta', 5, NONE)).toBeNull();
    expect(caretWordRange('alpha beta', 6, NONE)).toBeNull();
    expect(caretWordRange('alpha beta', 10, NONE)).toBeNull();
  });

  test('a hyphen, an underscore, and a period end a word; an apostrophe does not', () => {
    expect(caretWordRange('well-known', 2, NONE)).toEqual({ start: 0, end: 4 });
    expect(caretWordRange('snake_case', 7, NONE)).toEqual({ start: 6, end: 10 });
    expect(caretWordRange('a.bc', 3, NONE)).toEqual({ start: 2, end: 4 });
    expect(caretWordRange('don’t', 2, NONE)).toEqual({ start: 0, end: 5 });
  });

  test('combining marks, ZWJ, and astral letters stay inside the word', () => {
    // "e" + COMBINING ACUTE ACCENT, then "t".
    expect(caretWordRange('ét', 1, NONE)).toEqual({ start: 0, end: 3 });
    // MATHEMATICAL BOLD SMALL A (a surrogate pair) between Latin letters.
    expect(caretWordRange('a\u{1D41A}b', 1, NONE)).toEqual({ start: 0, end: 4 });
  });

  test('an object character ends a word', () => {
    expect(caretWordRange('ab￼cd', 1, NONE)).toEqual({ start: 0, end: 2 });
  });

  test('a stop ends the word on either side and refuses a caret on it', () => {
    const stops = new Set([3]);
    expect(caretWordRange('abcdef', 1, stops)).toEqual({ start: 0, end: 3 });
    expect(caretWordRange('abcdef', 4, stops)).toEqual({ start: 3, end: 6 });
    expect(caretWordRange('abcdef', 3, stops)).toBeNull();
  });

  test('Thai text narrows to the dictionary word that holds the caret', () => {
    // "สวัสดีครับ" is two dictionary words: "สวัสดี" (6 units) and "ครับ" (4 units).
    const thai = 'สวัสดีครับ';
    expect(caretWordRange(thai, 2, NONE)).toEqual({ start: 0, end: 6 });
    expect(caretWordRange(thai, 8, NONE)).toEqual({ start: 6, end: 10 });
    // The caret between the two dictionary words is an edge.
    expect(caretWordRange(thai, 6, NONE)).toBeNull();
  });
});
