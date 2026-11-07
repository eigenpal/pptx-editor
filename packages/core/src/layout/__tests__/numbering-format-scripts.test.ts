// Adapted from DOCX Editor `packages/core/src/layout/__tests__/numbering-format-scripts.test.ts`
// at commit adf0e178. The tests call the script formatters directly: PPTX Editor does not copy
// `numbering-format.ts`, and `hebrew1` is not a presentation scheme.

import { describe, expect, test } from 'bun:test';
import {
  formatArabicAbjad,
  formatArabicAlpha,
  formatHebrew1,
  formatHebrew2,
  formatHindiNumbers,
} from '../text/numbering-format-scripts.ts';

// Expected strings are the reference renderer's own list strings for these values.
const RLM = '‏';
const ZWNJ = '‌';

describe('Hebrew, Arabic and Devanagari numbering formats', () => {
  test('`hebrew1` writes Hebrew numerals, with 15 and 16 as טו and טז', () => {
    const cases: [number, string][] = [
      [1, 'א'],
      [10, 'י'],
      [11, 'יא'],
      [15, 'טו'],
      [16, 'טז'],
      [17, 'יז'],
      [40, 'מ'],
      [99, 'צט'],
      [101, 'קא'],
      [115, 'קטו'],
      [116, 'קטז'],
      [272, 'רעב'],
      [301, 'שא'],
    ];
    for (const [value, text] of cases) expect(formatHebrew1(value)).toBe(text);
  });

  test('`hebrew2` repeats the last letter past the alphabet', () => {
    expect(formatHebrew2(1)).toBe(`${RLM}א`);
    expect(formatHebrew2(22)).toBe(`${RLM}ת`);
    expect(formatHebrew2(23)).toBe(`${RLM}תא`);
    expect(formatHebrew2(32)).toBe(`${RLM}תי`);
    expect(formatHebrew2(98)).toBe(`${RLM}תתתתי`);
  });

  test('`arabicAlpha` and `arabicAbjad` repeat the letter each pass', () => {
    expect(formatArabicAlpha(1)).toBe(`أ${ZWNJ}`);
    expect(formatArabicAlpha(28)).toBe(`ي${ZWNJ}`);
    expect(formatArabicAlpha(29)).toBe(`أ${ZWNJ}أ${ZWNJ}`);
    expect(formatArabicAbjad(3)).toBe(`${ZWNJ}ج`);
    expect(formatArabicAbjad(28)).toBe(`${ZWNJ}غ`);
    expect(formatArabicAbjad(30)).toBe(`${ZWNJ}ب${ZWNJ}ب`);
  });

  test('the letter sequences restart after 392', () => {
    expect(formatHebrew1(398)).toBe('ו');
    expect(formatHebrew1(998)).toBe('ריד');
    expect(formatHebrew2(498)).toBe(`${RLM}תתתתצ`);
    expect(formatArabicAlpha(398)).toBe(`ح${ZWNJ}`);
    expect(formatArabicAbjad(398)).toBe(`${ZWNJ}و`);
  });

  test('`hindiNumbers` writes the decimal value in Devanagari digits', () => {
    expect(formatHindiNumbers('7')).toBe('७');
    expect(formatHindiNumbers('1001')).toBe('१००१');
  });
});
