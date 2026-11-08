// Line breaks never split a grapheme cluster (UAX #29): emoji sequences, flags, skin-tone
// modifiers, combining marks, and Hangul jamo stay whole, also where a word wider than the line
// breaks by force.

import { describe, expect, test } from 'bun:test';
import { buildPptx } from '../../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '../../../store/index.ts';
import { resolveSlide } from '../../../cascade/index.ts';
import { createFixedMeasurer, layoutSlide } from '../../index.ts';
import { breakOpportunities, forcedLineEnd } from '../line-break.ts';

const FAMILY = '\u{1F468}\u200D\u{1F469}\u200D\u{1F467}';
const FLAG_JP = '\u{1F1EF}\u{1F1F5}';
const FLAG_FR = '\u{1F1EB}\u{1F1F7}';
const THUMB_MEDIUM = '\u{1F44D}\u{1F3FD}';
const E_ACUTE = 'e\u0301';
const HAN = '\u1112\u1161\u11AB'; // U+D55C as conjoining jamo: L V T

/** The break opportunities inside `text`, with `latinLnBrk` on: between every two characters. */
const everyBreak = (text: string): number[] => {
  const b = breakOpportunities(text, true, true);
  const out: number[] = [];
  for (let i = 1; i < text.length; i += 1) if (b[i]) out.push(i);
  return out;
};

/** The UTF-16 offsets where each user-perceived character after the first starts. */
const clusterOffsets = (...parts: string[]): number[] => {
  const out: number[] = [];
  let at = 0;
  for (const p of parts) {
    if (at > 0) out.push(at);
    at += p.length;
  }
  return out;
};

describe('break opportunities keep grapheme clusters whole', () => {
  test('emoji ZWJ sequence', () => {
    expect(everyBreak(`a${FAMILY}b`)).toEqual(clusterOffsets('a', FAMILY, 'b'));
  });

  test('flag pairs: two regional indicators are one flag', () => {
    expect(everyBreak(`${FLAG_JP}${FLAG_FR}`)).toEqual(clusterOffsets(FLAG_JP, FLAG_FR));
  });

  test('skin-tone modifier', () => {
    expect(everyBreak(`a${THUMB_MEDIUM}b`)).toEqual(clusterOffsets('a', THUMB_MEDIUM, 'b'));
  });

  test('combining marks, also two on one base', () => {
    expect(everyBreak(`${E_ACUTE}\u0308x`)).toEqual([3]);
    expect(everyBreak(`a${E_ACUTE}b`)).toEqual(clusterOffsets('a', E_ACUTE, 'b'));
  });

  test('Hangul syllable in conjoining jamo', () => {
    expect(everyBreak(`${HAN}${HAN}`)).toEqual([3]);
  });

  test('an ideographic variation selector stays with its ideograph', () => {
    // Without latinLnBrk, ideographs break per character, but never before the selector.
    const b = breakOpportunities('漢\u{E0100}字');
    expect([1, 2, 3].filter((i) => b[i])).toEqual([3]);
  });

  // GB9c needs Unicode 15.1 in the runtime ICU: an older browser ICU clusters these differently.
  test('Devanagari conjuncts', () => {
    expect(everyBreak('\u0915\u094D\u0937\u0924\u094D\u0930\u093F\u092F')).toEqual([3, 7]);
  });

  test('Thai vowel signs stay with their consonant', () => {
    expect(everyBreak('\u0E01\u0E33\u0E25\u0E31\u0E07\u0E17\u0E33')).toEqual([2, 4, 5]);
  });

  test('three regional indicators', () => {
    expect(everyBreak('\u{1F1EF}\u{1F1F5}\u{1F1EB}')).toEqual([4]);
  });

  test('a carriage return and a line feed', () => {
    expect(everyBreak('a\r\nb')).toEqual([1, 3]);
  });

  test('the rules the cluster check had stay', () => {
    expect(everyBreak('abc')).toEqual([1, 2]);
    // A combining mark after a space belongs to the space.
    expect(everyBreak('a \u0301b')).toEqual([3]);
    // No break on either side of a zero-width joiner (UAX #14 LB8a).
    expect(everyBreak('a\u200Db')).toEqual([]);
    expect(everyBreak('𠀋𠀋')).toEqual([2]);
  });
});

describe('forcedLineEnd: a word wider than the line', () => {
  const cut = (text: string, start: number, at: number): number =>
    forcedLineEnd(text, breakOpportunities(text), start, at);

  test('cuts between letters where no cluster continues', () => {
    expect(cut('abcdefgh', 0, 5)).toBe(5);
    expect(cut('abcdefgh', 3, 3)).toBe(4);
  });

  test('moves the cut back to the start of the cluster that overflows', () => {
    const text = `abc${FAMILY}d`;
    for (let at = 4; at < 3 + FAMILY.length; at += 1) expect(cut(text, 0, at)).toBe(3);
    expect(cut(`ab${FLAG_JP}${FLAG_FR}`, 0, 4)).toBe(2);
    expect(cut(`ab${FLAG_JP}${FLAG_FR}`, 0, 7)).toBe(6);
    expect(cut(`ab${THUMB_MEDIUM}`, 0, 4)).toBe(2);
    expect(cut(`ab${E_ACUTE}c`, 0, 3)).toBe(2);
    expect(cut(`a${HAN}`, 0, 3)).toBe(1);
  });

  test('may cut after a joiner where a cluster ends', () => {
    // An Arabic letter and a zero-width joiner, twenty times: a cluster ends after each joiner.
    expect(cut('\u0628\u200D'.repeat(20), 0, 5)).toBe(4);
    expect(cut('\u{1F468}\u200Dab', 0, 3)).toBe(3);
  });

  test('three regional indicators: one flag and a lone indicator', () => {
    expect(cut('\u{1F1EF}\u{1F1F5}\u{1F1EB}', 0, 5)).toBe(4);
  });

  test('never between a carriage return and a line feed', () => {
    expect(cut('a\r\nb', 0, 2)).toBe(1);
  });

  test('keeps a first cluster wider than the line whole', () => {
    expect(cut(`${FAMILY}x`, 0, 1)).toBe(FAMILY.length);
    expect(cut(`${HAN}x`, 0, 0)).toBe(3);
    expect(cut(`x${THUMB_MEDIUM}y`, 1, 2)).toBe(5);
  });

  test('also without a breakOpportunities result for the text', () => {
    expect(forcedLineEnd(`ab${THUMB_MEDIUM}`, new Uint8Array(7), 0, 4)).toBe(2);
  });
});

describe('a long word breaks by force between clusters (text layout)', () => {
  // At 100 pt, the fixed measurer gives each UTF-16 unit but a low surrogate 50 pt. The shape
  // fits 9 units on a line: eight letters and the first unit of the next cluster.
  const lines = (word: string): string[] => {
    const sp =
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="S"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6000000" cy="4000000"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
      '<p:txBody><a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr><a:lstStyle/>' +
      `<a:p><a:r><a:rPr lang="en-US" sz="10000"/><a:t>${word}</a:t></a:r></a:p></p:txBody></p:sp>`;
    const r = readPptx(buildPptx({ slides: [{ shapes: sp, layout: 'blank' }] }));
    if (!r.ok) throw new Error('open failed');
    const item = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer()).items[0];
    if (item.kind !== 'shape' || !item.text) throw new Error('no text');
    return item.text.lines.map((l) => l.runs.map((g) => g.text).join(''));
  };

  test('a plain long word breaks between letters', () => {
    expect(lines('abcdefghijklmnop')).toEqual(['abcdefghi', 'jklmnop']);
  });

  test.each([
    ['skin-tone modifier', THUMB_MEDIUM],
    ['flag pair', FLAG_JP],
    ['combining marks', `${E_ACUTE}\u0308`],
    ['emoji ZWJ sequence', FAMILY],
    ['Hangul jamo', HAN],
  ])('%s at the line end moves to the next line whole', (_name, cluster) => {
    expect(lines(`aaaaaaaa${cluster}bb`)).toEqual(['aaaaaaaa', `${cluster}bb`]);
  });
});
