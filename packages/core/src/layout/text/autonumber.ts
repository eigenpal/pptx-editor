// `a:buAutoNum` schemes (ECMA-376 ST_TextAutonumberScheme): the number format plus the
// punctuation around it, and the counters that number a text body's paragraphs.

import type { ResolvedParagraph } from '../../cascade/index.ts';
import {
  formatArabicAbjad,
  formatArabicAlpha,
  formatHebrew2,
  formatHindiNumbers,
} from './numbering-format-scripts.ts';
import { mapSymbolBulletText } from './symbol-encoding.ts';

function roman(n: number): string {
  const table: [number, string][] = [
    [1000, 'm'],
    [900, 'cm'],
    [500, 'd'],
    [400, 'cd'],
    [100, 'c'],
    [90, 'xc'],
    [50, 'l'],
    [40, 'xl'],
    [10, 'x'],
    [9, 'ix'],
    [5, 'v'],
    [4, 'iv'],
    [1, 'i'],
  ];
  let out = '';
  let rest = Math.min(Math.max(1, n), 3999);
  for (const [v, s] of table)
    while (rest >= v) {
      out += s;
      rest -= v;
    }
  return out;
}

/**
 * The letters of `sequence` as digits: 1 is the first letter, and past the last come pairs
 * (Latin 27 is `aa`, 28 is `ab`). The standard gives no rule past the last letter. Latin, Thai,
 * and Hindi letters use this rule. The Arabic and Hebrew counters of DOCX Editor repeat a
 * letter instead (Arabic 29 is `أأ`, Hebrew 23 is `תא`).
 */
function letters(n: number, sequence: readonly string[]): string {
  let out = '';
  let rest = Math.max(1, n);
  while (rest > 0) {
    rest -= 1;
    out = sequence[rest % sequence.length]! + out;
    rest = Math.floor(rest / sequence.length);
  }
  return out;
}

const LATIN_LETTERS = [...'abcdefghijklmnopqrstuvwxyz'];

function alpha(n: number): string {
  return letters(n, LATIN_LETTERS);
}

/** The decimal digits of `n`, written with the ten characters of `digits` (zero first). */
function digitsOf(n: number, digits: string): string {
  return String(n).replace(/[0-9]/g, (d) => digits[Number(d)]!);
}

const FULL_WIDTH_DIGITS = '０１２３４５６７８９';
const THAI_DIGITS = '๐๑๒๓๔๕๖๗๘๙';
const IDEOGRAPHIC_DIGITS = '〇一二三四五六七八九';

// The Thai consonants without U+0E03, U+0E05, and U+0E06: the `thai-alphabetic` sequence of
// the W3C Ready-made Counter Styles.
const THAI_LETTERS = [...'กขคงจฉชซฌญฎฏฐฑฒณดตถทธนบปผฝพฟภมยรลวศษสหฬอฮ'];
// `hindiAlphaPeriod` counts with the Devanagari vowels, and `hindiAlpha1Period` with the
// consonants.
const HINDI_VOWELS = [...'अआइईउऊऋएऐओऔ'];
const HINDI_CONSONANTS = [...'कखगघङचछजझञटठडढणतथदधनपफबभमयरलवशषसह'];

/**
 * Ideographic numbers. Type A counts with 十 (11 is 十一, 99 is 九十九). Type C writes one
 * ideograph for each decimal digit (100 is 一〇〇). `typeAUpTo` is the last number in type A.
 */
function ideographic(n: number, typeAUpTo: number): string {
  if (n > typeAUpTo || n > 99) return digitsOf(n, IDEOGRAPHIC_DIGITS);
  const tens = Math.floor(n / 10);
  const ones = n % 10;
  const ten = tens === 0 ? '' : tens === 1 ? '十' : `${IDEOGRAPHIC_DIGITS[tens]!}十`;
  return ten + (ones === 0 ? '' : IDEOGRAPHIC_DIGITS[ones]!);
}

/** Circled numbers 1 to 10 from the code point of 1, and plain digits after 10. */
function circled(n: number, one: number): string {
  return n <= 10 ? String.fromCodePoint(one + n - 1) : String(n);
}

/** The counter of each scheme family: the scheme name without its punctuation. */
const COUNTERS: ReadonlyMap<string, (n: number) => string> = new Map([
  ['alphaLc', alpha],
  ['alphaUc', (n: number) => alpha(n).toUpperCase()],
  ['arabic', String],
  ['romanLc', roman],
  ['romanUc', (n: number) => roman(n).toUpperCase()],
  // ① (U+2460) to ⑩. The Wingdings schemes draw their Unicode twins, as symbol-font
  // bullets do where the symbol font is not installed.
  ['circleNumDb', (n: number) => circled(n, 0x2460)],
  ['circleNumWdWhite', (n: number) => circled(n, 0x2460)],
  // The negative circled digits ❶ (U+2776) to ❿.
  ['circleNumWdBlack', (n: number) => circled(n, 0x2776)],
  ['arabicDb', (n: number) => digitsOf(n, FULL_WIDTH_DIGITS)],
  // Simplified Chinese and Japanese: type A to 99. Traditional Chinese: type A to 19.
  // Japanese and Korean: type C from 1.
  ['ea1Chs', (n: number) => ideographic(n, 99)],
  ['ea1JpnChs', (n: number) => ideographic(n, 99)],
  ['ea1Cht', (n: number) => ideographic(n, 19)],
  ['ea1JpnKor', (n: number) => ideographic(n, 0)],
  // Arabic letters in alphabetical (1) and abjad (2) order, and Hebrew letters (2): the
  // `arabicAlpha`, `arabicAbjad`, and `hebrew2` counters of DOCX Editor.
  ['arabic1', formatArabicAlpha],
  ['arabic2', formatArabicAbjad],
  ['hebrew2', formatHebrew2],
  ['thaiAlpha', (n: number) => letters(n, THAI_LETTERS)],
  ['thaiNum', (n: number) => digitsOf(n, THAI_DIGITS)],
  ['hindiAlpha', (n: number) => letters(n, HINDI_VOWELS)],
  ['hindiAlpha1', (n: number) => letters(n, HINDI_CONSONANTS)],
  ['hindiNum', (n: number) => formatHindiNumbers(String(n))],
]);

/** The punctuation of each scheme suffix around a counter `c`. */
const PUNCTUATION: ReadonlyMap<string, (c: string) => string> = new Map([
  ['Period', (c: string) => `${c}.`],
  ['ParenR', (c: string) => `${c})`],
  ['ParenBoth', (c: string) => `(${c})`],
  ['Plain', (c: string) => c],
  // A hyphen-minus right after the counter, with no space (the `rtl-text` reference PDF
  // draws `א-` and `أ-`).
  ['Minus', (c: string) => `${c}-`],
  // A full-width period (U+FF0E).
  ['DbPeriod', (c: string) => `${c}．`],
]);

/**
 * Every `ST_TextAutonumberScheme` value (ECMA-376 Part 1, 20.1.10.61) as its counter family
 * and its punctuation. @internal
 */
export const AUTONUMBER_SCHEMES: ReadonlyMap<string, readonly [string, string]> = new Map(
  (
    [
      ['alphaLcParenBoth', 'alphaLc', 'ParenBoth'],
      ['alphaUcParenBoth', 'alphaUc', 'ParenBoth'],
      ['alphaLcParenR', 'alphaLc', 'ParenR'],
      ['alphaUcParenR', 'alphaUc', 'ParenR'],
      ['alphaLcPeriod', 'alphaLc', 'Period'],
      ['alphaUcPeriod', 'alphaUc', 'Period'],
      ['arabicParenBoth', 'arabic', 'ParenBoth'],
      ['arabicParenR', 'arabic', 'ParenR'],
      ['arabicPeriod', 'arabic', 'Period'],
      ['arabicPlain', 'arabic', 'Plain'],
      ['romanLcParenBoth', 'romanLc', 'ParenBoth'],
      ['romanUcParenBoth', 'romanUc', 'ParenBoth'],
      ['romanLcParenR', 'romanLc', 'ParenR'],
      ['romanUcParenR', 'romanUc', 'ParenR'],
      ['romanLcPeriod', 'romanLc', 'Period'],
      ['romanUcPeriod', 'romanUc', 'Period'],
      ['circleNumDbPlain', 'circleNumDb', 'Plain'],
      ['circleNumWdBlackPlain', 'circleNumWdBlack', 'Plain'],
      ['circleNumWdWhitePlain', 'circleNumWdWhite', 'Plain'],
      ['arabicDbPeriod', 'arabicDb', 'DbPeriod'],
      ['arabicDbPlain', 'arabicDb', 'Plain'],
      ['ea1ChsPeriod', 'ea1Chs', 'Period'],
      ['ea1ChsPlain', 'ea1Chs', 'Plain'],
      ['ea1ChtPeriod', 'ea1Cht', 'Period'],
      ['ea1ChtPlain', 'ea1Cht', 'Plain'],
      ['ea1JpnChsDbPeriod', 'ea1JpnChs', 'DbPeriod'],
      ['ea1JpnKorPlain', 'ea1JpnKor', 'Plain'],
      ['ea1JpnKorPeriod', 'ea1JpnKor', 'Period'],
      ['arabic1Minus', 'arabic1', 'Minus'],
      ['arabic2Minus', 'arabic2', 'Minus'],
      ['hebrew2Minus', 'hebrew2', 'Minus'],
      ['thaiAlphaPeriod', 'thaiAlpha', 'Period'],
      ['thaiAlphaParenR', 'thaiAlpha', 'ParenR'],
      ['thaiAlphaParenBoth', 'thaiAlpha', 'ParenBoth'],
      ['thaiNumPeriod', 'thaiNum', 'Period'],
      ['thaiNumParenR', 'thaiNum', 'ParenR'],
      ['thaiNumParenBoth', 'thaiNum', 'ParenBoth'],
      ['hindiAlphaPeriod', 'hindiAlpha', 'Period'],
      ['hindiNumPeriod', 'hindiNum', 'Period'],
      ['hindiNumParenR', 'hindiNum', 'ParenR'],
      ['hindiAlpha1Period', 'hindiAlpha1', 'Period'],
    ] as const
  ).map(([scheme, counter, punctuation]) => [scheme, [counter, punctuation]] as const)
);

/**
 * Whether a bullet font has a symbol character map (`charset="2"`: Wingdings, Webdings,
 * Symbol): its characters are code points of the font, not Unicode text. @internal
 */
export function isSymbolFont(family: string): boolean {
  return /^(wingdings( [23])?|webdings|symbol)$/i.test(family);
}

/**
 * Wingdings bullets of the Bullets gallery that DOCX Editor's table
 * (`symbol-encoding.ts`, a copy kept equal to its source) does not list, each compared with the
 * glyph the face draws: the Star bullet, `v` (0x76).
 */
const GALLERY_TWINS: ReadonlyMap<string, string> = new Map([['v', '❖']]);

/**
 * A bullet character as a face other than its symbol font draws it: the Unicode twin of a
 * Wingdings or Symbol byte, or the character itself. `isFamilyAvailable` returns true when the
 * symbol font itself draws, and the character is kept. @internal
 */
export function symbolBulletText(
  char: string,
  font: string | null | undefined,
  isFamilyAvailable?: (family: string) => boolean
): string {
  const twin = font?.toLowerCase() === 'wingdings' ? GALLERY_TWINS.get(char) : undefined;
  if (twin !== undefined && isFamilyAvailable?.(font!) !== true) return twin;
  return mapSymbolBulletText(char, font, isFamilyAvailable);
}

/**
 * The last item number that a reference PDF measures for a non-Latin scheme: the `rtl-text`
 * fixture draws Hebrew items 1 to 4 and Arabic items 1 and 2. Larger numbers (the repeated
 * letters and the wrap after 392) come from DOCX Editor's counters, with no measurement here.
 */
const MEASURED_UP_TO: ReadonlyMap<string, number> = new Map([
  ['hebrew2Minus', 4],
  ['arabic1Minus', 2],
]);

/**
 * Whether `autonumberText` draws a scheme as a reference PDF shows it, for every item number
 * up to `upTo`. Latin letters and Arabic and Roman numbers are exact at any number.
 * `hebrew2Minus` is exact for items 1 to 4 and `arabic1Minus` for items 1 and 2 (the
 * `rtl-text` fixture). With no `upTo`, the answer covers every number. The other schemes
 * follow the standard's description of their counter, with no reference measurement.
 * @internal
 */
export function isExactScheme(scheme: string, upTo = Number.POSITIVE_INFINITY): boolean {
  if (/^(alphaLc|alphaUc|arabic|romanLc|romanUc)(ParenBoth|ParenR|Period|Plain)$/.test(scheme))
    return true;
  const measured = MEASURED_UP_TO.get(scheme);
  return measured !== undefined && upTo <= measured;
}

/**
 * The marker text for item `n` (1-based) of a scheme. A value that `ST_TextAutonumberScheme`
 * does not list falls back to `n.`.
 */
export function autonumberText(scheme: string, n: number): string {
  const value = Number.isFinite(n) ? Math.max(1, Math.trunc(n)) : 1;
  const parts = AUTONUMBER_SCHEMES.get(scheme);
  if (!parts) return `${value}.`;
  const [counter, punctuation] = parts;
  return PUNCTUATION.get(punctuation)!(COUNTERS.get(counter)!(value));
}

interface Counter {
  readonly scheme: string;
  readonly startAt: number;
  readonly n: number;
}

/**
 * Autonumber counters for one text body, one per outline level. A numbered paragraph continues
 * its level's sequence when the scheme and `startAt` match the previous number on that level,
 * and restarts at `startAt` otherwise. Any paragraph with text ends the sequences of every
 * deeper level, so sub-items restart under each parent, and a paragraph without a number ends
 * the sequence on its own level. Empty paragraphs show no bullet and change nothing.
 */
export class AutonumberCounters {
  readonly #levels = new Map<number, Counter>();

  /** The number an empty numbered paragraph would show, without advancing any sequence. */
  peek(p: ResolvedParagraph): number | undefined {
    const b = p.bullet;
    if (!b || b.type !== 'autonum') return undefined;
    const prev = this.#levels.get(p.level);
    return prev && prev.scheme === b.scheme && prev.startAt === b.startAt ? prev.n + 1 : b.startAt;
  }

  /** Advance for a paragraph with text. The item number for a numbered paragraph. */
  next(p: ResolvedParagraph): number | undefined {
    for (const level of [...this.#levels.keys()]) if (level > p.level) this.#levels.delete(level);
    const b = p.bullet;
    if (!b || b.type !== 'autonum') {
      this.#levels.delete(p.level);
      return undefined;
    }
    const prev = this.#levels.get(p.level);
    const n =
      prev && prev.scheme === b.scheme && prev.startAt === b.startAt ? prev.n + 1 : b.startAt;
    this.#levels.set(p.level, { scheme: b.scheme, startAt: b.startAt, n });
    return n;
  }
}
