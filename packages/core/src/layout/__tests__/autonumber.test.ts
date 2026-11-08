import { describe, expect, test } from 'bun:test';
import { AUTONUMBER_SCHEMES, autonumberText, isExactScheme } from '../text/autonumber.ts';

const RLM = '‏';
const ZWNJ = '‌';

// The values of `ST_TextAutonumberScheme` in `dml-main.xsd` (ECMA-376 Part 4, transitional).
const SCHEMA_VALUES = [
  'alphaLcParenBoth',
  'alphaUcParenBoth',
  'alphaLcParenR',
  'alphaUcParenR',
  'alphaLcPeriod',
  'alphaUcPeriod',
  'arabicParenBoth',
  'arabicParenR',
  'arabicPeriod',
  'arabicPlain',
  'romanLcParenBoth',
  'romanUcParenBoth',
  'romanLcParenR',
  'romanUcParenR',
  'romanLcPeriod',
  'romanUcPeriod',
  'circleNumDbPlain',
  'circleNumWdBlackPlain',
  'circleNumWdWhitePlain',
  'arabicDbPeriod',
  'arabicDbPlain',
  'ea1ChsPeriod',
  'ea1ChsPlain',
  'ea1ChtPeriod',
  'ea1ChtPlain',
  'ea1JpnChsDbPeriod',
  'ea1JpnKorPlain',
  'ea1JpnKorPeriod',
  'arabic1Minus',
  'arabic2Minus',
  'hebrew2Minus',
  'thaiAlphaPeriod',
  'thaiAlphaParenR',
  'thaiAlphaParenBoth',
  'thaiNumPeriod',
  'thaiNumParenR',
  'thaiNumParenBoth',
  'hindiAlphaPeriod',
  'hindiNumPeriod',
  'hindiNumParenR',
  'hindiAlpha1Period',
];

/** Each scheme's marker at 1, 10, 15, 100, and its edge cases. */
const CASES: Record<string, [number, string][]> = {
  alphaLcParenBoth: [
    [1, '(a)'],
    [10, '(j)'],
    [15, '(o)'],
    [26, '(z)'],
    [27, '(aa)'],
    [100, '(cv)'],
  ],
  alphaUcParenBoth: [
    [1, '(A)'],
    [10, '(J)'],
    [15, '(O)'],
    [27, '(AA)'],
    [100, '(CV)'],
  ],
  alphaLcParenR: [
    [1, 'a)'],
    [10, 'j)'],
    [15, 'o)'],
    [52, 'az)'],
    [100, 'cv)'],
  ],
  alphaUcParenR: [
    [1, 'A)'],
    [10, 'J)'],
    [15, 'O)'],
    [53, 'BA)'],
    [100, 'CV)'],
  ],
  alphaLcPeriod: [
    [1, 'a.'],
    [10, 'j.'],
    [15, 'o.'],
    [100, 'cv.'],
    [702, 'zz.'],
    [703, 'aaa.'],
  ],
  alphaUcPeriod: [
    [1, 'A.'],
    [10, 'J.'],
    [15, 'O.'],
    [26, 'Z.'],
    [100, 'CV.'],
  ],
  arabicParenBoth: [
    [1, '(1)'],
    [10, '(10)'],
    [15, '(15)'],
    [100, '(100)'],
    [32767, '(32767)'],
  ],
  arabicParenR: [
    [1, '1)'],
    [10, '10)'],
    [15, '15)'],
    [100, '100)'],
  ],
  arabicPeriod: [
    [1, '1.'],
    [10, '10.'],
    [15, '15.'],
    [100, '100.'],
    [32767, '32767.'],
  ],
  arabicPlain: [
    [1, '1'],
    [10, '10'],
    [15, '15'],
    [100, '100'],
  ],
  romanLcParenBoth: [
    [1, '(i)'],
    [4, '(iv)'],
    [10, '(x)'],
    [15, '(xv)'],
    [100, '(c)'],
  ],
  romanUcParenBoth: [
    [1, '(I)'],
    [9, '(IX)'],
    [10, '(X)'],
    [15, '(XV)'],
    [100, '(C)'],
  ],
  romanLcParenR: [
    [1, 'i)'],
    [10, 'x)'],
    [15, 'xv)'],
    [49, 'xlix)'],
    [100, 'c)'],
  ],
  romanUcParenR: [
    [1, 'I)'],
    [10, 'X)'],
    [15, 'XV)'],
    [100, 'C)'],
    [1994, 'MCMXCIV)'],
  ],
  romanLcPeriod: [
    [1, 'i.'],
    [10, 'x.'],
    [15, 'xv.'],
    [100, 'c.'],
    [3999, 'mmmcmxcix.'],
  ],
  // Roman numbers stop at 3999.
  romanUcPeriod: [
    [1, 'I.'],
    [10, 'X.'],
    [15, 'XV.'],
    [100, 'C.'],
    [4000, 'MMMCMXCIX.'],
  ],
  circleNumDbPlain: [
    [1, '①'],
    [9, '⑨'],
    [10, '⑩'],
    [11, '11'],
    [15, '15'],
    [100, '100'],
  ],
  circleNumWdBlackPlain: [
    [1, '❶'],
    [10, '❿'],
    [11, '11'],
    [15, '15'],
    [100, '100'],
  ],
  circleNumWdWhitePlain: [
    [1, '①'],
    [10, '⑩'],
    [11, '11'],
    [15, '15'],
    [100, '100'],
  ],
  arabicDbPeriod: [
    [1, '１．'],
    [10, '１０．'],
    [15, '１５．'],
    [100, '１００．'],
  ],
  arabicDbPlain: [
    [1, '１'],
    [10, '１０'],
    [15, '１５'],
    [100, '１００'],
  ],
  ea1ChsPeriod: [
    [1, '一.'],
    [10, '十.'],
    [15, '十五.'],
    [20, '二十.'],
    [99, '九十九.'],
    [100, '一〇〇.'],
  ],
  ea1ChsPlain: [
    [1, '一'],
    [10, '十'],
    [15, '十五'],
    [21, '二十一'],
    [100, '一〇〇'],
    [205, '二〇五'],
  ],
  // Type A stops at 19.
  ea1ChtPeriod: [
    [1, '一.'],
    [10, '十.'],
    [15, '十五.'],
    [19, '十九.'],
    [20, '二〇.'],
    [100, '一〇〇.'],
  ],
  ea1ChtPlain: [
    [1, '一'],
    [10, '十'],
    [15, '十五'],
    [30, '三〇'],
    [100, '一〇〇'],
  ],
  ea1JpnChsDbPeriod: [
    [1, '一．'],
    [10, '十．'],
    [15, '十五．'],
    [99, '九十九．'],
    [100, '一〇〇．'],
  ],
  ea1JpnKorPlain: [
    [1, '一'],
    [10, '一〇'],
    [15, '一五'],
    [100, '一〇〇'],
  ],
  ea1JpnKorPeriod: [
    [1, '一.'],
    [10, '一〇.'],
    [15, '一五.'],
    [100, '一〇〇.'],
  ],
  arabic1Minus: [
    [1, `أ${ZWNJ}-`],
    [2, `ب${ZWNJ}-`],
    [10, `ر${ZWNJ}-`],
    [15, `ض${ZWNJ}-`],
    [29, `أ${ZWNJ}أ${ZWNJ}-`],
    [100, `ط${ZWNJ}`.repeat(4) + '-'],
  ],
  arabic2Minus: [
    [1, `${ZWNJ}أ-`],
    [3, `${ZWNJ}ج-`],
    [10, `${ZWNJ}ي-`],
    [15, `${ZWNJ}س-`],
    [100, `${ZWNJ}ع`.repeat(4) + '-'],
  ],
  hebrew2Minus: [
    [1, `${RLM}א-`],
    [4, `${RLM}ד-`],
    [10, `${RLM}י-`],
    [15, `${RLM}ס-`],
    [23, `${RLM}תא-`],
    [100, `${RLM}תתתתל-`],
    // The letter sequences restart after 392.
    [393, `${RLM}א-`],
  ],
  thaiAlphaPeriod: [
    [1, 'ก.'],
    [10, 'ญ.'],
    [15, 'ฒ.'],
    [41, 'ฮ.'],
    [42, 'กก.'],
    [100, 'ขต.'],
  ],
  thaiAlphaParenR: [
    [1, 'ก)'],
    [10, 'ญ)'],
    [15, 'ฒ)'],
    [100, 'ขต)'],
  ],
  thaiAlphaParenBoth: [
    [1, '(ก)'],
    [10, '(ญ)'],
    [15, '(ฒ)'],
    [100, '(ขต)'],
  ],
  thaiNumPeriod: [
    [1, '๑.'],
    [10, '๑๐.'],
    [15, '๑๕.'],
    [100, '๑๐๐.'],
  ],
  thaiNumParenR: [
    [1, '๑)'],
    [10, '๑๐)'],
    [15, '๑๕)'],
    [100, '๑๐๐)'],
  ],
  thaiNumParenBoth: [
    [1, '(๑)'],
    [10, '(๑๐)'],
    [15, '(๑๕)'],
    [100, '(๑๐๐)'],
  ],
  hindiAlphaPeriod: [
    [1, 'अ.'],
    [10, 'ओ.'],
    [11, 'औ.'],
    [12, 'अअ.'],
    [15, 'अई.'],
    [100, 'ऐअ.'],
  ],
  hindiNumPeriod: [
    [1, '१.'],
    [10, '१०.'],
    [15, '१५.'],
    [100, '१००.'],
  ],
  hindiNumParenR: [
    [1, '१)'],
    [10, '१०)'],
    [15, '१५)'],
    [100, '१००)'],
  ],
  hindiAlpha1Period: [
    [1, 'क.'],
    [10, 'ञ.'],
    [15, 'ण.'],
    [33, 'ह.'],
    [34, 'कक.'],
    [100, 'गक.'],
  ],
};

describe('autonumberText', () => {
  test('the table lists every ST_TextAutonumberScheme value, and nothing else', () => {
    expect([...AUTONUMBER_SCHEMES.keys()].sort()).toEqual([...SCHEMA_VALUES].sort());
    expect(Object.keys(CASES).sort()).toEqual([...SCHEMA_VALUES].sort());
  });

  for (const scheme of SCHEMA_VALUES)
    test(`${scheme} draws its counter and punctuation`, () => {
      for (const [n, text] of CASES[scheme]!)
        expect([n, autonumberText(scheme, n)]).toEqual([n, text]);
    });

  test('a value that the schema does not list falls back to n.', () => {
    expect(autonumberText('arabicColon', 3)).toBe('3.');
    expect(autonumberText('hebrew1Minus', 3)).toBe('3.');
    expect(autonumberText('ArabicPeriod', 3)).toBe('3.');
    expect(autonumberText('', 7)).toBe('7.');
    expect(autonumberText('__proto__', 2)).toBe('2.');
    expect(autonumberText('constructor', 2)).toBe('2.');
    expect(autonumberText('toString', 2)).toBe('2.');
  });

  test('a number outside the counter range draws as 1', () => {
    expect(autonumberText('arabicPeriod', 0)).toBe('1.');
    expect(autonumberText('alphaLcParenR', -4)).toBe('a)');
    expect(autonumberText('hebrew2Minus', Number.NaN)).toBe(`${RLM}א-`);
    expect(autonumberText('romanUcPeriod', Number.POSITIVE_INFINITY)).toBe('I.');
    expect(autonumberText('ea1ChsPlain', 2.7)).toBe('二');
    expect(autonumberText('unknownScheme', Number.NaN)).toBe('1.');
  });

  test('large counters stay short', () => {
    for (const scheme of SCHEMA_VALUES)
      expect(autonumberText(scheme, 1_000_000).length).toBeLessThanOrEqual(32);
  });
});

describe('isExactScheme', () => {
  test('lists the schemes that a reference PDF measures', () => {
    expect(isExactScheme('arabicPeriod')).toBe(true);
    expect(isExactScheme('romanUcParenR')).toBe(true);
    expect(isExactScheme('arabicPeriod', 32767)).toBe(true);
    // The rtl-text fixture measures Hebrew items 1 to 4 and Arabic items 1 and 2.
    expect(isExactScheme('hebrew2Minus', 4)).toBe(true);
    expect(isExactScheme('hebrew2Minus', 5)).toBe(false);
    expect(isExactScheme('hebrew2Minus')).toBe(false);
    expect(isExactScheme('arabic1Minus', 2)).toBe(true);
    expect(isExactScheme('arabic1Minus', 3)).toBe(false);
    expect(isExactScheme('arabic1Minus')).toBe(false);
    expect(isExactScheme('arabic2Minus', 1)).toBe(false);
    expect(isExactScheme('thaiAlphaPeriod')).toBe(false);
    expect(isExactScheme('ea1ChsPlain')).toBe(false);
    expect(isExactScheme('unknownScheme')).toBe(false);
  });
});
