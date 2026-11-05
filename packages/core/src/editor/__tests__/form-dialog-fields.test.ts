// The form models of the Columns dialog and the color fields, which the React and Vue dialogs
// share: the seed, the range checks, and the one command OK sends.

import { describe, expect, test } from 'bun:test';
import type { FontFormatRead } from '../../contracts/index.ts';
import {
  colorChange,
  colorPicks,
  columnsCommandOf,
  columnsFieldError,
  bulletFontChoices,
  bulletFontMixed,
  bulletsFieldError,
  editBulletGlyph,
  fontCommandOf,
  pickBulletStyle,
  seedBulletFields,
  seedColumnsFields,
  seedFontFields,
} from '../index.ts';

describe('color fields', () => {
  test('offer "no change" only when the field opens with it, and send only a change', () => {
    expect(colorPicks('auto')).not.toContain('keep');
    expect(colorPicks('keep').slice(0, 3)).toEqual(['auto', 'keep', 'tx1']);
    expect(colorChange('keep', 'keep')).toBeUndefined();
    expect(colorChange('auto', 'auto')).toBeUndefined();
    expect(colorChange('keep', 'auto')).toBeNull();
    expect(colorChange('auto', 'accent1')).toEqual({ theme: 'accent1' });
    // Colors that differ open at "Mixed", which sends no change.
    expect(colorPicks('mixed').slice(0, 3)).toEqual(['auto', 'mixed', 'tx1']);
    expect(colorChange('mixed', 'mixed')).toBeUndefined();
    expect(colorChange('mixed', 'auto')).toBeNull();
  });

  test('runs whose underline colors differ open the field at "Mixed"', () => {
    const one = seedFontFields(null);
    const read = {
      fontFamily: 'Arial',
      fontSizePt: 18,
      bold: false,
      italic: false,
      underline: 'single',
      strike: 'none',
      baseline: 0,
      caps: 'none',
      characterSpacingPt: 0,
      kerningPt: null,
      underlineColor: null,
      disagrees: {
        fontFamily: false,
        fontSizePt: false,
        bold: false,
        italic: false,
        underline: false,
        strike: false,
        baseline: false,
        caps: false,
        characterSpacingPt: false,
        kerningPt: false,
        underlineColor: true,
      },
    } satisfies FontFormatRead;
    const seed = seedFontFields(read);
    expect(seed.underlineColor).toBe('mixed');
    expect(fontCommandOf(seed, seed)).toBeNull();
    expect(one.underlineColor).toBe('mixed');
  });

  test('the Font dialog sends the font and underline colors that moved', () => {
    const seed = seedFontFields(null);
    expect([seed.color, seed.underlineColor]).toEqual(['keep', 'mixed']);
    expect(fontCommandOf(seed, { ...seed, color: 'accent2', underlineColor: 'auto' })).toEqual({
      color: { theme: 'accent2' },
      underlineColor: null,
    });
    expect(fontCommandOf(seed, seed)).toBeNull();
    // A dimmed underline color (the underline set to none) is not sent.
    expect(fontCommandOf(seed, { ...seed, underline: 'none', underlineColor: 'accent2' })).toEqual({
      underline: 'none',
    });
  });
});

describe('Bullets and Numbering dialog fields', () => {
  test('typing back the character the dialog opened with brings back its style', () => {
    const seed = seedBulletFields({
      kind: 'bullet',
      char: 'ü',
      font: 'Wingdings',
      sizePercent: 100,
    });
    expect(seed.bullet).toBe('checkmark');
    const typed = editBulletGlyph(seed, 'char', 'x', seed);
    expect(typed.bullet).toBe('current');
    expect(editBulletGlyph(typed, 'char', 'ü', seed).bullet).toBe('checkmark');
    const star = pickBulletStyle(seed, 'star');
    expect([star.char, star.font]).toEqual(['v', 'Wingdings']);
  });

  test('over bullets that differ, the Font field is mixed until a pick, and every pick asks for a character', () => {
    const seed = seedBulletFields(undefined);
    expect(bulletFontMixed(seed, seed)).toBe(true);
    for (const font of ['', 'Symbol']) {
      const picked = editBulletGlyph(seed, 'font', font, seed);
      expect(bulletFontMixed(seed, picked)).toBe(false);
      expect(bulletsFieldError(seed, picked)?.field).toBe('char');
    }
    const typed = editBulletGlyph(editBulletGlyph(seed, 'font', '', seed), 'char', '→', seed);
    expect(bulletsFieldError(seed, typed)).toBeNull();
  });

  test('the Font field always offers the gallery bullet fonts and Symbol', () => {
    const fonts = bulletFontChoices(['Arial', 'Calibri'], ['Calibri', 'Aptos']);
    expect(fonts.slice(0, 2)).toEqual(['Arial', 'Calibri']);
    for (const f of ['Wingdings', 'Courier New', 'Symbol', 'Aptos']) expect(fonts).toContain(f);
    expect(new Set(fonts).size).toBe(fonts.length);
    expect(bulletFontChoices(['', ''], [])).not.toContain('');
  });

  test('on paragraphs with no list, typing the first bullet back never selects None', () => {
    const seed = seedBulletFields(null);
    expect([seed.bullet, seed.char, seed.font]).toEqual(['none', '•', 'Arial']);
    const round = pickBulletStyle(seed, 'filledRound');
    const typed = editBulletGlyph(round, 'char', 'x', seed);
    expect(typed.bullet).toBe('current');
    const back = editBulletGlyph(typed, 'char', '•', seed);
    expect(back.bullet).not.toBe('none');
    expect(back.bullet).toBe(round.bullet);
  });
});

describe('Columns dialog fields', () => {
  test('seed in inches, check the engine limits, and send the spacing only when it moved', () => {
    const seed = seedColumnsFields({ count: 2, gapEmu: 457_200 });
    expect(seed).toEqual({ count: '2', spacing: '0.5' });
    expect(columnsCommandOf(seed, seed)).toBeNull();
    expect(columnsCommandOf(seed, { ...seed, count: '3' })).toEqual({
      type: 'setColumns',
      count: 3,
    });
    expect(columnsCommandOf(seed, { ...seed, spacing: '1' })).toEqual({
      type: 'setColumns',
      count: 2,
      gapEmu: 914_400,
    });
    for (const count of ['0', '17', '2.5', '']) {
      expect(columnsFieldError(seed, { ...seed, count })?.field).toBe('count');
    }
    expect(columnsFieldError(seed, { ...seed, spacing: '-1' })?.field).toBe('spacing');
    expect(columnsFieldError(seed, { ...seed, spacing: '57' })?.params).toEqual({
      min: 0,
      max: 56,
    });
    expect(columnsFieldError(seed, { count: '16', spacing: '0' })).toBeNull();
  });

  test('shapes that differ open blank; a spacing change needs a count, a count change keeps each gap', () => {
    const seed = seedColumnsFields(undefined);
    expect(seed).toEqual({ count: '', spacing: '' });
    expect(columnsCommandOf(seed, seed)).toBeNull();
    expect(columnsFieldError(seed, { ...seed, spacing: '0.25' })?.field).toBe('count');
    expect(columnsFieldError(seed, { count: '2', spacing: '0.25' })).toBeNull();
    expect(columnsCommandOf(seed, { ...seed, count: '2' })).toEqual({
      type: 'setColumns',
      count: 2,
    });
  });
});
