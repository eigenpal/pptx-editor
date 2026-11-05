// Tests for sfnt assembly and the structural validator (OpenType table directory rules).

import { describe, expect, test } from 'bun:test';
import { buildSfnt, calcChecksum, readSfntName, validateSfnt, writeU32 } from '../sfnt.ts';
import { FontFormatError } from '../types.ts';
import { buildPlainTtf } from './fixtures.ts';
import { readU32 } from './inspect.ts';

describe('calcChecksum', () => {
  test('pads the final partial word with zeros', () => {
    const d = Uint8Array.from([0, 0, 0, 1, 0, 0, 0, 2, 0x80]);
    expect(calcChecksum(d, 0, 9)).toBe(3 + 0x80000000);
  });
});

describe('buildSfnt', () => {
  test('sorts tables, aligns them and fixes checksums', () => {
    const font = buildPlainTtf();
    const v = validateSfnt(font);
    expect(v.errors).toEqual([]);
    expect(v.tables).toEqual([...v.tables].sort());
    expect(calcChecksum(font, 0, font.length)).toBe(0xb1b0afba);
  });

  test('rejects duplicate tags and size overflow', () => {
    const t = { tag: 'abcd', data: new Uint8Array(3) };
    expect(() => buildSfnt(0x00010000, [t, t], 1000)).toThrow(FontFormatError);
    expect(() => buildSfnt(0x00010000, [t], 10)).toThrow(FontFormatError);
  });
});

describe('validateSfnt', () => {
  const good = buildPlainTtf();

  const recordOf = (font: Uint8Array, tag: string): number => {
    for (let i = 0; i < ((font[4] << 8) | font[5]); i++) {
      const r = 12 + i * 16;
      if (String.fromCharCode(...font.subarray(r, r + 4)) === tag) return r;
    }
    throw new Error(`no ${tag}`);
  };

  test('accepts the reference font', () => {
    expect(validateSfnt(good).ok).toBe(true);
  });

  test('detects a table checksum mismatch', () => {
    const bad = good.slice();
    const rec = recordOf(bad, 'name');
    bad[readU32(bad, rec + 8)] ^= 1;
    const v = validateSfnt(bad);
    expect(v.ok).toBe(false);
    expect(v.errors.join()).toContain("'name' checksum mismatch");
  });

  test('detects a wrong head.checkSumAdjustment', () => {
    const bad = good.slice();
    const head = readU32(bad, recordOf(bad, 'head') + 8);
    writeU32(bad, head + 8, 0);
    expect(validateSfnt(bad).errors.join()).toContain('checkSumAdjustment');
  });

  test('detects unsorted directory entries', () => {
    const bad = good.slice();
    const a = 12;
    const b = 28;
    const tmp = bad.slice(a, a + 16);
    bad.copyWithin(a, b, b + 16);
    bad.set(tmp, b);
    expect(validateSfnt(bad).errors.join()).toContain('out of order');
  });

  test('detects misaligned and out-of-bounds offsets', () => {
    const bad = good.slice();
    const rec = recordOf(bad, 'post');
    writeU32(bad, rec + 8, readU32(bad, rec + 8) + 1);
    expect(validateSfnt(bad).errors.join()).toContain('not 4-byte aligned');
    writeU32(bad, rec + 8, bad.length);
    expect(validateSfnt(bad).errors.join()).toContain('out of bounds');
  });

  test('detects maxp.numGlyphs inconsistent with loca', () => {
    const bad = good.slice();
    const maxp = readU32(bad, recordOf(bad, 'maxp') + 8);
    bad[maxp + 5] = 5;
    expect(validateSfnt(bad).errors.join()).toContain('loca length');
  });

  test('handles tiny and garbage input', () => {
    expect(validateSfnt(new Uint8Array(4)).ok).toBe(false);
    expect(validateSfnt(Uint8Array.from([0, 1, 0, 0, 0xff, 0xff, 0, 0, 0, 0, 0, 0])).ok).toBe(
      false
    );
  });
});

describe('readSfntName', () => {
  test('reads the Windows English family name', () => {
    expect(readSfntName(buildPlainTtf(), 1)).toBe('Synth');
    expect(readSfntName(buildPlainTtf(), 2)).toBe('');
    expect(readSfntName(new Uint8Array(3), 1)).toBe('');
  });
});
