// End-to-end tests: synthetic EOT -> MTX -> CTF -> sfnt, plus hostile-input robustness.

import { describe, expect, test } from 'bun:test';
import { decodeEmbeddedFont, validateSfnt } from '../index.ts';
import {
  ADVANCES,
  buildEot,
  buildMtx,
  buildPlainTtf,
  CVT_VALUES,
  GLYPH1_CODE,
  GLYPH2_CODE,
  GLYPH3_BOX,
  GLYPH3_POINTS,
  HDMX_WIDTHS,
  prng,
  SQUARE,
  SQUARE_ENDPTS,
} from './fixtures.ts';
import { getTable, glyphBytes, parseGlyph, readI16, readU16, tableTags } from './inspect.ts';

const mtx = buildMtx();

function decodeOk(bytes: Uint8Array) {
  const r = decodeEmbeddedFont(bytes);
  if (!r.ok) throw new Error(`decode failed: ${r.reason}: ${r.detail}`);
  return r;
}

describe('decodeEmbeddedFont (MTX)', () => {
  const r = decodeOk(buildEot(mtx, { flags: 0x4, fsType: 8, family: 'Synth', style: 'Bold' }));

  test('returns header info and permission', () => {
    expect(r.info).toMatchObject({
      familyName: 'Synth',
      styleName: 'Bold',
      fsType: 8,
      permission: 'editable',
      compressed: true,
      numGlyphs: 4,
      eotVersion: 0x00020002,
    });
    expect(r.info.warnings).toEqual([]);
  });

  test('produces a structurally valid sfnt', () => {
    const v = validateSfnt(r.sfnt);
    expect(v.errors).toEqual([]);
    expect(v.ok).toBe(true);
    expect(v.numGlyphs).toBe(4);
    expect(tableTags(r.sfnt)).toEqual([
      'cmap',
      'cvt ',
      'glyf',
      'hdmx',
      'head',
      'hhea',
      'hmtx',
      'loca',
      'maxp',
      'name',
      'post',
    ]);
  });

  test('glyph 0 is empty', () => {
    expect(glyphBytes(r.sfnt, 0).length).toBe(0);
  });

  test('simple glyph: contours, computed bbox, points and instructions', () => {
    const g = parseGlyph(glyphBytes(r.sfnt, 1));
    expect(g).not.toBeNull();
    if (!g) return;
    expect(g.numContours).toBe(2);
    expect(g.endPts).toEqual(SQUARE_ENDPTS);
    expect(g.box).toEqual([0, 0, 500, 700]);
    expect(g.points).toEqual(SQUARE);
    // PUSHB[5] 5 7 5 9 5 (Hop3 expanded) followed by the glyph program bytes.
    expect(g.instructions).toEqual([0xb4, 5, 7, 5, 9, 5, ...GLYPH1_CODE]);
  });

  test('composite glyph keeps components and gets push + code instructions', () => {
    const g = parseGlyph(glyphBytes(r.sfnt, 2));
    expect(g).not.toBeNull();
    if (!g) return;
    expect(g.numContours).toBe(-1);
    expect(g.box).toEqual([10, -20, 510, 680]);
    expect(g.components).toEqual([{ flags: 0x103, glyphIndex: 1, args: [10, -20] }]);
    expect(g.instructions).toEqual([0xb9, 0xff, 0xff, 0x01, 0x2c, ...GLYPH2_CODE]);
  });

  test('explicit-bbox glyph decodes spec triplet vectors', () => {
    const g = parseGlyph(glyphBytes(r.sfnt, 3));
    expect(g).not.toBeNull();
    if (!g) return;
    expect(g.box).toEqual(GLYPH3_BOX);
    expect(g.points).toEqual(GLYPH3_POINTS);
    expect(g.instructions).toEqual([]);
  });

  test('cvt and hdmx are reconstructed', () => {
    const cvt = getTable(r.sfnt, 'cvt ');
    expect(cvt).toBeDefined();
    if (!cvt) return;
    expect(CVT_VALUES.map((_, i) => readI16(cvt, i * 2))).toEqual(CVT_VALUES);
    const hdmx = getTable(r.sfnt, 'hdmx');
    expect(hdmx).toBeDefined();
    if (!hdmx) return;
    expect(readU16(hdmx, 2)).toBe(1);
    const size = (readU16(hdmx, 4) << 16) | readU16(hdmx, 6);
    expect(size % 4).toBe(0);
    expect(size).toBeGreaterThanOrEqual(2 + ADVANCES.length);
    expect(hdmx[8]).toBe(12);
    expect([...hdmx.subarray(10, 10 + ADVANCES.length)]).toEqual(HDMX_WIDTHS);
  });

  test('XOR-obfuscated EOT decodes identically', () => {
    const x = decodeOk(buildEot(mtx, { flags: 0x10000004 }));
    expect(x.info.xorEncrypted).toBe(true);
    expect(x.sfnt).toEqual(r.sfnt);
  });

  test('run-length LZCOMP blocks decode identically', () => {
    const x = decodeOk(buildEot(buildMtx({ runLength: true })));
    expect(x.sfnt).toEqual(r.sfnt);
  });

  test('version 0x00010000 and 0x00020001 containers decode', () => {
    for (const version of [0x00010000, 0x00020001]) {
      expect(decodeOk(buildEot(mtx, { version })).sfnt).toEqual(r.sfnt);
    }
  });
});

describe('decodeEmbeddedFont (uncompressed)', () => {
  test('passes a plain TrueType font through', () => {
    const ttf = buildPlainTtf();
    const r = decodeOk(buildEot(ttf, { flags: 0 }));
    expect(r.sfnt).toEqual(ttf);
    expect(r.info.compressed).toBe(false);
    expect(r.info.numGlyphs).toBe(1);
    expect(validateSfnt(r.sfnt).ok).toBe(true);
  });
});

describe('decodeEmbeddedFont (hostile input)', () => {
  const eot = buildEot(mtx);

  test('rejects garbage and empty input', () => {
    expect(decodeEmbeddedFont(new Uint8Array(0)).ok).toBe(false);
    expect(decodeEmbeddedFont(new Uint8Array(1000)).ok).toBe(false);
  });

  test('every truncation is rejected without throwing', () => {
    for (let n = 0; n < eot.length; n += 3) {
      const r = decodeEmbeddedFont(eot.subarray(0, n));
      expect(r.ok).toBe(false);
    }
  });

  test('random corruption never throws and never yields an invalid sfnt', () => {
    const rand = prng(1234);
    let accepted = 0;
    for (let i = 0; i < 400; i++) {
      const bad = eot.slice();
      const flips = 1 + Math.floor(rand() * 6);
      for (let k = 0; k < flips; k++) {
        // Bias toward the MTX payload where the interesting parsers live.
        const pos =
          rand() < 0.85
            ? eot.length - mtx.length + Math.floor(rand() * mtx.length)
            : Math.floor(rand() * eot.length);
        bad[pos] = Math.floor(rand() * 256);
      }
      const r = decodeEmbeddedFont(bad);
      if (r.ok) {
        accepted++;
        expect(validateSfnt(r.sfnt).ok).toBe(true);
      } else {
        expect(typeof r.reason).toBe('string');
      }
    }
    expect(accepted).toBeLessThan(400);
  });

  test('honours the output size limit', () => {
    const r = decodeEmbeddedFont(eot, { maxOutputBytes: 300 });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.reason).toBe('limit-exceeded');
  });

  test('honours the glyph limit', () => {
    const r = decodeEmbeddedFont(eot, { maxGlyphs: 2 });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.reason).toBe('limit-exceeded');
  });

  test('rejects an MTX header with out-of-order offsets', () => {
    const bad = mtx.slice();
    bad[4] = 0;
    bad[5] = 0;
    bad[6] = 5;
    const r = decodeEmbeddedFont(buildEot(bad));
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.reason).toBe('invalid-mtx-header');
  });

  test('rejects an unknown MTX version', () => {
    const bad = mtx.slice();
    bad[0] = 9;
    const r = decodeEmbeddedFont(buildEot(bad));
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.reason).toBe('unsupported-version');
  });
});
