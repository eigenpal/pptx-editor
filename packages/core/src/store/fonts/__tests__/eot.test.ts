// Tests for EOT header parsing (W3C EOT Submission section 3) and fsType permissions.

import { describe, expect, test } from 'bun:test';
import { parseEot } from '../eot.ts';
import { fontEmbeddingPermission } from '../index.ts';
import { buildEot, buildPlainTtf, prng } from './fixtures.ts';

const payload = Uint8Array.from([1, 2, 3, 4, 5, 6, 7, 8, 9]);

describe('parseEot', () => {
  for (const version of [0x00010000, 0x00020001, 0x00020002]) {
    test(`parses version 0x${version.toString(16)}`, () => {
      const eot = buildEot(payload, {
        version,
        flags: 0x5,
        family: 'Fam',
        style: 'Bold',
        weight: 700,
        italic: true,
        fsType: 8,
        charset: 0,
        rootString: version === 0x00010000 ? '' : 'https://example.test/',
      });
      const r = parseEot(eot);
      expect(r.ok).toBe(true);
      if (!r.ok) return;
      expect(r.header.version).toBe(version);
      expect(r.header.familyName).toBe('Fam');
      expect(r.header.styleName).toBe('Bold');
      expect(r.header.fullName).toBe('Fam Bold');
      expect(r.header.versionName).toBe('Version 1.0');
      expect(r.header.weight).toBe(700);
      expect(r.header.italic).toBe(true);
      expect(r.header.fsType).toBe(8);
      expect(r.header.charset).toBe(0);
      expect(r.header.flags).toBe(5);
      expect(r.header.compressed).toBe(true);
      expect(r.header.subset).toBe(true);
      expect(r.header.xorEncrypted).toBe(false);
      expect([...r.fontData]).toEqual([...payload]);
      if (version !== 0x00010000) expect(r.header.rootString).toBe('https://example.test/');
      if (version === 0x00020002) expect(r.header.rootStringChecksumValid).toBe(true);
      else expect(r.header.rootStringChecksumValid).toBeNull();
    });
  }

  test('removes XOR obfuscation (TTEMBED_XORENCRYPTDATA)', () => {
    const eot = buildEot(payload, { flags: 0x10000004 });
    const r = parseEot(eot);
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.header.xorEncrypted).toBe(true);
    expect([...r.fontData]).toEqual([...payload]);
  });

  test('rejects bad magic', () => {
    const eot = buildEot(payload);
    eot[34] = 0;
    const r = parseEot(eot);
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.reason).toBe('bad-magic');
  });

  test('rejects unknown version', () => {
    const eot = buildEot(payload);
    eot[8] = 7;
    const r = parseEot(eot);
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.reason).toBe('unsupported-version');
  });

  test('rejects FontDataSize larger than the structure', () => {
    const eot = buildEot(payload);
    eot[4] = 0xff;
    const r = parseEot(eot);
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.reason).toBe('size-mismatch');
  });

  test('rejects every truncation without throwing', () => {
    const eot = buildEot(buildPlainTtf());
    for (let n = 0; n < eot.length; n++) {
      const r = parseEot(eot.subarray(0, n));
      expect(r.ok).toBe(false);
    }
  });

  test('rejects oversized name lengths', () => {
    const eot = buildEot(payload);
    eot[82] = 0xff;
    eot[83] = 0xff;
    expect(parseEot(eot).ok).toBe(false);
  });

  test('random garbage never throws', () => {
    const rand = prng(42);
    for (let i = 0; i < 500; i++) {
      const n = Math.floor(rand() * 300);
      const buf = new Uint8Array(n);
      for (let k = 0; k < n; k++) buf[k] = Math.floor(rand() * 256);
      if (n > 36 && rand() < 0.5) {
        buf[34] = 0x4c;
        buf[35] = 0x50;
        buf[8] = 2;
        buf[9] = 0;
        buf[10] = 2;
        buf[11] = 0;
      }
      expect(() => parseEot(buf)).not.toThrow();
    }
  });
});

describe('fontEmbeddingPermission', () => {
  test('maps OS/2 fsType bits', () => {
    expect(fontEmbeddingPermission(0)).toBe('installable');
    expect(fontEmbeddingPermission(0x0002)).toBe('restricted');
    expect(fontEmbeddingPermission(0x0004)).toBe('preview-print');
    expect(fontEmbeddingPermission(0x0008)).toBe('editable');
    expect(fontEmbeddingPermission(0x0100)).toBe('installable');
    expect(fontEmbeddingPermission(0x0108)).toBe('editable');
  });

  test('least restrictive bit wins for legacy multi-bit values', () => {
    expect(fontEmbeddingPermission(0x0006)).toBe('preview-print');
    expect(fontEmbeddingPermission(0x000e)).toBe('editable');
  });

  test('bitmap-only embedding is restricted for an outline renderer', () => {
    expect(fontEmbeddingPermission(0x0200)).toBe('restricted');
    expect(fontEmbeddingPermission(0x0208)).toBe('restricted');
  });
});
