// Adversarial tests for the OPC/ZIP/relationship normalization profile. Covers the
// trust-boundary rejections, the "two entries normalize to one name" scenario, owner-relative
// resolution without root escape, and the external-URI profile.
//
// Adapted from docx-editor/packages/core/src/store/__tests__/opc-names.test.ts at
// 26095c632 (EigenPal, Apache-2.0), with WordprocessingML part names replaced by PresentationML ones. The `zipItemName` block is
// the regression test for audit finding K1.

import { describe, expect, test } from 'bun:test';
import {
  normalizePartName,
  detectDuplicateNames,
  resolveInternalTarget,
  validateExternalTarget,
  partNameKey,
  zipItemName,
  isCanonicalPartName,
  type NameRejection,
} from '../opc-names.ts';

function reason(r: ReturnType<typeof normalizePartName>): NameRejection | 'ok' {
  return r.ok ? 'ok' : r.reason;
}

describe('normalizePartName', () => {
  test('canonicalizes with and without leading slash', () => {
    expect(normalizePartName('ppt/presentation.xml')).toEqual({
      ok: true,
      partName: '/ppt/presentation.xml',
    });
    expect(normalizePartName('/ppt/presentation.xml')).toEqual({
      ok: true,
      partName: '/ppt/presentation.xml',
    });
  });
  test('rejects the traversal/encoding/drive attack surface', () => {
    expect(reason(normalizePartName('ppt\\slides\\slide1.xml'))).toBe('backslash');
    expect(reason(normalizePartName('C:/ppt/x.xml'))).toBe('drive-or-unc');
    expect(reason(normalizePartName('//server/share/x'))).toBe('drive-or-unc');
    expect(reason(normalizePartName('ppt%2f..%2fsecret'))).toBe('encoded-separator');
    expect(reason(normalizePartName('%2e%2e/secret'))).toBe('encoded-dot');
    expect(reason(normalizePartName('ppt/../secret.xml'))).toBe('dot-segment');
    expect(reason(normalizePartName('ppt/./x.xml'))).toBe('dot-segment');
    expect(reason(normalizePartName('ppt//x.xml'))).toBe('empty-segment');
    expect(reason(normalizePartName('ppt/foo./x.xml'))).toBe('segment-trailing-dot');
    expect(reason(normalizePartName('%zz'))).toBe('bad-encoding');
    expect(reason(normalizePartName(''))).toBe('empty');
  });
  test('rejects a control character (including one hidden by percent-encoding)', () => {
    expect(reason(normalizePartName('ppt/\x00.xml'))).toBe('control-char');
    expect(reason(normalizePartName('ppt/%00.xml'))).toBe('control-char');
  });
  test('decodes safe percent-encoding canonically', () => {
    // %41 == "A" is a safe (non-separator, non-dot) encoding.
    expect(normalizePartName('ppt/%41.xml')).toEqual({ ok: true, partName: '/ppt/A.xml' });
  });
});

describe('duplicate detection before inflation', () => {
  test('two names colliding after case-folded normalization are flagged', () => {
    const { duplicates } = detectDuplicateNames(['ppt/x.xml', 'PPT/X.xml']);
    expect(duplicates).toContain(partNameKey('/ppt/x.xml'));
  });
  test('rejected names are reported, not silently dropped', () => {
    const { rejected } = detectDuplicateNames(['ok/a.xml', 'bad\\b.xml']);
    expect(rejected.map((r) => r.reason)).toContain('backslash');
  });
});

describe('resolveInternalTarget (owner-relative, no root escape)', () => {
  const owner = '/ppt/slides/slide1.xml';
  test('resolves relative media and parent references within root', () => {
    expect(resolveInternalTarget(owner, '../media/image1.png')).toEqual({
      ok: true,
      partName: '/ppt/media/image1.png',
    });
    expect(resolveInternalTarget(owner, '../../customXml/item1.xml')).toEqual({
      ok: true,
      partName: '/customXml/item1.xml',
    });
    expect(resolveInternalTarget(owner, './slide2.xml')).toEqual({
      ok: true,
      partName: '/ppt/slides/slide2.xml',
    });
  });
  test('package-absolute target resolves from root', () => {
    expect(resolveInternalTarget(owner, '/ppt/presentation.xml')).toEqual({
      ok: true,
      partName: '/ppt/presentation.xml',
    });
  });
  test('escaping the package root is rejected', () => {
    expect(reason(resolveInternalTarget(owner, '../../../../etc/passwd'))).toBe('traversal-escape');
    expect(reason(resolveInternalTarget(owner, '..%2f..%2fx'))).toBe('encoded-separator');
  });
});

describe('validateExternalTarget (absolute-URI profile, no fetch)', () => {
  test('accepts absolute http/mailto and preserves raw', () => {
    expect(validateExternalTarget('https://example.com/a')).toEqual({
      ok: true,
      partName: 'https://example.com/a',
    });
    expect(validateExternalTarget('mailto:a@b.com').ok).toBe(true);
  });
  test('rejects unsafe schemes and relative/control forms', () => {
    expect(reason(validateExternalTarget('javascript:alert(1)'))).toBe('unsafe-scheme');
    expect(reason(validateExternalTarget('data:text/html,<x>'))).toBe('unsafe-scheme');
    expect(reason(validateExternalTarget('file:///etc/passwd'))).toBe('unsafe-scheme');
    expect(reason(validateExternalTarget('relative/path'))).toBe('not-absolute-uri');
    expect(reason(validateExternalTarget('\x01http://x'))).toBe('control-char');
  });
});

describe('K1: zipItemName encodes, and never decodes a second time', () => {
  test('ordinary names are written unchanged', () => {
    expect(zipItemName('/ppt/slides/slide1.xml')).toEqual({
      ok: true,
      itemName: 'ppt/slides/slide1.xml',
    });
    expect(zipItemName('/[Content_Types].xml')).toEqual({
      ok: true,
      itemName: '[Content_Types].xml',
    });
  });
  test('a decoded percent sign or space is encoded back, and reads back the same', () => {
    for (const [partName, itemName] of [
      ['/a%41.xml', 'a%2541.xml'],
      ['/b%zz.bin', 'b%25zz.bin'],
      ['/ppt/media/image 1.png', 'ppt/media/image%201.png'],
      ['/ppt/media/a#b.png', 'ppt/media/a%23b.png'],
    ]) {
      expect(zipItemName(partName)).toEqual({ ok: true, itemName });
      expect(normalizePartName(itemName)).toEqual({ ok: true, partName });
    }
  });
  test('a name that is not canonical is refused, not renamed', () => {
    expect(zipItemName('ppt/x.xml')).toEqual({ ok: false, reason: 'not-canonical' });
    expect(zipItemName('/ppt/../x.xml')).toEqual({ ok: false, reason: 'dot-segment' });
    expect(isCanonicalPartName('/ppt/media/image1.png')).toBe(true);
    expect(isCanonicalPartName('/ppt//x.xml')).toBe(false);
  });
});
