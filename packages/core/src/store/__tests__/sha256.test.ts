// SHA-256 tests.
//
// The first block is ported from docx-editor/packages/core/src/store/__tests__/sha256.test.ts
// (EigenPal, Apache-2.0). The second is the regression test for audit finding P5
//: part fingerprints use the native hash and give the same digests.

import { describe, expect, test } from 'bun:test';
import { sha256FontBytes, sha256FontBytesPure, sha256IsNative } from '@eigenpal/ooxml-core';
import { bytesFingerprint, canonicalFingerprint, canonicalForm } from '../package/serialize.ts';
import { parseXmlPart } from '../package/tree.ts';

describe('sha256FontBytes', () => {
  test('matches the FIPS 180-4 vectors', () => {
    const text = (s: string) => new TextEncoder().encode(s);
    expect(sha256FontBytesPure(text(''))).toBe(
      'sha256:e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855'
    );
    expect(sha256FontBytesPure(text('abc'))).toBe(
      'sha256:ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad'
    );
    expect(sha256FontBytes(text('abc'))).toBe(sha256FontBytesPure(text('abc')));
  });

  test('the native path, when the host has one, agrees with the pure path byte for byte', () => {
    // Bun always has a native hasher, so this exercises the fast path in the test suite.
    expect(sha256IsNative).toBe(true);
    for (const length of [0, 1, 55, 56, 63, 64, 65, 1000, 4096, 70_000]) {
      const bytes = new Uint8Array(length);
      for (let i = 0; i < length; i++) bytes[i] = (i * 7919 + length) & 0xff;
      expect(sha256FontBytes(bytes)).toBe(sha256FontBytesPure(bytes));
    }
    // A subarray view hashes the view, not the whole buffer beneath it.
    const buffer = new Uint8Array(200).fill(9);
    const view = buffer.subarray(50, 120);
    expect(sha256FontBytes(view)).toBe(sha256FontBytesPure(view));
    expect(sha256FontBytes(view)).not.toBe(sha256FontBytes(buffer));
  });
});

describe('P5: fingerprints use the native hash', () => {
  test('fingerprints are unchanged: the same digest the pure hash gives', () => {
    const part = parseXmlPart('/x.xml', '<r a="1"><b>text</b></r>');
    if (!part.ok) throw new Error('fixture');
    expect(canonicalFingerprint(part.part)).toBe(
      sha256FontBytesPure(new TextEncoder().encode(canonicalForm(part.part.root)))
    );
    const bytes = Uint8Array.from([1, 2, 3, 4]);
    expect(bytesFingerprint(bytes)).toBe(sha256FontBytesPure(bytes));
  });

  test('a large binary part hashes far faster than the pure path', () => {
    const bytes = new Uint8Array(4 * 1024 * 1024);
    for (let i = 0; i < bytes.length; i += 4096) bytes[i] = i & 0xff;
    bytesFingerprint(bytes.subarray(0, 1024)); // warm up
    const time = (run: () => string): { ms: number; digest: string } => {
      const t = performance.now();
      const digest = run();
      return { ms: performance.now() - t, digest };
    };
    // Interleaved rounds see the same load, and the fastest round of each drops the stalls a
    // parallel test run puts into any one round. The comparison itself is unchanged.
    let pureMs = Number.POSITIVE_INFINITY;
    let fastMs = Number.POSITIVE_INFINITY;
    for (let round = 0; round < 5; round += 1) {
      const pure = time(() => sha256FontBytesPure(bytes));
      const fast = time(() => bytesFingerprint(bytes));
      expect(fast.digest).toBe(pure.digest);
      pureMs = Math.min(pureMs, pure.ms);
      fastMs = Math.min(fastMs, fast.ms);
    }
    // The native hash is 40 to 80 times faster; a 4x margin keeps this stable under load.
    expect(fastMs * 4).toBeLessThan(pureMs);
  });
});
