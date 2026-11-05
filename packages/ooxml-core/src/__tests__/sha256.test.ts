// Tests for `sha256.ts` (copied from DOCX Editor at aaebb18e): the pure path pads in a fixed
// tail buffer instead of copying the input, so check every padding case against node:crypto.

import { describe, expect, test } from 'bun:test';
import { createHash } from 'node:crypto';
import { sha256FontBytes, sha256FontBytesPure } from '../sha256.ts';

const reference = (bytes: Uint8Array): string =>
  `sha256:${createHash('sha256').update(bytes).digest('hex')}`;

describe('pure SHA-256', () => {
  test('agrees with node:crypto at every padding boundary', () => {
    for (const length of [0, 1, 55, 56, 57, 63, 64, 65, 119, 120, 127, 128, 129, 1000]) {
      const bytes = Uint8Array.from({ length }, (_, i) => (i * 7 + length) & 0xff);
      expect(sha256FontBytesPure(bytes)).toBe(reference(bytes));
    }
  });

  test('hashes a view into a larger buffer by the view only', () => {
    const buffer = Uint8Array.from({ length: 300 }, (_, i) => i & 0xff);
    const view = buffer.subarray(37, 37 + 130);
    expect(sha256FontBytesPure(view)).toBe(reference(view.slice()));
    expect(sha256FontBytes(view)).toBe(reference(view.slice()));
  });

  test('two calls in a row do not share state', () => {
    const a = Uint8Array.of(1, 2, 3);
    const b = Uint8Array.of(4, 5, 6);
    const first = sha256FontBytesPure(a);
    sha256FontBytesPure(b);
    expect(sha256FontBytesPure(a)).toBe(first);
  });
});
