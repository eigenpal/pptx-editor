// Tests for LZCOMP (MTX Submission section 3.3 / Appendix C) using a test-only encoder.

import { describe, expect, test } from 'bun:test';
import {
  AdaptiveHuffman,
  BitReader,
  BitWriter,
  LzcompError,
  lzcompDecompress,
  lzSymbolLayout,
} from '../mtx-lzcomp.ts';
import { lzcompEncode, prng } from './fixtures.ts';

const LIMIT = { maxOutputBytes: 1 << 22 };

function sample(n: number, seed: number): Uint8Array {
  const rand = prng(seed);
  const out = new Uint8Array(n);
  for (let i = 0; i < n; i++) {
    // Mix of repeated phrases, runs, and noise so every symbol kind is exercised.
    const r = rand();
    if (r < 0.3 && i > 40) out[i] = out[i - 1 - Math.floor(rand() * 40)];
    else if (r < 0.5 && i > 1000) out[i] = out[i - 700 - Math.floor(rand() * 10)];
    else if (r < 0.6) out[i] = 0;
    else out[i] = Math.floor(rand() * 256);
  }
  return out;
}

describe('AdaptiveHuffman', () => {
  test('encoder and decoder trees stay in sync', () => {
    for (const range of [8, 267, 323]) {
      const rand = prng(range);
      const symbols = Array.from(
        { length: 3000 },
        () => Math.floor(rand() < 0.7 ? rand() * 10 : rand() * range) % range
      );
      const w = new BitWriter();
      const enc = new AdaptiveHuffman(range);
      for (const s of symbols) enc.writeSymbol(w, s);
      const bits = new BitReader(w.finish());
      const dec = new AdaptiveHuffman(range);
      expect(symbols.map(() => dec.readSymbol(bits))).toEqual(symbols);
    }
  });
});

describe('lzSymbolLayout', () => {
  test('matches SetDistRange', () => {
    expect(lzSymbolLayout(8)).toMatchObject({ numDistRanges: 1, dup2: 264, numSyms: 267 });
    expect(lzSymbolLayout(9)).toMatchObject({ numDistRanges: 2, dup2: 272 });
    expect(lzSymbolLayout(100000)).toMatchObject({ numDistRanges: 6, dup2: 304, numSyms: 307 });
  });
});

describe('lzcompDecompress', () => {
  test('round-trips empty input', () => {
    const out = lzcompDecompress(lzcompEncode(new Uint8Array(0)), LIMIT);
    expect(out.length).toBe(0);
  });

  test('round-trips literals and DUP codes only', () => {
    const data = Uint8Array.from([1, 2, 1, 2, 3, 4, 3, 9, 3, 7, 7, 7, 250, 0, 250]);
    const out = lzcompDecompress(lzcompEncode(data, { literalsOnly: true }), LIMIT);
    expect([...out]).toEqual([...data]);
  });

  test('round-trips copy items with short and long (>= 512) distances', () => {
    const data = sample(6000, 7);
    const out = lzcompDecompress(lzcompEncode(data), LIMIT);
    expect(out).toEqual(data);
  });

  test('copies can reference the pre-loaded history', () => {
    // 0,0..0,95 pairs and runs of four equal bytes are in the preload.
    const data = Uint8Array.from([0, 10, 0, 11, 0, 12, 0, 13, 9, 9, 9, 9, 200, 200, 200, 200]);
    const out = lzcompDecompress(lzcompEncode(data), LIMIT);
    expect([...out]).toEqual([...data]);
  });

  test('round-trips with the run-length pre-pass', () => {
    const data = new Uint8Array(5000);
    for (let i = 0; i < data.length; i++) data[i] = i % 300 < 200 ? 0 : i & 0xff;
    const out = lzcompDecompress(lzcompEncode(data, { runLength: true }), LIMIT);
    expect(out).toEqual(data);
  });

  test('enforces the output limit', () => {
    const data = new Uint8Array(4096);
    const stream = lzcompEncode(data);
    expect(() => lzcompDecompress(stream, { maxOutputBytes: 1000 })).toThrow(LzcompError);
    const rle = lzcompEncode(data, { runLength: true });
    expect(() => lzcompDecompress(rle, { maxOutputBytes: 1000 })).toThrow(LzcompError);
  });

  test('truncated or corrupted streams fail with LzcompError only', () => {
    const stream = lzcompEncode(sample(3000, 11));
    for (let n = 0; n < stream.length; n += 7) {
      expect(() => lzcompDecompress(stream.subarray(0, n), LIMIT)).toThrow(LzcompError);
    }
    const rand = prng(99);
    for (let i = 0; i < 300; i++) {
      const bad = stream.slice();
      for (let k = 0; k < 4; k++)
        bad[Math.floor(rand() * bad.length)] ^= 1 << Math.floor(rand() * 8);
      try {
        lzcompDecompress(bad, LIMIT);
      } catch (err) {
        expect(err).toBeInstanceOf(LzcompError);
      }
    }
  });
});
