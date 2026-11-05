// Tests for CTF primitives (MTX Submission sections 5.2-5.5, 5.11, 6.1, 6.2) using
// hand-computed vectors from the specification text.

import { describe, expect, test } from 'bun:test';
import { decodeCvt, decodeVdmx, MagnitudeReader } from '../mtx-ctf.ts';
import {
  encodePushInstructions,
  read255Short,
  read255UShort,
  readPushData,
  TRIPLETS,
} from '../mtx-glyf.ts';
import { ByteReader } from '../sfnt.ts';
import { FontFormatError } from '../types.ts';
import { Bytes, enc255Short, enc255UShort } from './fixtures.ts';

const reader = (bytes: number[]): ByteReader =>
  new ByteReader(Uint8Array.from(bytes), 0, 'ctf-corrupt');

describe('255USHORT / 255SHORT (6.1)', () => {
  test('255USHORT vectors', () => {
    expect(read255UShort(reader([0]))).toBe(0);
    expect(read255UShort(reader([252]))).toBe(252);
    expect(read255UShort(reader([255, 0]))).toBe(253);
    expect(read255UShort(reader([255, 255]))).toBe(508);
    expect(read255UShort(reader([254, 0]))).toBe(506);
    expect(read255UShort(reader([254, 255]))).toBe(761);
    expect(read255UShort(reader([253, 0x12, 0x34]))).toBe(0x1234);
    for (const v of [0, 100, 252, 253, 505, 506, 761, 762, 65535]) {
      expect(read255UShort(reader(enc255UShort(v)))).toBe(v);
    }
  });

  test('255SHORT vectors', () => {
    expect(read255Short(reader([249]))).toBe(249);
    expect(read255Short(reader([250, 5]))).toBe(-5);
    expect(read255Short(reader([255, 0]))).toBe(250);
    expect(read255Short(reader([255, 255]))).toBe(505);
    expect(read255Short(reader([250, 255, 255]))).toBe(-505);
    expect(read255Short(reader([254, 0]))).toBe(500);
    expect(read255Short(reader([250, 254, 255]))).toBe(-755);
    expect(read255Short(reader([253, 0x80, 0x00]))).toBe(-32768);
    for (const v of [-32768, -1000, -300, -1, 0, 1, 249, 250, 600, 32767]) {
      expect(read255Short(reader(enc255Short(v)))).toBe(v);
    }
  });

  test('truncated values throw FontFormatError', () => {
    expect(() => read255UShort(reader([253, 1]))).toThrow(FontFormatError);
    expect(() => read255Short(reader([250]))).toThrow(FontFormatError);
  });
});

describe('push data (6.2)', () => {
  test('Hop3Code: A,X1,Hop3,X2 -> A,X1,A,X2,A', () => {
    expect([...readPushData(reader([5, 7, 251, 9]), 5)]).toEqual([5, 7, 5, 9, 5]);
  });

  test('Hop4Code: A,X1,Hop4,X2,X3 -> A,X1,A,X2,A,X3,A', () => {
    expect([...readPushData(reader([3, 1, 252, 2, 4]), 7)]).toEqual([3, 1, 3, 2, 3, 4, 3]);
  });

  test('rejects hop codes without history or overrunning the count', () => {
    expect(() => readPushData(reader([251, 1]), 3)).toThrow(FontFormatError);
    expect(() => readPushData(reader([1, 2, 251, 3]), 4)).toThrow(FontFormatError);
    expect(() => readPushData(reader([1]), 2)).toThrow(FontFormatError);
  });

  test('re-encodes as TrueType PUSH instructions', () => {
    expect([...encodePushInstructions(Int32Array.from([5, 7, 5, 9, 5]))]).toEqual([
      0xb4, 5, 7, 5, 9, 5,
    ]);
    expect([...encodePushInstructions(Int32Array.from([-1, 300]))]).toEqual([
      0xb9, 0xff, 0xff, 0x01, 0x2c,
    ]);
    const many = Int32Array.from({ length: 20 }, (_, i) => i);
    const enc = encodePushInstructions(many);
    expect(enc[0]).toBe(0x40); // NPUSHB
    expect(enc[1]).toBe(20);
    expect(encodePushInstructions(new Int32Array(0)).length).toBe(0);
  });
});

describe('triplet table (5.11)', () => {
  test('has 128 entries consistent with the specification table', () => {
    expect(TRIPLETS.length).toBe(128);
    expect(TRIPLETS[0]).toMatchObject({ byteCount: 2, xBits: 0, yBits: 8, deltaY: 0, ySign: -1 });
    expect(TRIPLETS[9]).toMatchObject({ deltaY: 1024, ySign: 1 });
    expect(TRIPLETS[10]).toMatchObject({ xBits: 8, yBits: 0, deltaX: 0, xSign: -1 });
    expect(TRIPLETS[19]).toMatchObject({ deltaX: 1024, xSign: 1 });
    expect(TRIPLETS[20]).toMatchObject({ xBits: 4, deltaX: 1, deltaY: 1, xSign: -1, ySign: -1 });
    expect(TRIPLETS[21]).toMatchObject({ xSign: 1, ySign: -1 });
    expect(TRIPLETS[22]).toMatchObject({ xSign: -1, ySign: 1 });
    expect(TRIPLETS[24]).toMatchObject({ deltaX: 1, deltaY: 17 });
    expect(TRIPLETS[36]).toMatchObject({ deltaX: 17, deltaY: 1 });
    expect(TRIPLETS[83]).toMatchObject({ deltaX: 49, deltaY: 49, xSign: 1, ySign: 1 });
    expect(TRIPLETS[84]).toMatchObject({ byteCount: 3, xBits: 8, deltaX: 1, deltaY: 1 });
    expect(TRIPLETS[88]).toMatchObject({ deltaX: 1, deltaY: 257 });
    expect(TRIPLETS[96]).toMatchObject({ deltaX: 257, deltaY: 1 });
    expect(TRIPLETS[119]).toMatchObject({ deltaX: 513, deltaY: 513, xSign: 1, ySign: 1 });
    expect(TRIPLETS[120]).toMatchObject({ byteCount: 4, xBits: 12, yBits: 12, deltaX: 0 });
    expect(TRIPLETS[127]).toMatchObject({ byteCount: 5, xBits: 16, xSign: 1, ySign: 1 });
  });
});

describe('cvt (5.2)', () => {
  test('decodes 1-3 byte codes as running deltas', () => {
    const ctf = new Bytes()
      .u16(5)
      .u8(10) // +10 -> 10
      .u8(248, 5) // +(238+5) -> 253
      .u8(239, 3) // -3 -> 250
      .u8(241, 24) // -(2*238+24) = -500 -> -250
      .u8(238, 0x80, 0x00) // -32768 -> -250-32768 wraps to 32518
      .done();
    const out = decodeCvt(ctf);
    const values = [0, 2, 4, 6, 8].map((p) => ((out[p] << 24) >> 16) | out[p + 1]);
    expect(values).toEqual([10, 253, 250, -250, 32518]);
  });

  test('rejects truncated cvt', () => {
    expect(() => decodeCvt(new Bytes().u16(3).u8(1, 248).done())).toThrow(FontFormatError);
  });
});

describe('magnitude dependent encoding (5.3)', () => {
  test('decodes the specification example 0x90 0x48 0x84', () => {
    const r = new MagnitudeReader(Uint8Array.from([0x90, 0x48, 0x84]), 0);
    const values = Array.from({ length: 13 }, () => r.next());
    expect(values).toEqual([0, 0, 0, 0, 1, 1, 0, 1, 1, 0, 1, 0, 0]);
  });

  test('runs past the end throw', () => {
    const r = new MagnitudeReader(Uint8Array.from([0xff]), 0);
    expect(() => r.next()).toThrow(FontFormatError);
  });
});

describe('VDMX (5.5)', () => {
  test('rebuilds groups from predicted values plus surprises', () => {
    // One ratio, one group: entries (10, 11, -3) and (11, 11, -4) with yMaxMult=2048 and
    // yMinMult=512 encode as surprises (+2,+1,0) and (0,0,-1).
    const ctf = new Bytes()
      .u16(1, 1, 1)
      .u8(1, 1, 1, 1)
      .u16(12)
      .u16(2, 2048, 512)
      .u8(0x13, 0x14)
      .done();
    const out = decodeVdmx(ctf);
    const expected = new Bytes()
      .u16(1, 1, 1)
      .u8(1, 1, 1, 1)
      .u16(12)
      .u16(2)
      .u8(10, 11)
      .u16(10, 11, 0xfffd, 11, 11, 0xfffc)
      .done();
    expect([...out]).toEqual([...expected]);
  });

  test('uncompressed tables are copied with the version restored', () => {
    const ctf = new Bytes().u16(0xffff - 1, 0, 0).done();
    const out = decodeVdmx(ctf);
    expect([...out]).toEqual([0, 1, 0, 0, 0, 0]);
  });
});
