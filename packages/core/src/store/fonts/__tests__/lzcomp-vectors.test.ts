// Test vectors for the LZCOMP decoder and the test-side encoder.
// The expected values follow the LZCOMP compression of the W3C MTX submission, section 3.
// Model state is not exported, so the slot tables are checked through the codes they imply.

import { createHash } from 'node:crypto';
import { describe, expect, test } from 'bun:test';
import {
  AdaptiveHuffman,
  BitReader,
  BitWriter,
  LzcompError,
  lzcompDecompress,
  LZ_MAX_2BYTE_DIST,
  LZ_PRELOAD_SIZE,
  lzSymbolLayout,
  writePreload,
  type LzcompOptions,
} from '../mtx-lzcomp.ts';
import { lzcompEncode, rlePack } from './fixtures.ts';

// --- helpers -------------------------------------------------------------------------------

const DEFAULT: LzcompOptions = { maxOutputBytes: 4194304 };

function hex(text: string): Uint8Array {
  const tokens = text.split(/\s+/).filter((t) => t.length > 0);
  return Uint8Array.from(tokens, (t) => parseInt(t, 16));
}

function toHex(bytes: Uint8Array): string {
  return Array.from(bytes, (b) => b.toString(16).padStart(2, '0')).join(' ');
}

function sha256(data: Uint8Array | string): string {
  return createHash('sha256').update(data).digest('hex');
}

function repeat(unit: number[], times: number): number[] {
  const out: number[] = [];
  for (let i = 0; i < times; i++) for (const b of unit) out.push(b);
  return out;
}

/** 'ok', or the kind of the LzcompError. Any other exception fails the test. */
function outcome(data: Uint8Array, options: LzcompOptions = DEFAULT): string {
  try {
    lzcompDecompress(data, options);
    return 'ok';
  } catch (err) {
    if (err instanceof LzcompError) return err.kind;
    throw err;
  }
}

function decode(data: Uint8Array, options: LzcompOptions = DEFAULT): string {
  return toHex(lzcompDecompress(data, options));
}

/** Writes `symbol` with the model (which then updates) and returns the emitted code bits. */
function emit(model: AdaptiveHuffman, symbol: number): string {
  const w = new BitWriter();
  model.writeSymbol(w, symbol);
  w.writeBit(1); // sentinel: the code ends before the last 1 bit
  const bits = Array.from(w.finish(), (b) => b.toString(2).padStart(8, '0')).join('');
  return bits.slice(0, bits.lastIndexOf('1'));
}

function modelAfter(size: number, history: readonly number[]): AdaptiveHuffman {
  const model = new AdaptiveHuffman(size);
  const discard = new BitWriter();
  for (const s of history) model.writeSymbol(discard, s);
  return model;
}

/** Code of every symbol in the state reached after `history`, by replay on fresh models. */
function codeTable(size: number, history: readonly number[] = []): string[] {
  return Array.from({ length: size }, (_, symbol) => emit(modelAfter(size, history), symbol));
}

function codesOf(size: number, history: readonly number[], symbols: number[]): string[] {
  return symbols.map((s) => emit(modelAfter(size, history), s));
}

function lengthHistogram(codes: string[]): Record<number, number> {
  const out: Record<number, number> = {};
  for (const c of codes) out[c.length] = (out[c.length] ?? 0) + 1;
  return out;
}

// --- 2, 10.1 bit input and output, errors --------------------------------------------------

describe('bit input and output (sections 2, 10.1)', () => {
  test('reads fields first bit most significant, then fails past the end', () => {
    const r = new BitReader(hex('a5 0f'));
    expect(r.readBit()).toBe(1);
    expect(r.readValue(3)).toBe(2);
    expect(r.readValue(12)).toBe(1295);
    expect(() => r.readBit()).toThrow(LzcompError);
    try {
      r.readBit();
    } catch (err) {
      expect((err as LzcompError).kind).toBe('corrupt');
    }
  });

  test('reads only inside a subarray view', () => {
    const backing = hex('ff a5 0f ff');
    const r = new BitReader(backing.subarray(1, 3));
    expect(r.readValue(16)).toBe(0xa50f);
    expect(() => r.readBit()).toThrow(LzcompError);
  });

  test('writes in the same order and pads the last byte with zeros', () => {
    const w = new BitWriter();
    w.writeBit(1);
    w.writeValue(5, 3);
    w.writeValue(0xabc, 12);
    expect(toHex(w.finish())).toBe('da bc');
    const w2 = new BitWriter();
    w2.writeValue(0x1ff, 9);
    expect(toHex(w2.finish())).toBe('ff 80');
    expect(toHex(w2.finish())).toBe('ff 80');
  });

  test('readValue(0) on an empty reader is 0', () => {
    expect(new BitReader(new Uint8Array(0)).readValue(0)).toBe(0);
  });

  test('readSymbol on an empty reader fails with corrupt', () => {
    const m = new AdaptiveHuffman(8);
    expect(() => m.readSymbol(new BitReader(new Uint8Array(0)))).toThrow(LzcompError);
  });

  test('writeSymbol outside the alphabet fails with corrupt and writes nothing', () => {
    for (const bad of [8, -1, 2.5]) {
      const m = new AdaptiveHuffman(8);
      const w = new BitWriter();
      let kind = '';
      try {
        m.writeSymbol(w, bad);
      } catch (err) {
        kind = err instanceof LzcompError ? err.kind : 'other';
      }
      expect(kind).toBe('corrupt');
      expect(w.finish().length).toBe(0);
    }
  });

  test('model sizes are validated', () => {
    for (const bad of [0, 1, 2.5, 4097, 257, Number.NaN]) {
      let kind = '';
      try {
        new AdaptiveHuffman(bad);
      } catch (err) {
        kind = err instanceof LzcompError ? err.kind : 'other';
      }
      expect(kind).toBe('corrupt');
    }
    for (const ok of [2, 3, 256, 258, 511, 512, 4096]) {
      expect(() => new AdaptiveHuffman(ok)).not.toThrow();
    }
  });

  test('LzcompError defaults to corrupt', () => {
    const err = new LzcompError('x');
    expect(err.kind).toBe('corrupt');
    expect(err.name).toBe('LzcompError');
    expect(err).toBeInstanceOf(Error);
    expect(new LzcompError('y', 'limit').kind).toBe('limit');
  });
});

// --- 4.2, 4.5 layout and preload -----------------------------------------------------------

describe('block layout and preload (sections 1.3, 4.2, 4.5)', () => {
  test('constants', () => {
    expect(LZ_PRELOAD_SIZE).toBe(7168);
    expect(LZ_MAX_2BYTE_DIST).toBe(512);
  });

  test('lzSymbolLayout follows the size table', () => {
    const rows: Array<[number, number, number]> = [
      [0, 1, 267],
      [8, 1, 267],
      [9, 2, 275],
      [64, 2, 275],
      [65, 3, 283],
      [512, 3, 283],
      [513, 4, 291],
      [4096, 4, 291],
      [4097, 5, 299],
      [32768, 5, 299],
      [32769, 6, 307],
      [262144, 6, 307],
      [262145, 7, 315],
      [2097152, 7, 315],
      [2097153, 8, 323],
      [16777215, 8, 323],
    ];
    for (const [length, n, size] of rows) {
      const dup2 = 256 + 8 * n;
      expect(lzSymbolLayout(length)).toEqual({
        numDistRanges: n,
        dup2,
        dup4: dup2 + 1,
        dup6: dup2 + 2,
        numSyms: size,
      });
    }
    expect(lzSymbolLayout(100000)).toMatchObject({ numDistRanges: 6, dup2: 304, numSyms: 307 });
  });

  test('preload spot values and hash', () => {
    const p = new Uint8Array(LZ_PRELOAD_SIZE + 4).fill(0xaa);
    writePreload(p);
    expect(toHex(p.subarray(0, 8))).toBe('00 00 00 01 00 02 00 03');
    expect(toHex(p.subarray(190, 194))).toBe('00 5f 01 00');
    expect(toHex(p.subarray(6140, 6152))).toBe('1f 5e 1f 5f 00 00 00 00 01 01 01 01');
    expect(toHex(p.subarray(7156, 7168))).toBe('fd fd fd fd fe fe fe fe ff ff ff ff');
    expect(sha256(p.subarray(0, LZ_PRELOAD_SIZE))).toBe(
      '0789240870c8fdf8a355edadedf35bc0d63b535c9e8e28388bc9f884ad49f843'
    );
    expect(toHex(p.subarray(LZ_PRELOAD_SIZE))).toBe('aa aa aa aa');
  });

  test('preload into a short target writes only what exists', () => {
    const short = new Uint8Array(5);
    writePreload(short);
    expect(toHex(short)).toBe('00 00 00 01 00');
  });
});

// --- 3.5, 10.2 size-8 model ----------------------------------------------------------------

describe('size-8 model (sections 3.5, 10.2)', () => {
  test('fresh codes are the 3-bit binary values', () => {
    expect(codeTable(8)).toEqual(['000', '001', '010', '011', '100', '101', '110', '111']);
  });

  test('increment steps', () => {
    const steps: Array<[number, string, string]> = [
      [3, '011', '011 001 010 000 100 101 110 111'],
      [3, '000', '011 001 010 000 100 101 110 111'],
      [5, '101', '011 101 010 000 100 001 110 111'],
      [0, '011', '010 101 011 000 100 001 110 111'],
      [3, '000', '010 101 011 000 100 001 110 111'],
      [7, '111', '010 101 111 000 100 001 110 011'],
      [7, '011', '010 101 111 000 100 011 110 001'],
      [7, '001', '010 101 111 000 100 011 110 001'],
    ];
    const live = new AdaptiveHuffman(8);
    const history: number[] = [];
    for (const [symbol, emitted, after] of steps) {
      expect(emit(live, symbol)).toBe(emitted);
      history.push(symbol);
      expect(codeTable(8, history).join(' ')).toBe(after);
    }
  });

  test('the final slot table implies the final codes', () => {
    // Slot -> symbol from the section 10.2 table; the path to slot k follows k's binary digits.
    const slotOfSymbol = [10, 13, 15, 8, 12, 11, 14, 9];
    const implied = slotOfSymbol.map((slot) => slot.toString(2).slice(1));
    expect(codeTable(8, [3, 3, 5, 0, 3, 7, 7, 7])).toEqual(implied);
  });

  test('a decoder reads the codes back', () => {
    const symbols = [3, 3, 5, 0, 3, 7, 7, 7];
    const w = new BitWriter();
    const enc = new AdaptiveHuffman(8);
    for (const s of symbols) enc.writeSymbol(w, s);
    const r = new BitReader(w.finish());
    const dec = new AdaptiveHuffman(8);
    expect(symbols.map(() => dec.readSymbol(r))).toEqual(symbols);
  });
});

// --- 10.3 main models ----------------------------------------------------------------------

describe('main models (section 10.3)', () => {
  test('fresh size 267', () => {
    const codes = codeTable(267);
    expect(lengthHistogram(codes)).toEqual({ 5: 1, 6: 1, 8: 223, 9: 42 });
    const expected: Record<number, string> = {
      0: '000001011',
      1: '000001100',
      2: '000010011',
      65: '01001100',
      127: '10001010',
      128: '10001011',
      254: '000001001',
      255: '000001010',
      256: '00001011',
      257: '00001100',
      258: '000001101',
      263: '000010010',
      264: '00011',
      265: '001000',
      266: '000010101',
    };
    for (const [symbol, code] of Object.entries(expected)) expect(codes[Number(symbol)]).toBe(code);
  });

  test('size 267 after increments', () => {
    const sequence = [65, 65, 264, 0, 256, 65, 266, 266, 266];
    const live = new AdaptiveHuffman(267);
    expect(sequence.map((s) => emit(live, s))).toEqual([
      '01001100',
      '00010111',
      '00011',
      '000001011',
      '00001011',
      '0010010',
      '000010101',
      '00011011',
      '0010101',
    ]);
    const symbols = [0, 1, 64, 65, 66, 255, 256, 257, 264, 265, 266];
    expect(codesOf(267, sequence, symbols)).toEqual([
      '00011010',
      '000001100',
      '01001011',
      '0001100',
      '01001101',
      '000001010',
      '0010101',
      '00001100',
      '00100',
      '000111',
      '0010100',
    ]);
  });

  test('fresh size 323', () => {
    const codes = codeTable(323);
    expect(lengthHistogram(codes)).toEqual({ 5: 1, 6: 1, 8: 167, 9: 154 });
    const expected: Record<number, string> = {
      0: '001000011',
      1: '001000100',
      2: '010000011',
      65: '10000100',
      127: '11000010',
      128: '11000011',
      254: '001000001',
      255: '001000010',
      256: '01000011',
      257: '01000100',
      258: '001000101',
      263: '001001010',
      319: '010000010',
      320: '01010',
      321: '010110',
      322: '010000101',
    };
    for (const [symbol, code] of Object.entries(expected)) expect(codes[Number(symbol)]).toBe(code);
  });

  test('size 323 after increments', () => {
    const sequence = [320, 320, 0, 0, 0, 300, 255];
    const live = new AdaptiveHuffman(323);
    expect(sequence.map((s) => emit(live, s))).toEqual([
      '01010',
      '01010',
      '001000011',
      '01001111',
      '0101110',
      '001101111',
      '001000010',
    ]);
    const symbols = [0, 1, 255, 256, 257, 300, 319, 320, 321, 322];
    expect(codesOf(323, sequence, symbols)).toEqual([
      '0101100',
      '001000100',
      '01011011',
      '01000011',
      '01000100',
      '01011010',
      '010000010',
      '01010',
      '010111',
      '010000101',
    ]);
  });

  test('fresh codes of the other main sizes', () => {
    const rows: Array<[number, string[]]> = [
      [275, ['000010011', '000010010', '00010011', '00100', '001010', '000100101']],
      [283, ['000011011', '000011010', '00011011', '00101', '001100', '000110101']],
      [291, ['000100011', '000100010', '00100011', '00110', '001110', '001000101']],
      [299, ['000101011', '000101010', '00101011', '00111', '010000', '001010101']],
      [307, ['000110011', '000110010', '00110011', '01000', '010010', '001100101']],
      [315, ['000111011', '000111010', '00111011', '01001', '010100', '001110101']],
    ];
    for (const [size, expected] of rows) {
      expect(codesOf(size, [], [0, 255, 256, size - 3, size - 2, size - 1])).toEqual(expected);
    }
  });
});

// --- 10.4 model hash vectors ---------------------------------------------------------------

describe('model hash vectors (section 10.4)', () => {
  const rows: Array<[number, string, number, string, string, string]> = [
    [
      8,
      '74c691ad79435c1b63bc666b42ae131acd2d5e2b9e80eafcb728b9df5831125d',
      685,
      '0d 91 d5 74 93 52 24 96 1b b8 4d 66 a0 32 c4 ed',
      'b9176cddc0e806481f524ff448d3a8496d159ebd46873be744eb8d59f92c7a47',
      '94ec565f89a8fffd56d5e86f0d3aabab380c26a2292afd485ca2da6b6ec58693',
    ],
    [
      267,
      'd5a0512aef4ddb80064c25508f25941831bfe454a876997419e1c5671fe56b53',
      1501,
      '05 87 83 81 42 e3 a0 62 c3 50 2e 52 b0 27 c4 00',
      '0498c6ae560e3c3ca61a14b18aeb2fc17ce8701aec0c21ffdb4f1d2e57879df5',
      '0f53b6b450eb34cc374f8df640ef75016da65bd808d8fb09b0682f9b2833d63d',
    ],
    [
      275,
      'cc0256eaa0edd64483c797e40e137ce54325482f57a0200e4dbd815dd4743589',
      1527,
      '09 8b 85 82 43 64 20 a3 03 90 4e 72 d0 47 d4 08',
      '1d2baaba1c2887687dc9980f9ab5f8fe645df54caca58e708cb6d1cedc2bc826',
      'd2566c6cd9fcc0096695bc16e02bcb49b23678654b9ba5d6a7ea1a241ab40f60',
    ],
    [
      283,
      '5a555d714ca13acabbca27789e96fbcea5209e05a544aaba6f83e20613eb8309',
      1698,
      '0d 8f 87 83 43 e4 a0 e3 43 d0 6e 92 f0 67 e4 10',
      'dc8d3f5d895199bf5d3b7be982f576f2276119037baaaa36fb4e3c7714306dad',
      'd6f4583dcebf1478356f81328cda3f299e8b9d6637be515ce70b56cf29c73cfb',
    ],
    [
      291,
      '58dce91e9f231a58ac2c28c1d191ea329ad5f7002dcb7a436288868b4b672d77',
      1529,
      '11 93 89 84 44 65 21 23 84 10 8e b3 10 87 f4 18',
      'e40f176e09b609f39bf4f1679e6d7b277af3dcf9fd57a17c3700ff5b64a7ca81',
      'b4d06824355bdad5bfa462642c8537848f520698171dd16edf3d6d6ff8d4ab79',
    ],
    [
      299,
      '5e5943a9a7c886d0b1ddfc55e7bd787b60e1fba62089904c39432ccc58ae6602',
      1610,
      '15 97 8b 85 44 e5 a1 63 c4 50 ae d3 30 a6 02 10',
      'c21d0c2ff38c2857a075560310ab35db2ae47306f976bcb2200c0d42fc34241a',
      '84c6fa7322eecf76b135a60a26e7e91d9cc173b81fac7597ab6fa9960cbda3aa',
    ],
    [
      307,
      'f03ee4a90376d9dd47dc19b7b9507f94380a91bbcb7df2125ed71f21dcfc20a2',
      1719,
      '19 9b 8d 86 45 66 21 a4 04 90 ce f3 50 c6 0a 14',
      'b7e92e522d60b64ff849c669e5fa566c9c0085f1d543cbe2d45532039ca64060',
      '2afdbee35dbeee31d269c26309afde968799b666b85c1f08429f13d7c52736fb',
    ],
    [
      315,
      'c3e637d23d285319f36ff7bc7e3ecf6e3a83952314e645adffe20733ba407f2c',
      1442,
      '1d 9f 8f 87 45 e6 a1 e4 44 d0 ef 13 70 e6 12 18',
      'b03e6cabe2d6ff7b9c6b87d3bc4b16c8c39a74e5b564197a115ebcd7419f7e56',
      'd6fb91ed5c3f36abe31081d7d3470a955a79b826d8b071621f0e5a8122934f98',
    ],
    [
      323,
      'e9655a5f59531b205baf53a3c2a7633cb1000402a33f98bef8780d1c9f5e7be0',
      1641,
      '21 a3 91 88 46 67 22 24 85 11 0f 33 91 06 1a 1c',
      '8125a572af568668265bd5e8ae6b1f52e3c555e488cd7b143996f6b1e729a978',
      'fc6ccc26a3f7339d859b2d76425fdad072f3d87c603767031a2b9bcd7a33b00f',
    ],
  ];

  const tableHash = (codes: string[]): string => sha256(codes.map((c) => `${c}\n`).join(''));

  for (const [size, freshHash, streamLength, head, streamHash, afterHash] of rows) {
    test(`size ${size}`, () => {
      const sequence = Array.from({ length: 2000 }, (_, k) =>
        k % 3 === 0 ? k % 5 : (k * k + 3 * k + 7) % size
      );
      expect(tableHash(codeTable(size))).toBe(freshHash);

      const w = new BitWriter();
      const enc = new AdaptiveHuffman(size);
      for (const s of sequence) enc.writeSymbol(w, s);
      const stream = w.finish();
      expect(stream.length).toBe(streamLength);
      expect(toHex(stream.subarray(0, 16))).toBe(head);
      expect(sha256(stream)).toBe(streamHash);

      const r = new BitReader(stream);
      const dec = new AdaptiveHuffman(size);
      expect(sequence.map(() => dec.readSymbol(r))).toEqual(sequence);

      expect(tableHash(codeTable(size, sequence))).toBe(afterHash);
    });
  }
});

// --- 10.5 decoder vectors ------------------------------------------------------------------

const V1 = hex('00 00 02 a9 b8 3b 92 bd 00');
const HELLO = '48 65 6c 6c 6f';
const V8D = hex('80 00 02 44 8b 82 93 00');

/** The same stream with the run-length flag bit cleared: the decoder returns the LZ bytes. */
function withoutFlag(stream: Uint8Array): Uint8Array {
  const out = stream.slice();
  out[0] &= 0x7f;
  return out;
}

describe('decoder vectors (section 10.5)', () => {
  test('V1 literals only', () => {
    expect(decode(V1)).toBe(HELLO);
    expect(lzcompEncode(hex(HELLO), { literalsOnly: true })).toEqual(V1);
  });

  test('V2 DUP2, DUP4, and DUP6 inside the data', () => {
    expect(decode(hex('00 00 04 85 04 62 40 a0 8c 48 52 44 a0'))).toBe(
      '01 02 03 04 05 06 05 04 03'
    );
  });

  test('V3 single-byte copies from the preload', () => {
    expect(decode(hex('00 00 01 85 48 18'))).toBe('fe ff fe');
  });

  test('V4 short-distance copies', () => {
    const out = lzcompDecompress(hex('00 00 1f 3a 3a bb 09 91 42 18 44 34 01 17 44'), DEFAULT);
    expect(toHex(out)).toBe(
      '61 62 63 61 62 63 61 62 61 62 63 61 62 63 61 62 61 62 63 61 62 63 61 62 61 62 63 61 62 63 61 62 ' +
        '62 63 61 62 63 61 62 61 62 63 61 62 63 61 62 61 62 63 61 62 63 61 62 61 62 63 61 62 63 61'
    );
  });

  test('V5 copies at distances 511, 512, and 600', () => {
    const stream = hex(
      '00 01 46 88 c4 84 32 20 8c 48 43 22 08 c4 82 90 05 11 03 da 81 f1 e0 24 70 29 4b 5b 9a 33 e7 26 13 90 31 f2 c0'
    );
    const expected: number[] = [];
    for (let i = 0; i < 640; i++) expected.push(i % 10);
    expected.push(0xaa, 0x09, 0x00, 0x09, 0x00, 0x01, 0, 1, 2, 3, 4, 5, 6);
    expect(Array.from(lzcompDecompress(stream, DEFAULT))).toEqual(expected);
  });

  test('V6 copy from the end of the preload in a short block', () => {
    expect(decode(hex('00 00 02 83 7a 60'))).toBe('fd fd fd fe 41');
  });

  test('V7 deep preload references', () => {
    const stream = hex(
      '00 08 04 13 0d e6 8b 58 c1 11 90 07 bc 47 c4 49 09 54 b5 27 55 94 88 42 60 40 50 3b 15 10'
    );
    const out = lzcompDecompress(stream, DEFAULT);
    const expected = repeat([0x00, 0x0a, 0x00, 0x0b], 1024);
    expected.push(0x41, 0x41, 0x41, 0x41, 0x42, 0x42, 0x42, 0x42);
    expect(Array.from(out)).toEqual(expected);
    expect(sha256(out)).toBe('3de3d2aee37a4efc4147073e4b302ab689c65af8e2ebd30d47479f7cbc20c050');
  });

  test('V8 run-length stage', () => {
    const stream = hex('80 00 05 c8 aa 0f aa 2a 93 09 ab 11 12 2b 80');
    expect(decode(stream)).toBe('41 42 42 42 42 42 7e 43 44 44 44');
    expect(decode(withoutFlag(stream))).toBe('7e 41 7e 05 42 7e 00 43 7e 03 44');
  });

  test('V8b pending escape at the end', () => {
    const stream = hex('80 00 01 c4 a6 0b 80');
    expect(decode(stream)).toBe('41');
    expect(decode(withoutFlag(stream))).toBe('7e 41 7e');
  });

  test('V8c pending count at the end', () => {
    const stream = hex('80 00 02 44 a6 0b 86 80');
    expect(decode(stream)).toBe('41');
    expect(decode(withoutFlag(stream))).toBe('7e 41 7e 09');
  });

  test('V8d a run of 255 at the cap', () => {
    expect(decode(withoutFlag(V8D))).toBe('7e 7e ff 41');
    const out = lzcompDecompress(V8D, { maxOutputBytes: 255 });
    expect(Array.from(out)).toEqual(repeat([0x41], 255));
    expect(outcome(V8D, { maxOutputBytes: 254 })).toBe('limit');
  });

  test('V8f the escape value as data, and a run of the escape value', () => {
    const stream = hex('80 00 03 82 c5 c9 12 05 94 0b');
    expect(decode(stream)).toBe('00 00 00 00 00 05');
    expect(decode(withoutFlag(stream))).toBe('00 00 00 00 04 00 05');
  });

  test('V9 version 1 has no flag bit', () => {
    const v1 = { maxOutputBytes: 4194304, version: 1 };
    expect(decode(hex('00 00 05 53 70 77 25 7a'), v1)).toBe(HELLO);
    expect(decode(V1, v1)).toBe('9e ad');
    expect(decode(V1, { maxOutputBytes: 4194304, version: 3 })).toBe(HELLO);
  });

  test('V10 empty blocks', () => {
    expect(decode(hex('00 00 00 00'))).toBe('');
    expect(decode(hex('00 00 00'), { maxOutputBytes: 4194304, version: 1 })).toBe('');
  });

  test('V11 a zero first length chunk with more chunks', () => {
    expect(decode(hex('00 00 02 a6 26 83 c9'))).toBe('41 42 ff ff 41');
  });

  test('V12 more distance groups than the distance needs', () => {
    expect(decode(hex('00 00 04 ba 3a bb 07 01 26 00'))).toBe('61 62 63 ff 61 62 ff 61 62');
  });

  test('V13 trailing bytes are ignored', () => {
    const padded = new Uint8Array(V1.length + 2).fill(0xff);
    padded.set(V1);
    expect(decode(padded)).toBe(HELLO);
  });

  test('the output does not share memory with the input and the input is unchanged', () => {
    const input = V1.slice();
    const out = lzcompDecompress(input, DEFAULT);
    expect(out.buffer).not.toBe(input.buffer);
    expect(toHex(input)).toBe(toHex(V1));
  });
});

// --- 10.6 error vectors --------------------------------------------------------------------

describe('error vectors (section 10.6)', () => {
  test('E1 bits run out', () => {
    expect(outcome(V1.subarray(0, 8))).toBe('corrupt');
  });

  test('E2 declared length over the cap; E2b at the cap', () => {
    expect(outcome(V1, { maxOutputBytes: 4 })).toBe('limit');
    expect(decode(V1, { maxOutputBytes: 5 })).toBe(HELLO);
  });

  test('E3 source before the window; E3b source at index 0', () => {
    const e3 = hex(
      '00 08 02 13 0d e8 8b 68 c1 31 90 47 bc 03 e2 24 84 b2 5a a3 aa ca 44 25 30 20 38 00'
    );
    expect(outcome(e3)).toBe('corrupt');
    const e3b = hex(
      '00 08 02 13 0d ea 0b 68 c1 41 90 27 bc 03 e2 34 84 b2 5a 93 aa ca 44 25 30 20 38 00'
    );
    const out = lzcompDecompress(e3b, DEFAULT);
    expect(Array.from(out)).toEqual(repeat([0x00, 0x00, 0x00, 0x01], 1025));
    expect(sha256(out)).toBe('744f891ae1ef653322cced21e968801730e7b201cf0a6d5fe873cc2c45cad258');
  });

  test('E4 copy past the declared length; E4b length 2 from the preload', () => {
    expect(outcome(hex('00 00 01 a6 06 00'))).toBe('corrupt');
    expect(decode(hex('00 00 01 a6 05 80'))).toBe('41 ff 41');
  });

  test('E5 length guard; E6 long copy past the declared length', () => {
    expect(outcome(hex('00 00 04 04 b8 01 50 30'))).toBe('corrupt');
    expect(outcome(hex('00 00 04 04 b8 60'))).toBe('corrupt');
  });

  test('E7 run-length expansion over the cap', () => {
    expect(outcome(V8D, { maxOutputBytes: 254 })).toBe('limit');
  });

  test('E8 and E9 on the medium stream', () => {
    const stream = lzcompEncode(mediumInput());
    expect(outcome(stream, { maxOutputBytes: 2999 })).toBe('limit');
    expect(outcome(stream.subarray(0, 200))).toBe('corrupt');
  });
});

// --- 10.7 encoder vectors ------------------------------------------------------------------

function mediumInput(): Uint8Array {
  const out = new Uint8Array(3000);
  for (let i = 0; i < out.length; i++) {
    if (i % 700 < 500) out[i] = (7 * i + (Math.floor(i / 32) % 4)) % 256;
    else if (i % 13 === 0) out[i] = 0;
    else out[i] = (i * i) % 64;
  }
  return out;
}

const MEDIUM_STREAM = `
00 05 dc 08 cb 85 92 81 5d 19 35 51 6d 89 a5 c1 dd fa 16 32 4e 6a 86 a2 be da f7 13 2f 4b 67 83
9f bb d7 f0 08 0b 09 06 44 02 32 64 2d b7 38 64 32 58 27 b7 15 5d 1e e0 a2 64 25 e7 a9 6b 2c ee
b0 72 33 f5 b7 79 7b 3c fe 80 20 80 78 58 3a 24 cb 12 50 e9 01 23 77 65 f2 3e b2 e5 cb a4 e9 52
a4 dd 15 4d 85 bd f6 36 6e a6 df 17 4f 87 bf f0 14 18 13 0d 08 6a 80 bd 37 bd c2 00 c9 68 a7 ca
16 02 74 96 76 c1 85 50 b6 d9 70 b2 50 2b 8c 31 28 95 c4 13 ee 02 6d f1 12 c9 85 b2 96 8e c6 ff
37 6f a7 d8 08 12 10 0b 87 88 7f 85 52 c9 44 16 1e 07 f2 93 81 1d 3f 41 9e 6f e0 29 68 18 02 2e
97 41 21 90 47 90 20 ba 5a 28 fc 7f 81 81 a1 40 d8 88 b6 10 0d 10 c7 00 b6 ca 60 a0 88 60 3e 32
f3 b9 9d 00 fb 01 2f 36 fc 20 23 35 51 e0 00 1f c8 a0 2a e0 61 21 c0 82 84 88 fd 85 d0 87 08 50
28 80 c5 64 84 6d 30 61 7e 63 80 c3 7e 2d 44 87 0c 2e 45 88 1d 24 60 a2 a2 06 88 c4 03 47 60 87
fd 09 20 3d 25 c0 89 95 07 23 cc 22 30 b1 18 e2 9a 1e a2 27 a5 c9 16 eb b1 13 34 f1 c5 e4 2d 22
4b 96 11 40 eb 18 9e f8 02 30 e3 a3 10 87 96 73 14 d2 5b e9 32 0c 4b 40 f5 a4 0a 9b a2 80 ec 18
7c 38 7f 52 e0 af 9a 6e 5f e4 ac 89 50 7f 6c e7 4b ca 41 82 0e b1 9e 9b 20 9d 48 a4 d8 13 13 b7
05 4a 01 12 6c 89 0f 28 e7 29 fd fd b0 b4 ef 12 28 fe ee 55 2e 0a 10 40 f0 98`;

function expectRoundTrip(input: Uint8Array, stream: Uint8Array): void {
  expect(lzcompDecompress(stream, DEFAULT)).toEqual(input);
}

describe('encoder vectors (section 10.7)', () => {
  test('ENC1 literals and DUP2 only', () => {
    const input = hex('01 02 01 02 03 04 03 09 03 07 07 07 fa 00 fa');
    const expected = '00 00 07 85 04 64 20 90 28 41 51 02 88 08 0d 09 a0';
    expect(toHex(lzcompEncode(input, { literalsOnly: true }))).toBe(expected);
    expect(toHex(lzcompEncode(input))).toBe(expected);
    expectRoundTrip(input, hex(expected));
  });

  test('ENC2 preload-friendly input', () => {
    const input = hex('00 0a 00 0b 00 0c 00 0d 09 09 09 09 c8 c8 c8 c8');
    const stream = lzcompEncode(input);
    expect(toHex(stream)).toBe('00 00 08 04 c5 a4 2e 21 79 07 02 a3 02 26 d9 8a 20');
    expectRoundTrip(input, stream);
  });

  const abc = new TextEncoder().encode('abcabcabcabc-xyzxyzxyz-abcabcabcabc!');

  test('ENC3 copies', () => {
    const stream = lzcompEncode(abc);
    expect(toHex(stream)).toBe('00 00 12 3a 3a bb 0a 00 c0 08 11 71 91 a6 41 30 46 21 44 4f 00');
    expectRoundTrip(abc, stream);
  });

  test('ENC4 copies with the run-length stage', () => {
    const payload = rlePack(abc);
    expect(payload.length).toBe(37);
    expect(payload[0]).toBe(0);
    expect(payload.subarray(1)).toEqual(abc);
    const stream = lzcompEncode(abc, { runLength: true });
    expect(toHex(stream)).toBe('80 00 12 84 dd 1d 5d 85 00 60 04 08 b8 c8 d3 40 d0 24 10 a2 23 80');
    expectRoundTrip(abc, stream);
  });

  test('ENC5 empty input', () => {
    const empty = new Uint8Array(0);
    expect(toHex(lzcompEncode(empty))).toBe('00 00 00 00');
    expect(toHex(lzcompEncode(empty, { runLength: true }))).toBe('80 00 00 82 c0');
    expectRoundTrip(empty, lzcompEncode(empty));
    expectRoundTrip(empty, lzcompEncode(empty, { runLength: true }));
  });

  test('ENC6 medium vector', () => {
    const input = mediumInput();
    expect(sha256(input)).toBe('ade8a8302595668f53947e257d48faf7b292df2c3d038cfcf291bcbf78c8952b');
    const stream = lzcompEncode(input);
    expect(stream.length).toBe(410);
    expect(sha256(stream)).toBe('bd104bb9019e5a5e5a3b1cbde5a01981d392d70028abac4141b3e52aa009c845');
    expect(toHex(stream)).toBe(toHex(hex(MEDIUM_STREAM)));
    expectRoundTrip(input, hex(MEDIUM_STREAM));
  });

  test('ENC7 medium vector with the run-length stage', () => {
    const input = mediumInput();
    const payload = rlePack(input);
    expect(payload.length).toBe(3001);
    expect(payload[0]).toBe(0x03);
    const stream = lzcompEncode(input, { runLength: true });
    expect(stream.length).toBe(411);
    expect(sha256(stream)).toBe('a5b0d75fd79dfd8c7eff9352bf34d6ce2a4b6f91fb65968423843573115c814c');
    expectRoundTrip(input, stream);
  });

  test('ENC8 long runs with the run-length stage', () => {
    const input = new Uint8Array(1000);
    for (let i = 0; i < input.length; i++) input[i] = i % 300 < 200 ? 0 : i % 256;
    expect(sha256(input)).toBe('84bb4fe846a4df57e67d08898de4e6c01f37b39bab1ebe2fe0eec4ecaeb41c95');
    expect(rlePack(input).length).toBe(313);
    const stream = lzcompEncode(input, { runLength: true });
    expect(stream.length).toBe(219);
    expect(sha256(stream)).toBe('975800dd61937f276be4556ce71fe225ec4479bb7861ebd643b9cee96ff1d7a5');
    expectRoundTrip(input, stream);
  });
});

// --- 10.8 rlePack --------------------------------------------------------------------------

describe('rlePack (section 10.8)', () => {
  test('vectors', () => {
    expect(toHex(rlePack(new Uint8Array(0)))).toBe('00');
    expect(toHex(rlePack(hex('05')))).toBe('00 05');
    expect(toHex(rlePack(hex('41 41 41 41 41 00 42 43 43 43')))).toBe('01 01 05 41 00 42 43 43 43');
    expect(toHex(rlePack(hex('00 01 02 03 00 00 00 00 00 00')))).toBe('04 00 01 02 03 04 06 00');
    expect(toHex(rlePack(new Uint8Array(300).fill(9)))).toBe('00 00 ff 09 00 2d 09');
  });
});
