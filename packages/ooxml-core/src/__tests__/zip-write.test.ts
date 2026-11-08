// Tests for the stepped ZIP writer (`zip-write.ts`). The first block is adapted from
// docx-editor/packages/core/src/store/package/__tests__/zip-write.test.ts at 7cd3776a (EigenPal), with WordprocessingML
// part names replaced by PresentationML ones. The second block reads every saved archive with
// an independent reader: `node:zlib` inflates each entry and `Bun.hash.crc32` checks each CRC,
// because fflate's own reader does not check CRCs.

import { describe, expect, test } from 'bun:test';
import { inflateRawSync } from 'node:zlib';
import { deflateSync, strToU8, unzipSync, zipSync } from 'fflate';
import {
  DEFAULT_ZIP_LIMITS,
  ZIP_RATIO_EXEMPT_BYTES,
  readZip,
  writeZip,
  writeZipSteps,
} from '../zip.ts';
import { ZIP_WRITE_SLICE_BYTES, writeZipEntriesSteps } from '../zip-write.ts';

const MTIME = new Date(Date.UTC(1980, 0, 2, 12));

function drain<T>(steps: Generator<void, T>): { value: T; yields: number } {
  let yields = 0;
  for (;;) {
    const next = steps.next();
    if (next.done) return { value: next.value, yields };
    yields += 1;
  }
}

/** Compressible bytes longer than one slice, with enough variety to need real deflating. */
function longEntry(length: number): Uint8Array {
  const bytes = new Uint8Array(length);
  for (let index = 0; index < length; index += 1) bytes[index] = (index * 31 + (index >> 7)) & 0x7f;
  return bytes;
}

describe('ZIP writer', () => {
  test('writes the zipSync layout byte for byte', () => {
    const entries = new Map<string, Uint8Array>([
      ['[Content_Types].xml', strToU8('<Types/>')],
      ['ppt/presentation.xml', strToU8('<p:presentation/>'.repeat(50))],
      ['ppt/média/émpty.bin', new Uint8Array(0)],
    ]);
    const { value } = drain(
      writeZipEntriesSteps(entries, MTIME, function* (data) {
        return deflateSync(data);
      })
    );
    expect([...value]).toEqual([...zipSync(Object.fromEntries(entries), { mtime: MTIME })]);
  });

  test('deflates a long entry across steps and reads back unchanged', () => {
    const long = longEntry(ZIP_WRITE_SLICE_BYTES * 4 + 123);
    const entries = new Map<string, Uint8Array>([
      ['/ppt/presentation.xml', long],
      ['/[Content_Types].xml', strToU8('<Types/>')],
    ]);
    const { value, yields } = drain(writeZipSteps(entries));
    expect(yields).toBeGreaterThanOrEqual(8);
    expect([...value]).toEqual([...writeZip(entries)]);

    const independent = unzipSync(value);
    expect([...independent['ppt/presentation.xml']!]).toEqual([...long]);
    const bounded = readZip(value);
    if (!bounded.ok) throw new Error(bounded.reason);
    expect(bounded.entries.get('/ppt/presentation.xml')).toEqual(long);
  });

  test('refuses unsafe part names before writing anything', () => {
    expect(() => writeZipSteps(new Map([['/../evil.xml', new Uint8Array(1)]]))).toThrow(
      /unsafe part name/
    );
  });
});

interface Entry {
  readonly name: string;
  readonly method: number;
  readonly crc: number;
  readonly data: Uint8Array;
}

/** Every entry of `zip`, read from the central directory and checked against its local header. */
function readIndependently(zip: Uint8Array): Entry[] {
  const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
  const end = zip.length - 22;
  expect(view.getUint32(end, true)).toBe(0x06054b50);
  const count = view.getUint16(end + 10, true);
  let central = view.getUint32(end + 16, true);
  const entries: Entry[] = [];
  for (let i = 0; i < count; i += 1) {
    expect(view.getUint32(central, true)).toBe(0x02014b50);
    const flags = view.getUint16(central + 8, true);
    const method = view.getUint16(central + 10, true);
    const crc = view.getUint32(central + 16, true);
    const compressedSize = view.getUint32(central + 20, true);
    const size = view.getUint32(central + 24, true);
    const nameLength = view.getUint16(central + 28, true);
    const extraLength = view.getUint16(central + 30, true);
    const commentLength = view.getUint16(central + 32, true);
    const local = view.getUint32(central + 42, true);
    const nameBytes = zip.subarray(central + 46, central + 46 + nameLength);
    // The local header repeats the same fields, and no data descriptor follows the data.
    expect(view.getUint32(local, true)).toBe(0x04034b50);
    expect(flags & 0x08).toBe(0);
    for (const [at, width] of [
      [6, 2],
      [8, 2],
      [14, 4],
      [18, 4],
      [22, 4],
      [26, 2],
    ] as const) {
      const l = width === 2 ? view.getUint16(local + at, true) : view.getUint32(local + at, true);
      const c =
        width === 2
          ? view.getUint16(central + 2 + at, true)
          : view.getUint32(central + 2 + at, true);
      expect(l).toBe(c);
    }
    const start = local + 30 + nameLength + view.getUint16(local + 28, true);
    const raw = zip.subarray(start, start + compressedSize);
    const data = method === 0 ? raw : new Uint8Array(inflateRawSync(raw));
    expect(data.length).toBe(size);
    entries.push({ name: new TextDecoder().decode(nameBytes), method, crc, data });
    central += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

describe('a saved archive reads back with an independent reader', () => {
  const PNG = Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82]);

  test('names, methods, CRCs, and bytes of every entry, around the slice size', () => {
    const parts = new Map<string, Uint8Array>([
      ['/[Content_Types].xml', strToU8('<Types/>')],
      ['/ppt/empty.bin', new Uint8Array(0)],
      ['/ppt/one-slice.bin', longEntry(ZIP_WRITE_SLICE_BYTES)],
      ['/ppt/slice-and-one.bin', longEntry(ZIP_WRITE_SLICE_BYTES + 1)],
      ['/ppt/several-slices.xml', longEntry(ZIP_WRITE_SLICE_BYTES * 3 + 17)],
      ['/ppt/media/image 1.png', PNG],
      ['/ppt/média/a%41.bin', strToU8('x')],
    ]);
    const entries = readIndependently(writeZip(parts));
    expect(entries.map((e) => e.name)).toEqual([
      '[Content_Types].xml',
      'ppt/empty.bin',
      'ppt/one-slice.bin',
      'ppt/slice-and-one.bin',
      'ppt/several-slices.xml',
      'ppt/media/image%201.png',
      'ppt/média/a%2541.bin',
    ]);
    const originals = [...parts.values()];
    entries.forEach((entry, index) => {
      expect([...entry.data]).toEqual([...originals[index]!]);
      expect(entry.crc).toBe(Bun.hash.crc32(originals[index]!));
      expect(entry.method).toBe(entry.name.endsWith('.png') ? 0 : 8);
    });
  });
});

describe('a saved archive always reopens under the default limits', () => {
  test('a blank part that opened stored is stored again, not deflated past the ratio cap', () => {
    // 12 MB of one byte, stored: about 1:1, so it opens. Deflated it would be about 12 KB.
    const blank = new Uint8Array(12_000_000);
    const source = zipSync({
      '[Content_Types].xml': strToU8('<Types/>'),
      'ppt/media/image1.bmp': [blank, { level: 0 }],
      'ppt/slides/slide1.xml': strToU8('<p:sld/>'.repeat(100)),
    });
    const opened = readZip(source);
    if (!opened.ok) throw new Error(opened.reason);
    const saved = writeZip(opened.entries);
    const reopened = readZip(saved);
    expect(reopened.ok).toBe(true);
    const methods = readIndependently(saved).map((e) => [e.name, e.method]);
    expect(methods).toEqual([
      ['[Content_Types].xml', 8],
      ['ppt/media/image1.bmp', 0],
      ['ppt/slides/slide1.xml', 8],
    ]);
    // It deflates 12 MB twice; give a loaded machine room.
  }, 30_000);

  test('an archive under the cap is written exactly as before', () => {
    const parts = new Map([
      ['/[Content_Types].xml', strToU8('<Types/>')],
      ['/ppt/media/image1.bmp', new Uint8Array(4_000_000)],
    ]);
    const saved = writeZip(parts);
    expect(readIndependently(saved).map((e) => e.method)).toEqual([8, 8]);
    expect(readZip(saved).ok).toBe(true);
  });
});

describe('entries that share one byte array', () => {
  test('are written as separate entries, and only the one that must be is stored', () => {
    const blank = new Uint8Array(6_000_000);
    const parts = new Map([
      ['/[Content_Types].xml', strToU8('<Types/>')],
      ['/ppt/media/a.bin', blank],
      ['/ppt/media/b.bin', blank],
    ]);
    const saved = writeZip(parts);
    expect(readZip(saved).ok).toBe(true);
    const methods = readIndependently(saved).map((e) => e.method);
    expect(methods.filter((m) => m === 0)).toHaveLength(1);
    // A second save of what the first one reads back gives the same bytes.
    const again = readZip(saved);
    if (!again.ok) throw new Error(again.reason);
    expect([...writeZip(again.entries)]).toEqual([...saved]);
  }, 30_000);
});

describe('the writer stores enough to stay within the reader cap', () => {
  /** Bytes that deflate cannot shrink, from a fixed seed. */
  function noise(n: number): Uint8Array {
    const out = new Uint8Array(n);
    let state = 12345;
    for (let i = 0; i < n; i += 1) {
      state ^= state << 13;
      state ^= state >>> 17;
      state ^= state << 5;
      out[i] = state & 0xff;
    }
    return out;
  }

  test('an archive just past the cap when deflated is stored enough to reopen', () => {
    const parts = new Map([
      ['/a.bin', new Uint8Array(12_000_000)],
      ['/b.bin', noise(45_000)],
    ]);
    // Deflated as a whole (here by zipSync), this archive is past 200:1.
    const all = zipSync({ 'a.bin': parts.get('/a.bin')!, 'b.bin': parts.get('/b.bin')! });
    const ratio = 12_045_000 / all.length;
    expect(ratio).toBeGreaterThan(200);
    expect(ratio).toBeLessThan(230);
    expect(readZip(writeZip(parts)).ok).toBe(true);
  }, 60_000);
});

describe('the writer cap at its exact boundaries', () => {
  /** Bytes that deflate cannot shrink, from a fixed seed. */
  function noise(n: number): Uint8Array {
    const out = new Uint8Array(n);
    let state = 99991;
    for (let i = 0; i < n; i += 1) {
      state ^= state << 13;
      state ^= state >>> 17;
      state ^= state << 5;
      out[i] = state & 0xff;
    }
    return out;
  }

  /** The stored-entry count of `writeZip`, and whether deflating all would pass the cap. */
  function probe(parts: Map<string, Uint8Array>): { stored: number; pastCap: boolean } {
    const named = new Map([...parts].map(([name, data]) => [name.slice(1), data]));
    const all = drain(writeZipEntriesSteps(named, MTIME)).value;
    // Count each entry as the reader does: the larger of its compressed and declared size.
    const view = new DataView(all.buffer, all.byteOffset, all.byteLength);
    let at = view.getUint32(all.length - 22 + 16, true);
    let counted = 0;
    for (let i = view.getUint16(all.length - 22 + 10, true); i > 0; i -= 1) {
      counted += Math.max(view.getUint32(at + 20, true), view.getUint32(at + 24, true));
      at += 46 + view.getUint16(at + 28, true);
    }
    const pastCap =
      counted > ZIP_RATIO_EXEMPT_BYTES &&
      counted / all.length > (DEFAULT_ZIP_LIMITS.maxRatio ?? 200);
    const saved = writeZip(parts);
    expect(readZip(saved).ok).toBe(true);
    return { stored: readIndependently(saved).filter((e) => e.method === 0).length, pastCap };
  }

  test('an archive of exactly 8 MiB is exempt; one byte more is stored to stay readable', () => {
    expect(probe(new Map([['/a.bin', new Uint8Array(ZIP_RATIO_EXEMPT_BYTES)]])).stored).toBe(0);
    expect(probe(new Map([['/a.bin', new Uint8Array(ZIP_RATIO_EXEMPT_BYTES + 1)]])).stored).toBe(1);
  }, 60_000);

  test('at the 200:1 ratio, the writer stores exactly when deflating all would pass it', () => {
    // 9 MB of zeros and a little noise. Deflated whole (by this writer), these sizes give
    // about 200.09, 200.002, 199.998, and 199.91 to 1: two just past the cap, two within it.
    const results = [35_679, 35_699, 35_700, 35_720].map((extra) =>
      probe(
        new Map([
          ['/a.bin', new Uint8Array(9_000_000)],
          ['/b.bin', noise(extra)],
        ])
      )
    );
    expect(results.map((r) => r.pastCap)).toEqual([true, true, false, false]);
    expect(results.map((r) => r.stored > 0)).toEqual([true, true, false, false]);
  }, 120_000);
});
