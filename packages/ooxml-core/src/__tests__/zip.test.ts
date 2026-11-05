// Tests for the bounded ZIP reader at the trust boundary: the size and ratio caps from DOCX
// Editor (`zip.ts` at 7cd3776a), the layout and limit checks from PPTX Editor (S1, S6), the
// path rules, and OPC-equivalent duplicates.

import { describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import {
  DEFAULT_ZIP_LIMITS,
  ZIP_HARD_LIMITS,
  ZIP_RATIO_EXEMPT_BYTES,
  readZip,
  writeZip,
  writeZipSteps,
  type ZipLimits,
} from '../zip.ts';

/** Bytes that deflate cannot shrink, from a fixed seed. */
function incompressible(length: number, seed: number): Uint8Array {
  const out = new Uint8Array(length);
  let state = seed >>> 0 || 1;
  for (let i = 0; i < length; i += 1) {
    state ^= state << 13;
    state ^= state >>> 17;
    state ^= state << 5;
    out[i] = state & 0xff;
  }
  return out;
}

describe('size caps count what fflate will really allocate', () => {
  test('a deflated entry larger than its declared size counts its compressed size', () => {
    const files: Record<string, Uint8Array> = {};
    for (let i = 0; i < 8; i += 1) files[`f${i}.bin`] = incompressible(64 * 1024, i + 1);
    const zip = zipSync(files);
    const declared = 8 * 64 * 1024;
    // Each entry deflates to a little more than it declares. Counting only the declared size
    // would admit the archive at exactly its declared total.
    expect(readZip(zip, { maxEntries: 10, maxTotalBytes: declared })).toMatchObject({
      ok: false,
      reason: 'too-large',
      limit: 'zip.maxTotalBytes',
    });
    expect(readZip(zip, { maxEntries: 10, maxTotalBytes: declared + 4096 }).ok).toBe(true);
  });

  test('the ratio cap applies to the whole archive above 8 MiB', () => {
    expect(ZIP_RATIO_EXEMPT_BYTES).toBe(8 * 1024 * 1024);
    const below = zipSync({ 'a.bin': new Uint8Array(ZIP_RATIO_EXEMPT_BYTES - 1024) });
    expect(readZip(below).ok).toBe(true);
    // 9 MiB of zeros deflates to about 9 KiB: about 1000:1 for the whole archive.
    const above = zipSync({ 'a.bin': new Uint8Array(9 * 1024 * 1024) });
    expect(readZip(above)).toMatchObject({ ok: false, reason: 'too-large', limit: 'zip.maxRatio' });
  });

  test('a small, highly compressible part opens under a strict ratio', () => {
    const xml = strToU8('<a:p><a:r><a:t>x</a:t></a:r></a:p>'.repeat(4000));
    const zip = zipSync({ 'a.xml': xml });
    expect(xml.length / zip.length).toBeGreaterThan(200);
    expect(readZip(zip, { maxEntries: 10, maxTotalBytes: 1e9, maxRatio: 2 }).ok).toBe(true);
  });

  test('the entry count cap refuses before inflation', () => {
    const zip = zipSync({ 'a.xml': strToU8('<a/>'), 'b.xml': strToU8('<b/>') });
    expect(readZip(zip, { maxEntries: 1, maxTotalBytes: 1e6 })).toMatchObject({
      ok: false,
      reason: 'too-many-entries',
    });
  });

  test('the limits are frozen and the hard ceilings stay finite', () => {
    expect(Object.isFrozen(DEFAULT_ZIP_LIMITS)).toBe(true);
    expect(DEFAULT_ZIP_LIMITS.maxRatio).toBe(200);
    expect(Object.isFrozen(ZIP_HARD_LIMITS)).toBe(true);
    expect(Number.isFinite(ZIP_HARD_LIMITS.maxTotalBytes)).toBe(true);
    expect(
      readZip(zipSync({ 'a.bin': new Uint8Array(1) }), { maxEntries: Infinity } as never)
    ).toMatchObject({ ok: false, reason: 'invalid-limits' });
  });
});

/**
 * An archive with `records` central-directory records that all point at ONE stored local entry
 * of `payload` bytes, each declaring an uncompressed size of `declared`.
 */
function overlappingStoredZip(records: number, payload: number, declared: number): Uint8Array {
  const local = new Uint8Array(30 + 1 + payload);
  const lv = new DataView(local.buffer);
  lv.setUint32(0, 0x04034b50, true);
  lv.setUint32(18, payload, true);
  lv.setUint32(22, declared, true);
  lv.setUint16(26, 1, true);
  local[30] = 0x61;
  const directory: Uint8Array[] = [];
  for (let i = 0; i < records; i += 1) {
    const name = strToU8(`f${i}.bin`);
    const cd = new Uint8Array(46 + name.length);
    const cv = new DataView(cd.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint32(20, payload, true);
    cv.setUint32(24, declared, true);
    cv.setUint16(28, name.length, true);
    cd.set(name, 46);
    directory.push(cd);
  }
  const cdLength = directory.reduce((n, c) => n + c.length, 0);
  const out = new Uint8Array(local.length + cdLength + 22);
  out.set(local, 0);
  let p = local.length;
  for (const cd of directory) {
    out.set(cd, p);
    p += cd.length;
  }
  const ev = new DataView(out.buffer, p, 22);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, records, true);
  ev.setUint16(10, records, true);
  ev.setUint32(12, cdLength, true);
  ev.setUint32(16, local.length, true);
  return out;
}

describe('S1: entries must lie inside the input and must not share bytes', () => {
  test('directory records that share one stored entry are refused before inflation', () => {
    // 1 MiB of input that would otherwise inflate to 200 MiB under a 1 MiB cap.
    const zip = overlappingStoredZip(200, 1024 * 1024, 0);
    const r = readZip(zip, { maxEntries: 10_000, maxTotalBytes: 1024 * 1024 });
    expect(r.ok).toBe(false);
    expect(r.ok ? 'opened' : r.reason).toBe('overlapping-entries');
  });

  test('two records over one small stored entry are refused by the span check alone', () => {
    // 2 × 10 compressed bytes fit in the input, so only the overlapping spans reveal it.
    const r = readZip(overlappingStoredZip(2, 10, 10));
    expect(r.ok ? 'opened' : r.reason).toBe('overlapping-entries');
  });

  test('an entry whose data runs past the end of the input is refused', () => {
    const zip = zipSync({ 'a.bin': new Uint8Array(1000) }, { level: 0 });
    const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
    view.setUint32(zip.length - 22 - (46 + 5) + 20, 5000, true);
    expect(readZip(zip)).toMatchObject({ ok: false, reason: 'inflate-error' });
  });
});

describe('entry names follow the OPC path rules', () => {
  /** A one-entry archive whose entry has exactly `name`, written raw. */
  function named(name: string): Uint8Array {
    const zip = zipSync({ [`${'x'.repeat(name.length)}`]: strToU8('<a/>') });
    const encoded = strToU8(name);
    const out = zip.slice();
    // The placeholder name occurs once in the local header and once in the directory record.
    for (let p = 0; p + encoded.length <= out.length; p += 1)
      if (out.subarray(p, p + encoded.length).every((b) => b === 0x78)) {
        out.set(encoded, p);
        p += encoded.length;
      }
    return out;
  }

  test.each([
    ['../evil.xml'],
    ['a/../../evil.xml'],
    ['a\\b.xml'],
    ['C:/evil.xml'],
    ['//server/share.xml'],
    ['a%2fb.xml'],
    ['a%5Cb.xml'],
    ['a/%2e%2e/b.xml'],
    ['a/./b.xml'],
    ['a//b.xml'],
    ['a/b.'],
    ['a/__proto__'],
    ['bad%zz.xml'],
  ])('%s is refused before inflation', (name) => {
    expect(readZip(named(name))).toMatchObject({ ok: false, reason: 'bad-name' });
  });

  test('a directory entry is skipped, never inflated', () => {
    const r = readZip(zipSync({ 'dir/': new Uint8Array(0), 'dir/a.xml': strToU8('<a/>') }));
    expect(r.ok && [...r.entries.keys()]).toEqual(['/dir/a.xml']);
  });
});

describe('OPC-equivalent duplicates are refused', () => {
  test.each([
    ['ppt/slides/slide1.xml', 'PPT/Slides/Slide1.XML'],
    ['ppt/slides/slide1.xml', 'ppt/slides/%73lide1.xml'],
    ['a.xml', '/a.xml'],
  ])('%s and %s', (first, second) => {
    const zip = zipSync({ [first]: strToU8('<a/>'), [second]: strToU8('<b/>') });
    expect(readZip(zip)).toMatchObject({ ok: false, reason: 'bad-name' });
  });

  test('on write too', () => {
    const entries = new Map([
      ['/ppt/a.xml', strToU8('<a/>')],
      ['/PPT/A.xml', strToU8('<b/>')],
    ]);
    expect(() => writeZip(entries)).toThrow(/OPC-equivalent duplicate/);
  });

  test('a traversal or non-canonical name never reaches a written entry', () => {
    expect(() => writeZip(new Map([['/../evil.xml', new Uint8Array(1)]]))).toThrow(/unsafe/);
    expect(() => writeZip(new Map([['/a//b.xml', new Uint8Array(1)]]))).toThrow(/unsafe/);
    // A name without its leading slash is not canonical: refused, never normalized again.
    expect(() => writeZip(new Map([['a.xml', new Uint8Array(1)]]))).toThrow(/not-canonical/);
    expect(() => writeZipSteps(new Map([['/a/./b.xml', new Uint8Array(1)]])).next()).toThrow(
      /unsafe/
    );
  });
});

// Ported from docx-editor/packages/core/src/store/__tests__/package-open-tolerance.test.ts at
// 8f541b37 (EigenPal): the ratio cap at its exact boundaries.
describe('the ratio cap at its boundaries', () => {
  const generous: ZipLimits = { maxEntries: 100, maxTotalBytes: 256 * 1024 * 1024, maxRatio: 200 };

  test('an archive up to 8 MiB is exempt; one byte more is not', () => {
    const lone = (size: number) => zipSync({ 'pad.bin': new Uint8Array(size) }, { level: 9 });
    expect(readZip(lone(ZIP_RATIO_EXEMPT_BYTES), generous).ok).toBe(true);
    expect(readZip(lone(ZIP_RATIO_EXEMPT_BYTES + 1), generous)).toMatchObject({
      ok: false,
      limit: 'zip.maxRatio',
    });
  });

  test('an archive at exactly the ratio opens; one just below the ratio is refused', () => {
    const bytes = zipSync({ 'pad.bin': new Uint8Array(ZIP_RATIO_EXEMPT_BYTES + 1024) });
    const ratio = (ZIP_RATIO_EXEMPT_BYTES + 1024) / bytes.byteLength;
    expect(readZip(bytes, { ...generous, maxRatio: ratio }).ok).toBe(true);
    expect(readZip(bytes, { ...generous, maxRatio: ratio * (1 - 1e-9) })).toMatchObject({
      ok: false,
      limit: 'zip.maxRatio',
    });
  });
});
