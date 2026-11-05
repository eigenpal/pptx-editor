// Regression tests for audit finding S2: the CTF device-metrics tables
// (MTX Submission 5.4 hdmx, 5.5 VDMX) are bounded BEFORE they are decoded or allocated. LZCOMP
// can turn a few kilobytes of font data into a 16 MiB block, so the bound cannot rest on the
// size of the input.
//
// The inputs are built directly as decompressed CTF blocks, so these tests do not depend on the
// LZCOMP encoder.

import { describe, expect, test } from 'bun:test';
import { ctfToSfnt, decodeHdmx, decodeVdmx } from '../mtx-ctf.ts';
import { FontFormatError } from '../types.ts';

/** VDMX groups of 65,000 entries each, 3 bits per entry: the table from the audit. */
const RECS = 65_000;
const PERIOD = 6 + Math.ceil((RECS * 3) / 8);

function vdmxTable(groups: number): Uint8Array {
  const t = new Uint8Array(12 + groups * PERIOD);
  const view = new DataView(t.buffer);
  view.setUint16(2, groups);
  view.setUint16(4, 1);
  view.setUint16(10, 12); // offset of the first group
  for (let g = 0; g < groups; g += 1) view.setUint16(12 + g * PERIOD, RECS);
  return t;
}

const SIXTEEN_MIB = 16 * 1024 * 1024;

function reasonOf(run: () => unknown): string {
  try {
    run();
  } catch (e) {
    if (e instanceof FontFormatError) return `${e.reason}: ${e.message}`;
    throw e;
  }
  return 'ok';
}

describe('S2: VDMX', () => {
  test('a 16 MiB table is refused before its groups are decoded', () => {
    const table = vdmxTable(Math.floor((SIXTEEN_MIB - 12) / PERIOD));
    const t = performance.now();
    expect(reasonOf(() => decodeVdmx(table))).toBe('ctf-corrupt: VDMX too large');
    // Decoding first took 0.7 to 0.9 s and up to 1.6 GB.
    expect(performance.now() - t).toBeLessThan(100);
  });

  test('ctfToSfnt drops the table and keeps the font, quickly', () => {
    const dirLength = 12 + 3 * 16;
    const vdmxOffset = dirLength + 56 + 8;
    const block1 = new Uint8Array(SIXTEEN_MIB);
    const view = new DataView(block1.buffer);
    view.setUint32(0, 0x00010000);
    view.setUint16(4, 3);
    const record = (i: number, tag: string, offset: number, length: number): void => {
      for (let k = 0; k < 4; k += 1) block1[12 + i * 16 + k] = tag.charCodeAt(k);
      view.setUint32(12 + i * 16 + 8, offset);
      view.setUint32(12 + i * 16 + 12, length);
    };
    record(0, 'VDMX', vdmxOffset, SIXTEEN_MIB - vdmxOffset);
    record(1, 'head', dirLength, 54);
    record(2, 'maxp', dirLength + 56, 6);
    view.setUint32(dirLength + 12, 0x5f0f3cf5);
    view.setUint16(dirLength + 18, 1000);
    view.setUint32(dirLength + 56, 0x00005000);
    block1.set(vdmxTable(Math.floor((SIXTEEN_MIB - vdmxOffset - 12) / PERIOD)), vdmxOffset);
    const t = performance.now();
    const r = ctfToSfnt({
      block1,
      block2: new Uint8Array(0),
      block3: new Uint8Array(0),
      maxOutputBytes: 32 * 1024 * 1024,
      maxGlyphs: 65535,
    });
    expect(performance.now() - t).toBeLessThan(200);
    expect(r.warnings).toContain("dropped 'VDMX': VDMX too large");
  });
});

describe('S2: hdmx', () => {
  // 4 device records of 65,535 glyphs: 262,168 bytes decoded from 32 KB of zero bits.
  const numGlyphs = 65535;
  const records = 4;
  const metrics = { numGlyphs, unitsPerEm: 1000, advanceWidths: new Uint16Array(numGlyphs) };
  const hdmx = new Uint8Array(8 + records * 2 + Math.ceil((records * numGlyphs) / 8) + 8);
  const view = new DataView(hdmx.buffer);
  view.setUint16(2, records);
  for (let i = 0; i < records; i += 1) hdmx[8 + i * 2] = 10 + i;

  test('the decoded table counts against the output budget before it is allocated', () => {
    expect(reasonOf(() => decodeHdmx(hdmx, metrics, 100_000))).toBe(
      'limit-exceeded: hdmx over output limit'
    );
    expect(decodeHdmx(hdmx, metrics, 300_000).length).toBe(8 + records * 65540);
  });

  test('a stored (uncompressed) table counts too', () => {
    const stored = new Uint8Array(1000);
    new DataView(stored.buffer).setUint16(0, 0xffff - 1);
    expect(reasonOf(() => decodeHdmx(stored, metrics, 999))).toBe(
      'limit-exceeded: hdmx over output limit'
    );
  });
});
