// Compact Table Format (CTF) -> sfnt reconstruction for MicroType Express.
// Implements W3C Member Submission "MicroType Express (MTX) Font Format" (2008):
//   4    Compact Table Format overview (block 1 is an sfnt with transformed tables)
//   5.1  'loca' elided (offset/length zero in the CTF directory) and rebuilt
//   5.2  'cvt ' delta + 1-3 byte value encoding
//   5.3  magnitude dependent bit encoding (LSB first)
//   5.4  'hdmx' prediction/surprise encoding
//   5.5  'VDMX' prediction/surprise encoding
// Glyph data (5.6-7) is handled by mtx-glyf.ts.

import { decodeCtfGlyf } from './mtx-glyf.ts';
import {
  ByteReader,
  buildSfnt,
  readTableDirectory,
  readU16,
  writeU16,
  type SfntTable,
} from './sfnt.ts';
import { FontFormatError } from './types.ts';

const ERR = 'ctf-corrupt' as const;

// --- 5.2 cvt ---------------------------------------------------------------------------

const CVT_WORD_CODE = 238;
const CVT_NEG0 = 239;
const CVT_NEG8 = 247;
const CVT_POS1 = 248;
const CVT_LOWEST_CODE = 238;

export function decodeCvt(ctf: Uint8Array): Uint8Array {
  const r = new ByteReader(ctf, 0, ERR);
  const numEntries = r.u16();
  if (numEntries > r.remaining) throw new FontFormatError(ERR, 'cvt entry count exceeds data');
  const out = new Uint8Array(numEntries * 2);
  let prev = 0;
  for (let i = 0; i < numEntries; i++) {
    const code = r.u8();
    let value: number;
    if (code < CVT_LOWEST_CODE) value = code;
    else if (code >= CVT_POS1) value = r.u8() + (code - CVT_POS1 + 1) * CVT_LOWEST_CODE;
    else if (code >= CVT_NEG0 && code <= CVT_NEG8) {
      value = -(r.u8() + (code - CVT_NEG0) * CVT_LOWEST_CODE);
    } else {
      if (code !== CVT_WORD_CODE) throw new FontFormatError(ERR, 'bad cvt code');
      value = r.i16();
    }
    prev = (prev + value) & 0xffff;
    writeU16(out, i * 2, prev);
  }
  return out;
}

// --- 5.3 magnitude dependent encoding -------------------------------------------------

export class MagnitudeReader {
  private readonly data: Uint8Array;
  private bitPos: number;

  constructor(data: Uint8Array, byteOffset: number) {
    this.data = data;
    this.bitPos = byteOffset * 8;
  }

  private bit(): number {
    const byte = this.bitPos >> 3;
    if (byte >= this.data.length) throw new FontFormatError(ERR, 'magnitude stream truncated');
    const b = (this.data[byte] >> (this.bitPos & 7)) & 1;
    this.bitPos++;
    return b;
  }

  next(): number {
    let magnitude = 0;
    while (this.bit() === 1) magnitude++;
    if (magnitude === 0) return 0;
    return this.bit() === 1 ? -magnitude : magnitude;
  }

  /** Byte offset of the next whole byte (skips a partially consumed byte). */
  nextByte(): number {
    return (this.bitPos + 7) >> 3;
  }
}

// --- metrics helpers ------------------------------------------------------------------

interface Metrics {
  numGlyphs: number;
  unitsPerEm: number;
  advanceWidths: Uint16Array;
}

/** The largest VDMX table: its group offsets are 16-bit. */
const VDMX_MAX_BYTES = 0xffff;

function readAdvanceWidths(hhea: Uint8Array | undefined, hmtx: Uint8Array | undefined, n: number) {
  if (!hhea || hhea.length < 36 || !hmtx) throw new FontFormatError(ERR, 'hhea/hmtx missing');
  const numberOfHMetrics = readU16(hhea, 34);
  if (numberOfHMetrics === 0 || numberOfHMetrics * 4 > hmtx.length) {
    throw new FontFormatError(ERR, 'hmtx shorter than numberOfHMetrics');
  }
  const out = new Uint16Array(n);
  let last = 0;
  for (let g = 0; g < n; g++) {
    if (g < numberOfHMetrics) last = readU16(hmtx, g * 4);
    out[g] = last;
  }
  return out;
}

// --- 5.4 hdmx -------------------------------------------------------------------------

/**
 * Decode `hdmx`. `maxOutputBytes` is what is left of the font's output budget: the decoded
 * table is counted against it BEFORE it is allocated, because the device records can be about
 * 67 MB (1,020 records of 65,535 glyphs) from a few kilobytes of LZCOMP input.
 */
export function decodeHdmx(
  ctf: Uint8Array,
  m: Metrics,
  maxOutputBytes: number = Number.MAX_SAFE_INTEGER
): Uint8Array {
  const r = new ByteReader(ctf, 0, ERR);
  const version = r.u16();
  if (version >= 0x8000) {
    // Stored uncompressed: version = 0xFFFF - real_version.
    if (ctf.length > maxOutputBytes)
      throw new FontFormatError('limit-exceeded', 'hdmx over output limit');
    const copy = ctf.slice();
    writeU16(copy, 0, 0xffff - version);
    return copy;
  }
  const numRecords = r.u16();
  r.u32(); // sizeDeviceRecord (recomputed below)
  if (numRecords > 255 * 4) throw new FontFormatError(ERR, 'implausible hdmx record count');
  r.need(numRecords * 2);
  // Every width costs at least one bit of surprise data.
  if (numRecords * m.numGlyphs > (r.remaining - numRecords * 2) * 8) {
    throw new FontFormatError(ERR, 'hdmx data truncated');
  }
  const ppems: number[] = [];
  const maxWidths: number[] = [];
  for (let i = 0; i < numRecords; i++) {
    ppems.push(r.u8());
    maxWidths.push(r.u8());
  }
  const recordSize = (2 + m.numGlyphs + 3) & ~3;
  if (8 + recordSize * numRecords > maxOutputBytes)
    throw new FontFormatError('limit-exceeded', 'hdmx over output limit');
  const out = new Uint8Array(8 + recordSize * numRecords);
  writeU16(out, 0, version);
  writeU16(out, 2, numRecords);
  out[4] = (recordSize >>> 24) & 0xff;
  out[5] = (recordSize >>> 16) & 0xff;
  out[6] = (recordSize >>> 8) & 0xff;
  out[7] = recordSize & 0xff;
  const bits = new MagnitudeReader(ctf, r.pos);
  const upem = m.unitsPerEm;
  for (let rec = 0; rec < numRecords; rec++) {
    const base = 8 + rec * recordSize;
    const ppem = ppems[rec];
    out[base] = ppem;
    out[base + 1] = maxWidths[rec];
    for (let g = 0; g < m.numGlyphs; g++) {
      const scaled = Math.floor((64 * ppem * m.advanceWidths[g] + Math.floor(upem / 2)) / upem);
      const predicted = Math.floor((scaled + 32) / 64);
      const width = predicted + bits.next();
      if (width < 0 || width > 255) throw new FontFormatError(ERR, 'hdmx width out of range');
      out[base + 2 + g] = width;
    }
  }
  return out;
}

// --- 5.5 VDMX -------------------------------------------------------------------------

export function decodeVdmx(ctf: Uint8Array): Uint8Array {
  const r = new ByteReader(ctf, 0, ERR);
  const version = r.u16();
  if (version >= 0x8000) {
    const copy = ctf.slice();
    writeU16(copy, 0, 0xffff - version);
    return copy;
  }
  const numRecs = r.u16();
  const numRatios = r.u16();
  r.need(numRatios * 6);
  const ratios = r.bytes(numRatios * 4);
  const ctfOffsets: number[] = [];
  for (let i = 0; i < numRatios; i++) ctfOffsets.push(r.u16());
  if (numRecs > 0 && numRatios === 0) throw new FontFormatError(ERR, 'VDMX groups without ratios');

  interface Entry {
    ppem: number;
    yMax: number;
    yMin: number;
  }
  const groups: Entry[][] = [];
  // The decoded table cannot exceed 64 KiB, so check the running size BEFORE each group is
  // decoded. Checking only at the end let a 3.4 KB font decode 44.7 million entries first.
  let size = 6 + numRatios * 6;
  if (size > VDMX_MAX_BYTES) throw new FontFormatError(ERR, 'VDMX too large');
  let pos = numRatios > 0 ? ctfOffsets[0] : r.pos;
  for (let gi = 0; gi < numRecs; gi++) {
    const g = new ByteReader(ctf, pos, ERR);
    const recs = g.u16();
    const yMaxMult = g.i16();
    const yMinMult = g.i16();
    size += 4 + 6 * recs;
    if (size > VDMX_MAX_BYTES) throw new FontFormatError(ERR, 'VDMX too large');
    // Each entry needs at least 3 bits.
    if (recs * 3 > g.remaining * 8) throw new FontFormatError(ERR, 'VDMX group truncated');
    const bits = new MagnitudeReader(ctf, g.pos);
    const entries: Entry[] = [];
    let predictedPpem = 8;
    for (let j = 0; j < recs; j++) {
      const ppem = predictedPpem + bits.next();
      const yMaxErr = bits.next();
      const yMinErr = bits.next();
      const predictedYMax = Math.trunc((ppem * yMaxMult + 1024) / 2048);
      const predictedYMin = -Math.trunc((ppem * yMinMult + 1024) / 2048);
      const e = { ppem, yMax: predictedYMax + yMaxErr, yMin: predictedYMin + yMinErr };
      if (e.ppem < 0 || e.ppem > 0xffff) throw new FontFormatError(ERR, 'VDMX ppem out of range');
      entries.push(e);
      predictedPpem = ppem + 1;
    }
    groups.push(entries);
    pos = bits.nextByte();
  }

  // Lay out the TrueType table: header, ratios, offsets, then groups back to back.
  const groupOffsets: number[] = [];
  let offset = 6 + numRatios * 6;
  for (const g of groups) {
    groupOffsets.push(offset);
    offset += 4 + 6 * g.length;
  }
  const out = new Uint8Array(size);
  writeU16(out, 0, version);
  writeU16(out, 2, numRecs);
  writeU16(out, 4, numRatios);
  out.set(ratios, 6);
  // Ratios reference groups by offset; map the original offsets to group indices by rank.
  const distinct = [...new Set(ctfOffsets)].sort((a, b) => a - b);
  for (let i = 0; i < numRatios; i++) {
    const idx = Math.min(distinct.indexOf(ctfOffsets[i]), groups.length - 1);
    writeU16(out, 6 + numRatios * 4 + i * 2, idx >= 0 ? groupOffsets[idx] : 0);
  }
  groups.forEach((entries, gi) => {
    let p = groupOffsets[gi];
    writeU16(out, p, entries.length);
    out[p + 2] = entries.length ? entries[0].ppem & 0xff : 0;
    out[p + 3] = entries.length ? entries[entries.length - 1].ppem & 0xff : 0;
    p += 4;
    for (const e of entries) {
      writeU16(out, p, e.ppem);
      writeU16(out, p + 2, e.yMax & 0xffff);
      writeU16(out, p + 4, e.yMin & 0xffff);
      p += 6;
    }
  });
  return out;
}

// --- Assembly -------------------------------------------------------------------------

export interface CtfInput {
  block1: Uint8Array;
  block2: Uint8Array;
  block3: Uint8Array;
  maxOutputBytes: number;
  maxGlyphs: number;
}

export interface CtfResult {
  sfnt: Uint8Array;
  numGlyphs: number;
  warnings: string[];
}

/** Rebuilds a TrueType sfnt from the three decompressed CTF blocks. */
export function ctfToSfnt(input: CtfInput): CtfResult {
  const { block1 } = input;
  const dir = readTableDirectory(block1, 0, ERR);
  const warnings: string[] = [];
  const raw = new Map<string, Uint8Array>();
  for (const t of dir.tables) {
    if (t.tag === 'loca') continue;
    if (t.offset > block1.length || t.length > block1.length - t.offset) {
      throw new FontFormatError(ERR, `CTF table '${t.tag}' out of bounds`);
    }
    if (raw.has(t.tag)) throw new FontFormatError(ERR, `duplicate CTF table '${t.tag}'`);
    raw.set(t.tag, block1.subarray(t.offset, t.offset + t.length));
  }

  const head = raw.get('head');
  const maxp = raw.get('maxp');
  if (!head || head.length < 54) throw new FontFormatError(ERR, 'missing head table');
  if (!maxp || maxp.length < 6) throw new FontFormatError(ERR, 'missing maxp table');
  const numGlyphs = readU16(maxp, 4);
  if (numGlyphs > input.maxGlyphs) {
    throw new FontFormatError('limit-exceeded', `numGlyphs ${numGlyphs} over limit`);
  }
  const unitsPerEm = readU16(head, 18);
  const tables: SfntTable[] = [];
  const headCopy = head.slice();

  const glyfCtf = raw.get('glyf');
  if (glyfCtf) {
    const g = decodeCtfGlyf({
      glyf: glyfCtf,
      push: input.block2,
      code: input.block3,
      numGlyphs,
      indexToLocFormat: readU16(head, 50),
      maxOutputBytes: input.maxOutputBytes,
    });
    writeU16(headCopy, 50, g.indexToLocFormat);
    tables.push({ tag: 'glyf', data: g.glyf }, { tag: 'loca', data: g.loca });
    const { leftover } = g;
    if (leftover.glyf || leftover.push || leftover.code) {
      warnings.push(
        `unconsumed CTF bytes: glyf=${leftover.glyf} push=${leftover.push} code=${leftover.code}`
      );
    }
  }
  tables.push({ tag: 'head', data: headCopy });
  // Every decoded table counts against the output budget as it is built, not only the sfnt
  // that `buildSfnt` assembles at the end.
  let used = 0;
  for (const t of tables) used += t.data.length;

  for (const [tag, data] of raw) {
    if (tag === 'glyf' || tag === 'head') continue;
    if (tag === 'cvt ') {
      const cvt = decodeCvt(data);
      tables.push({ tag, data: cvt });
      used += cvt.length;
    } else if (tag === 'hdmx' || tag === 'VDMX') {
      try {
        const decoded =
          tag === 'hdmx'
            ? decodeHdmx(
                data,
                {
                  numGlyphs,
                  unitsPerEm: unitsPerEm || 1000,
                  advanceWidths: readAdvanceWidths(raw.get('hhea'), raw.get('hmtx'), numGlyphs),
                },
                input.maxOutputBytes - used
              )
            : decodeVdmx(data);
        tables.push({ tag, data: decoded });
        used += decoded.length;
      } catch (err) {
        // Device metrics are optional hints; drop the table rather than the whole font.
        if (!(err instanceof FontFormatError) || err.reason === 'limit-exceeded') throw err;
        warnings.push(`dropped '${tag}': ${err.message}`);
      }
    } else {
      tables.push({ tag, data });
      used += data.length;
    }
  }

  const sfnt = buildSfnt(dir.sfntVersion, tables, input.maxOutputBytes);
  return { sfnt, numGlyphs, warnings };
}
