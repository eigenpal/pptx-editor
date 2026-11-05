// sfnt (TrueType/OpenType) container helpers: big-endian reader, table-directory parsing,
// sfnt assembly with padding and checksums, and a structural validator.
// Implements the OpenType spec "Organization of an OpenType Font" (table directory,
// table records, CalcTableChecksum, head.checkSumAdjustment) plus the 'loca'/'maxp'
// consistency rules from the 'loca' and 'maxp' chapters.

import { FontFormatError, type FontRejectReason } from './types.ts';

/** Bounds-checked big-endian reader. Every overrun throws FontFormatError. */
export class ByteReader {
  readonly data: Uint8Array;
  pos: number;
  private readonly reason: FontRejectReason;

  constructor(data: Uint8Array, pos = 0, reason: FontRejectReason = 'truncated') {
    this.data = data;
    this.pos = pos;
    this.reason = reason;
  }

  get remaining(): number {
    return this.data.length - this.pos;
  }

  need(count: number): void {
    if (count < 0 || this.pos + count > this.data.length) {
      throw new FontFormatError(this.reason, `read of ${count} bytes at ${this.pos} overruns data`);
    }
  }

  u8(): number {
    this.need(1);
    return this.data[this.pos++];
  }

  u16(): number {
    this.need(2);
    const v = (this.data[this.pos] << 8) | this.data[this.pos + 1];
    this.pos += 2;
    return v;
  }

  i16(): number {
    const v = this.u16();
    return v >= 0x8000 ? v - 0x10000 : v;
  }

  u32(): number {
    this.need(4);
    const d = this.data;
    const p = this.pos;
    this.pos += 4;
    return ((d[p] << 24) >>> 0) + ((d[p + 1] << 16) | (d[p + 2] << 8) | d[p + 3]);
  }

  bytes(count: number): Uint8Array {
    this.need(count);
    const out = this.data.subarray(this.pos, this.pos + count);
    this.pos += count;
    return out;
  }
}

export function readU16(d: Uint8Array, p: number): number {
  return (d[p] << 8) | d[p + 1];
}

export function readU32(d: Uint8Array, p: number): number {
  return ((d[p] << 24) >>> 0) + ((d[p + 1] << 16) | (d[p + 2] << 8) | d[p + 3]);
}

export function writeU16(d: Uint8Array, p: number, v: number): void {
  d[p] = (v >>> 8) & 0xff;
  d[p + 1] = v & 0xff;
}

export function writeU32(d: Uint8Array, p: number, v: number): void {
  d[p] = (v >>> 24) & 0xff;
  d[p + 1] = (v >>> 16) & 0xff;
  d[p + 2] = (v >>> 8) & 0xff;
  d[p + 3] = v & 0xff;
}

export function tagToString(tag: number): string {
  return String.fromCharCode(
    (tag >>> 24) & 0xff,
    (tag >>> 16) & 0xff,
    (tag >>> 8) & 0xff,
    tag & 0xff
  );
}

export function tagToNumber(tag: string): number {
  let v = 0;
  for (let i = 0; i < 4; i++) v = v * 256 + (tag.charCodeAt(i) & 0xff);
  return v;
}

/** OpenType CalcTableChecksum over [offset, offset+length), zero-padded to 4 bytes. */
export function calcChecksum(data: Uint8Array, offset: number, length: number): number {
  let sum = 0;
  const end = offset + length;
  const whole = offset + (length & ~3);
  let p = offset;
  for (; p < whole; p += 4) sum = (sum + readU32(data, p)) >>> 0;
  if (p < end) {
    let last = 0;
    for (let i = 0; i < 4; i++) last = last * 256 + (p + i < end ? data[p + i] : 0);
    sum = (sum + last) >>> 0;
  }
  return sum;
}

export interface TableRecord {
  tag: string;
  checksum: number;
  offset: number;
  length: number;
}

export interface TableDirectory {
  sfntVersion: number;
  tables: TableRecord[];
}

/** Parses an sfnt table directory located at `base` within `data` (offsets relative to base). */
export function readTableDirectory(
  data: Uint8Array,
  base = 0,
  reason: FontRejectReason = 'invalid-sfnt'
): TableDirectory {
  const r = new ByteReader(data, base, reason);
  const sfntVersion = r.u32();
  const numTables = r.u16();
  r.u16();
  r.u16();
  r.u16();
  if (numTables === 0 || numTables > 512) {
    throw new FontFormatError(reason, `implausible table count ${numTables}`);
  }
  r.need(numTables * 16);
  const tables: TableRecord[] = [];
  for (let i = 0; i < numTables; i++) {
    const tag = tagToString(r.u32());
    const checksum = r.u32();
    const offset = r.u32();
    const length = r.u32();
    tables.push({ tag, checksum, offset, length });
  }
  return { sfntVersion, tables };
}

export interface SfntTable {
  tag: string;
  data: Uint8Array;
}

/**
 * Assembles an sfnt: directory sorted by tag, tables 4-byte aligned with zero padding,
 * per-table checksums, and head.checkSumAdjustment = 0xB1B0AFBA - sum(font).
 */
export function buildSfnt(sfntVersion: number, input: SfntTable[], maxBytes: number): Uint8Array {
  const tables = [...input].sort((a, b) => tagToNumber(a.tag) - tagToNumber(b.tag));
  for (let i = 1; i < tables.length; i++) {
    if (tables[i].tag === tables[i - 1].tag) {
      throw new FontFormatError('invalid-sfnt', `duplicate table '${tables[i].tag}'`);
    }
  }
  const numTables = tables.length;
  const headerSize = 12 + 16 * numTables;
  let total = headerSize;
  for (const t of tables) total += (t.data.length + 3) & ~3;
  if (total > maxBytes) throw new FontFormatError('limit-exceeded', `sfnt would be ${total} bytes`);

  const out = new Uint8Array(total);
  let entrySelector = 0;
  while (2 ** (entrySelector + 1) <= numTables) entrySelector++;
  const searchRange = 2 ** entrySelector * 16;
  writeU32(out, 0, sfntVersion);
  writeU16(out, 4, numTables);
  writeU16(out, 6, searchRange);
  writeU16(out, 8, entrySelector);
  writeU16(out, 10, numTables * 16 - searchRange);

  let offset = headerSize;
  let headOffset = -1;
  tables.forEach((t, i) => {
    out.set(t.data, offset);
    if (t.tag === 'head' && t.data.length >= 12) {
      headOffset = offset;
      writeU32(out, offset + 8, 0);
    }
    const rec = 12 + 16 * i;
    writeU32(out, rec, tagToNumber(t.tag));
    writeU32(out, rec + 4, calcChecksum(out, offset, t.data.length));
    writeU32(out, rec + 8, offset);
    writeU32(out, rec + 12, t.data.length);
    offset += (t.data.length + 3) & ~3;
  });
  if (headOffset >= 0) {
    const adjustment = (0xb1b0afba - calcChecksum(out, 0, total)) >>> 0;
    writeU32(out, headOffset + 8, adjustment);
  }
  return out;
}

/**
 * Reads an English name record (OpenType 'name' table) from an sfnt, preferring Windows
 * Unicode (3/1 or 3/10, language 0x409), then Unicode platform 0, then Mac Roman 1/0.
 * Returns '' when absent or malformed.
 */
export function readSfntName(font: Uint8Array, nameId: number): string {
  try {
    const dir = readTableDirectory(font, 0, 'invalid-sfnt');
    const rec = dir.tables.find((t) => t.tag === 'name');
    if (!rec || rec.offset + rec.length > font.length) return '';
    const name = font.subarray(rec.offset, rec.offset + rec.length);
    const r = new ByteReader(name, 0, 'invalid-sfnt');
    r.u16();
    const count = r.u16();
    const storage = r.u16();
    let best = -1;
    let bestScore = 0;
    let bestUtf16 = true;
    for (let i = 0; i < count; i++) {
      const platform = r.u16();
      const encoding = r.u16();
      const language = r.u16();
      const id = r.u16();
      r.u16(); // length (re-read below)
      r.u16(); // offset
      if (id !== nameId) continue;
      let score = 0;
      if (platform === 3 && (encoding === 1 || encoding === 10)) score = language === 0x409 ? 4 : 3;
      else if (platform === 0) score = 2;
      else if (platform === 1 && encoding === 0 && language === 0) score = 1;
      if (score > bestScore) {
        bestScore = score;
        best = 6 + i * 12;
        bestUtf16 = platform !== 1;
      }
    }
    if (best < 0) return '';
    const length = readU16(name, best + 8);
    const start = storage + readU16(name, best + 10);
    if (start + length > name.length) return '';
    let s = '';
    if (bestUtf16) {
      for (let i = 0; i + 1 < length; i += 2) s += String.fromCharCode(readU16(name, start + i));
    } else {
      for (let i = 0; i < length; i++) s += String.fromCharCode(name[start + i]);
    }
    return s;
  } catch {
    return '';
  }
}

export interface SfntValidation {
  ok: boolean;
  errors: string[];
  tables: string[];
  numGlyphs: number;
}

/**
 * Structural validator: directory sorted/unique, searchRange fields, offsets aligned and in
 * bounds, table checksums, whole-font checksum (head.checkSumAdjustment), maxp.numGlyphs vs
 * loca length, loca monotonic and within glyf.
 */
export function validateSfnt(font: Uint8Array): SfntValidation {
  const errors: string[] = [];
  const result = (numGlyphs = 0, tags: string[] = []): SfntValidation => ({
    ok: errors.length === 0,
    errors,
    tables: tags,
    numGlyphs,
  });
  if (font.length < 12) {
    errors.push('shorter than sfnt header');
    return result();
  }
  const version = readU32(font, 0);
  if (version !== 0x00010000 && version !== 0x74727565 && version !== 0x4f54544f) {
    errors.push(`unknown sfntVersion 0x${version.toString(16)}`);
  }
  const numTables = readU16(font, 4);
  if (numTables === 0) errors.push('no tables');
  if (12 + numTables * 16 > font.length) {
    errors.push('table directory overruns file');
    return result();
  }
  let entrySelector = 0;
  while (2 ** (entrySelector + 1) <= numTables) entrySelector++;
  const searchRange = 2 ** entrySelector * 16;
  if (readU16(font, 6) !== searchRange) errors.push('bad searchRange');
  if (readU16(font, 8) !== entrySelector) errors.push('bad entrySelector');
  if (readU16(font, 10) !== numTables * 16 - searchRange) errors.push('bad rangeShift');

  const byTag = new Map<string, TableRecord>();
  const tags: string[] = [];
  let prevTag = -1;
  for (let i = 0; i < numTables; i++) {
    const rec = 12 + 16 * i;
    const tagNum = readU32(font, rec);
    const tag = tagToString(tagNum);
    const checksum = readU32(font, rec + 4);
    const offset = readU32(font, rec + 8);
    const length = readU32(font, rec + 12);
    tags.push(tag);
    if (tagNum <= prevTag) errors.push(`table '${tag}' out of order or duplicated`);
    prevTag = tagNum;
    if (offset % 4 !== 0) errors.push(`table '${tag}' offset not 4-byte aligned`);
    if (offset + length > font.length) {
      errors.push(`table '${tag}' out of bounds`);
      continue;
    }
    let actual: number;
    if (tag === 'head' && length >= 12) {
      const copy = font.slice(offset, offset + length);
      writeU32(copy, 8, 0);
      actual = calcChecksum(copy, 0, length);
    } else {
      actual = calcChecksum(font, offset, length);
    }
    if (actual !== checksum) errors.push(`table '${tag}' checksum mismatch`);
    byTag.set(tag, { tag, checksum, offset, length });
  }

  const head = byTag.get('head');
  if (!head || head.length < 54) {
    errors.push('missing or short head table');
  } else {
    if (readU32(font, head.offset + 12) !== 0x5f0f3cf5) errors.push('bad head.magicNumber');
    if (calcChecksum(font, 0, font.length) !== 0xb1b0afba) {
      errors.push('whole-font checksum (head.checkSumAdjustment) mismatch');
    }
  }
  const maxp = byTag.get('maxp');
  let numGlyphs = 0;
  if (!maxp || maxp.length < 6) errors.push('missing or short maxp table');
  else numGlyphs = readU16(font, maxp.offset + 4);

  const loca = byTag.get('loca');
  const glyf = byTag.get('glyf');
  if (glyf || loca) {
    if (!glyf || !loca) errors.push('glyf and loca must both be present');
    else if (head && head.length >= 54 && maxp) {
      const longFormat = readU16(font, head.offset + 50) === 1;
      const entry = longFormat ? 4 : 2;
      if (loca.length !== (numGlyphs + 1) * entry) {
        errors.push(`loca length ${loca.length} != (numGlyphs+1)*${entry}`);
      } else {
        let prev = 0;
        for (let g = 0; g <= numGlyphs; g++) {
          const p = loca.offset + g * entry;
          const off = longFormat ? readU32(font, p) : readU16(font, p) * 2;
          if (off < prev) {
            errors.push(`loca not monotonic at glyph ${g}`);
            break;
          }
          prev = off;
        }
        if (prev > glyf.length) errors.push('loca points past end of glyf');
      }
    }
  }
  return result(numGlyphs, tags);
}
