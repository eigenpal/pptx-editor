// CTF 'glyf' reconstruction for MicroType Express.
// Implements W3C Member Submission "MicroType Express (MTX) Font Format" (2008):
//   5.1  'loca' table (rebuilt from the sequential CTF glyph stream)
//   5.6-5.11 CTF glyph records, simple/composite glyph tables, triplet coordinate encoding
//   6.1  255USHORT / 255SHORT data types
//   6.2  push data with Hop3Code / Hop4Code, re-emitted as TrueType PUSH instructions
//   7    glyph instruction stream
// and the OpenType 'glyf'/'loca' table formats for the output.

import { ByteReader, writeU16, writeU32 } from './sfnt.ts';
import { FontFormatError } from './types.ts';

const ERR = 'ctf-corrupt' as const;

// --- 6.1 255USHORT / 255SHORT --------------------------------------------------------------

export function read255UShort(r: ByteReader): number {
  const code = r.u8();
  if (code === 253) return r.u16();
  if (code === 255) return 253 + r.u8();
  if (code === 254) return 506 + r.u8();
  return code;
}

export function read255Short(r: ByteReader): number {
  let code = r.u8();
  if (code === 253) return r.i16();
  let sign = 1;
  if (code === 250) {
    sign = -1;
    code = r.u8();
  }
  let value: number;
  if (code === 255) value = 250 + r.u8();
  else if (code === 254) value = 500 + r.u8();
  else value = code;
  return value * sign;
}

const HOP3 = 251;
const HOP4 = 252;

/** 6.2.4: decodes `count` push values, expanding Hop3/Hop4 codes. */
export function readPushData(r: ByteReader, count: number): Int32Array {
  // Every value needs at least one input byte except Hop expansions (3 or 5 values from
  // 2 or 3 bytes); bound the allocation by the input before trusting the count.
  if (count > r.remaining * 3) throw new FontFormatError(ERR, 'push count exceeds push stream');
  const out = new Int32Array(count);
  let i = 0;
  while (i < count) {
    if (r.remaining <= 0) throw new FontFormatError(ERR, 'push stream truncated');
    const code = r.data[r.pos];
    if (code === HOP3 || code === HOP4) {
      const n = code === HOP3 ? 3 : 5;
      if (i < 2 || i + n > count) throw new FontFormatError(ERR, 'invalid hop code position');
      const a = out[i - 2];
      r.pos++;
      out[i++] = a;
      out[i++] = read255Short(r);
      out[i++] = a;
      if (n === 5) {
        out[i++] = read255Short(r);
        out[i++] = a;
      }
    } else {
      out[i++] = read255Short(r);
    }
  }
  return out;
}

/** Re-encodes initial stack values as TrueType PUSHB/PUSHW/NPUSHB/NPUSHW instructions. */
export function encodePushInstructions(values: Int32Array): Uint8Array {
  const out: number[] = [];
  let i = 0;
  while (i < values.length) {
    const isByte = (v: number): boolean => v >= 0 && v <= 255;
    const byteRun = isByte(values[i]);
    let j = i;
    while (j < values.length && j - i < 255 && isByte(values[j]) === byteRun) j++;
    const n = j - i;
    if (byteRun) {
      if (n <= 8) out.push(0xb0 + n - 1);
      else out.push(0x40, n);
      for (let k = i; k < j; k++) out.push(values[k]);
    } else {
      if (n <= 8) out.push(0xb8 + n - 1);
      else out.push(0x41, n);
      for (let k = i; k < j; k++) out.push((values[k] >> 8) & 0xff, values[k] & 0xff);
    }
    i = j;
  }
  return Uint8Array.from(out);
}

// --- 5.11 Triplet encoding ---------------------------------------------------------------

interface Triplet {
  byteCount: number;
  xBits: number;
  yBits: number;
  deltaX: number;
  deltaY: number;
  xSign: number;
  ySign: number;
}

/** Builds the 128-entry table from section 5.11 (signs are -,+ alternating, x before y). */
function buildTripletTable(): Triplet[] {
  const t: Triplet[] = [];
  const signs = [-1, 1];
  // 0..9: dx = 0, 8-bit dy with deltaY 0,256,512,768,1024
  for (let i = 0; i < 10; i++) {
    t.push({
      byteCount: 2,
      xBits: 0,
      yBits: 8,
      deltaX: 0,
      deltaY: 256 * (i >> 1),
      xSign: 1,
      ySign: signs[i & 1],
    });
  }
  // 10..19: dy = 0, 8-bit dx
  for (let i = 0; i < 10; i++) {
    t.push({
      byteCount: 2,
      xBits: 8,
      yBits: 0,
      deltaX: 256 * (i >> 1),
      deltaY: 0,
      xSign: signs[i & 1],
      ySign: 1,
    });
  }
  // 20..83: 4/4 bits, deltas 1,17,33,49
  for (let dx = 0; dx < 4; dx++) {
    for (let dy = 0; dy < 4; dy++) {
      for (let s = 0; s < 4; s++) {
        t.push({
          byteCount: 2,
          xBits: 4,
          yBits: 4,
          deltaX: 1 + 16 * dx,
          deltaY: 1 + 16 * dy,
          xSign: signs[s & 1],
          ySign: signs[s >> 1],
        });
      }
    }
  }
  // 84..119: 8/8 bits, deltas 1,257,513
  for (let dx = 0; dx < 3; dx++) {
    for (let dy = 0; dy < 3; dy++) {
      for (let s = 0; s < 4; s++) {
        t.push({
          byteCount: 3,
          xBits: 8,
          yBits: 8,
          deltaX: 1 + 256 * dx,
          deltaY: 1 + 256 * dy,
          xSign: signs[s & 1],
          ySign: signs[s >> 1],
        });
      }
    }
  }
  // 120..123: 12/12 bits; 124..127: 16/16 bits
  for (const [byteCount, nb] of [
    [4, 12],
    [5, 16],
  ] as const) {
    for (let s = 0; s < 4; s++) {
      t.push({
        byteCount,
        xBits: nb,
        yBits: nb,
        deltaX: 0,
        deltaY: 0,
        xSign: signs[s & 1],
        ySign: signs[s >> 1],
      });
    }
  }
  return t;
}

export const TRIPLETS: readonly Triplet[] = buildTripletTable();

// --- Output buffer -----------------------------------------------------------------------

class GrowBuffer {
  buf: Uint8Array;
  length = 0;
  private readonly cap: number;

  constructor(initial: number, cap: number) {
    this.cap = cap;
    this.buf = new Uint8Array(Math.max(64, Math.min(initial, cap)));
  }

  reserve(extra: number): void {
    const need = this.length + extra;
    if (need > this.cap) throw new FontFormatError('limit-exceeded', 'glyf output exceeds limit');
    if (need <= this.buf.length) return;
    let size = this.buf.length;
    while (size < need) size = Math.min(this.cap, size * 2);
    const next = new Uint8Array(size);
    next.set(this.buf.subarray(0, this.length));
    this.buf = next;
  }

  u8(v: number): void {
    this.reserve(1);
    this.buf[this.length++] = v & 0xff;
  }

  u16(v: number): void {
    this.reserve(2);
    writeU16(this.buf, this.length, v & 0xffff);
    this.length += 2;
  }

  bytes(b: Uint8Array): void {
    this.reserve(b.length);
    this.buf.set(b, this.length);
    this.length += b.length;
  }

  align(n: number): void {
    while (this.length % n !== 0) this.u8(0);
  }
}

// --- Glyph decoding ----------------------------------------------------------------------

const toI16 = (v: number): number => (v << 16) >> 16;

const ARG_1_AND_2_ARE_WORDS = 0x0001;
const WE_HAVE_A_SCALE = 0x0008;
const MORE_COMPONENTS = 0x0020;
const WE_HAVE_AN_X_AND_Y_SCALE = 0x0040;
const WE_HAVE_A_TWO_BY_TWO = 0x0080;
const WE_HAVE_INSTRUCTIONS = 0x0100;

interface Streams {
  glyf: ByteReader;
  push: ByteReader;
  code: ByteReader;
}

function readInstructions(s: Streams): Uint8Array {
  const pushCount = read255UShort(s.glyf);
  const codeSize = read255UShort(s.glyf);
  const push = encodePushInstructions(readPushData(s.push, pushCount));
  const code = s.code.bytes(codeSize);
  const total = push.length + code.length;
  if (total > 0xffff) throw new FontFormatError(ERR, 'glyph instructions too long');
  const out = new Uint8Array(total);
  out.set(push, 0);
  out.set(code, push.length);
  return out;
}

function writeSimpleGlyph(
  s: Streams,
  out: GrowBuffer,
  numContours: number,
  explicitBox: number[] | null
): void {
  const g = s.glyf;
  const endPts = new Array<number>(numContours);
  let last = -1;
  for (let c = 0; c < numContours; c++) {
    const v = read255UShort(g);
    last = c === 0 ? v : last + v;
    if (last > 0xffff) throw new FontFormatError(ERR, 'contour end point overflow');
    endPts[c] = last;
  }
  const numPoints = numContours > 0 ? last + 1 : 0;
  // Each point needs one flag byte and at least one coordinate byte.
  if (numPoints * 2 > g.remaining) throw new FontFormatError(ERR, 'glyph point data truncated');
  const flags = g.bytes(numPoints);
  const xs = new Int32Array(numPoints);
  const ys = new Int32Array(numPoints);
  const onCurve = new Uint8Array(numPoints);
  let x = 0;
  let y = 0;
  let xMin = 0x7fff;
  let yMin = 0x7fff;
  let xMax = -0x8000;
  let yMax = -0x8000;
  for (let i = 0; i < numPoints; i++) {
    const flag = flags[i];
    onCurve[i] = (flag & 0x80) === 0 ? 1 : 0;
    const t = TRIPLETS[flag & 0x7f];
    const n = t.byteCount - 1;
    let data = 0;
    for (let j = 0; j < n; j++) data = data * 256 + g.u8();
    const totalBits = n * 8;
    const dxRaw = t.xBits ? Math.floor(data / 2 ** (totalBits - t.xBits)) % 2 ** t.xBits : 0;
    const dyRaw = t.yBits
      ? Math.floor(data / 2 ** (totalBits - t.xBits - t.yBits)) % 2 ** t.yBits
      : 0;
    const dx = t.xSign * (dxRaw + t.deltaX);
    const dy = t.ySign * (dyRaw + t.deltaY);
    x = toI16(x + dx);
    y = toI16(y + dy);
    xs[i] = x;
    ys[i] = y;
    if (x < xMin) xMin = x;
    if (x > xMax) xMax = x;
    if (y < yMin) yMin = y;
    if (y > yMax) yMax = y;
  }
  if (numPoints === 0) xMin = yMin = xMax = yMax = 0;
  const instructions = readInstructions(s);
  const box = explicitBox ?? [xMin, yMin, xMax, yMax];

  out.u16(numContours);
  for (const v of box) out.u16(v);
  for (const e of endPts) out.u16(e);
  out.u16(instructions.length);
  out.bytes(instructions);

  // TrueType flags: ON_CURVE 1, X_SHORT 2, Y_SHORT 4, REPEAT 8, X_SAME/POS 16, Y_SAME/POS 32.
  const ttFlags = new Uint8Array(numPoints);
  const xBytes: number[] = [];
  const yBytes: number[] = [];
  let px = 0;
  let py = 0;
  for (let i = 0; i < numPoints; i++) {
    let f = onCurve[i];
    const dx = toI16(xs[i] - px);
    const dy = toI16(ys[i] - py);
    px = xs[i];
    py = ys[i];
    if (dx === 0) f |= 0x10;
    else if (dx >= -255 && dx <= 255) {
      f |= 0x02 | (dx > 0 ? 0x10 : 0);
      xBytes.push(Math.abs(dx));
    } else xBytes.push((dx >> 8) & 0xff, dx & 0xff);
    if (dy === 0) f |= 0x20;
    else if (dy >= -255 && dy <= 255) {
      f |= 0x04 | (dy > 0 ? 0x20 : 0);
      yBytes.push(Math.abs(dy));
    } else yBytes.push((dy >> 8) & 0xff, dy & 0xff);
    ttFlags[i] = f;
  }
  for (let i = 0; i < numPoints;) {
    const f = ttFlags[i];
    let run = 1;
    while (i + run < numPoints && ttFlags[i + run] === f && run < 256) run++;
    if (run > 1) {
      out.u8(f | 0x08);
      out.u8(run - 1);
    } else {
      out.u8(f);
    }
    i += run;
  }
  out.bytes(Uint8Array.from(xBytes));
  out.bytes(Uint8Array.from(yBytes));
}

function writeCompositeGlyph(s: Streams, out: GrowBuffer): void {
  const g = s.glyf;
  const box = [g.i16(), g.i16(), g.i16(), g.i16()];
  out.u16(-1);
  for (const v of box) out.u16(v);
  let flags: number;
  let haveInstructions = false;
  do {
    flags = g.u16();
    const glyphIndex = g.u16();
    let argBytes = flags & ARG_1_AND_2_ARE_WORDS ? 4 : 2;
    if (flags & WE_HAVE_A_SCALE) argBytes += 2;
    else if (flags & WE_HAVE_AN_X_AND_Y_SCALE) argBytes += 4;
    else if (flags & WE_HAVE_A_TWO_BY_TWO) argBytes += 8;
    if (flags & WE_HAVE_INSTRUCTIONS) haveInstructions = true;
    out.u16(flags);
    out.u16(glyphIndex);
    out.bytes(g.bytes(argBytes));
  } while (flags & MORE_COMPONENTS);
  if (haveInstructions) {
    const instructions = readInstructions(s);
    out.u16(instructions.length);
    out.bytes(instructions);
  }
}

export interface GlyfResult {
  glyf: Uint8Array;
  loca: Uint8Array;
  indexToLocFormat: 0 | 1;
  /** Bytes left unread in the three CTF streams (0 for a cleanly consumed font). */
  leftover: { glyf: number; push: number; code: number };
}

export interface GlyfInput {
  glyf: Uint8Array;
  push: Uint8Array;
  code: Uint8Array;
  numGlyphs: number;
  indexToLocFormat: number;
  maxOutputBytes: number;
}

/** Decodes the sequential CTF glyph stream into TrueType 'glyf' + 'loca'. */
export function decodeCtfGlyf(input: GlyfInput): GlyfResult {
  const s: Streams = {
    glyf: new ByteReader(input.glyf, 0, ERR),
    push: new ByteReader(input.push, 0, ERR),
    code: new ByteReader(input.code, 0, ERR),
  };
  // Every CTF glyph record needs at least two bytes (numContours).
  if (input.numGlyphs * 2 > input.glyf.length) {
    throw new FontFormatError(ERR, 'glyf stream shorter than maxp.numGlyphs');
  }
  const out = new GrowBuffer(input.glyf.length * 2, input.maxOutputBytes);
  const offsets = new Uint32Array(input.numGlyphs + 1);
  for (let gid = 0; gid < input.numGlyphs; gid++) {
    offsets[gid] = out.length;
    let numContours = s.glyf.i16();
    if (numContours < 0) {
      if (numContours !== -1) throw new FontFormatError(ERR, `glyph ${gid}: bad contour count`);
      writeCompositeGlyph(s, out);
    } else if (numContours === 0x7fff) {
      numContours = s.glyf.i16();
      if (numContours < 0) throw new FontFormatError(ERR, `glyph ${gid}: bad contour count`);
      const box = [s.glyf.i16(), s.glyf.i16(), s.glyf.i16(), s.glyf.i16()];
      writeSimpleGlyph(s, out, numContours, box);
    } else if (numContours > 0) {
      writeSimpleGlyph(s, out, numContours, null);
    }
    // numContours === 0: empty glyph, zero-length entry.
    out.align(4);
  }
  offsets[input.numGlyphs] = out.length;

  let format: 0 | 1 = input.indexToLocFormat === 0 ? 0 : 1;
  if (format === 0 && out.length > 0x1fffe) format = 1;
  const loca = new Uint8Array((input.numGlyphs + 1) * (format === 0 ? 2 : 4));
  for (let i = 0; i <= input.numGlyphs; i++) {
    if (format === 0) writeU16(loca, i * 2, offsets[i] / 2);
    else writeU32(loca, i * 4, offsets[i]);
  }
  return {
    glyf: out.buf.slice(0, out.length),
    loca,
    indexToLocFormat: format,
    leftover: { glyf: s.glyf.remaining, push: s.push.remaining, code: s.code.remaining },
  };
}
