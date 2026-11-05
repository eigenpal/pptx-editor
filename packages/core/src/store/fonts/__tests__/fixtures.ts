// Test-only builders for synthetic EOT/MTX/CTF data (no corpus bytes).
// The LZCOMP encoder is a simple greedy encoder that emits the symbol layout described in
// MTX Submission section 3.3 (literals, DUP2/4/6, copy items with length/distance trees);
// it is not size-optimal, only format-correct, which is all the decoder tests need.

import {
  AdaptiveHuffman,
  BitWriter,
  LZ_MAX_2BYTE_DIST,
  LZ_PRELOAD_SIZE,
  lzSymbolLayout,
  writePreload,
} from '../mtx-lzcomp.ts';
import { buildSfnt, type SfntTable } from '../sfnt.ts';

// --- byte helpers ------------------------------------------------------------------------

export class Bytes {
  private readonly out: number[] = [];

  u8(...v: number[]): this {
    for (const x of v) this.out.push(x & 0xff);
    return this;
  }

  u16(...v: number[]): this {
    for (const x of v) this.out.push((x >> 8) & 0xff, x & 0xff);
    return this;
  }

  u32(...v: number[]): this {
    for (const x of v)
      this.out.push((x >>> 24) & 0xff, (x >>> 16) & 0xff, (x >>> 8) & 0xff, x & 0xff);
    return this;
  }

  le16(...v: number[]): this {
    for (const x of v) this.out.push(x & 0xff, (x >> 8) & 0xff);
    return this;
  }

  le32(...v: number[]): this {
    for (const x of v) {
      this.out.push(x & 0xff, (x >>> 8) & 0xff, (x >>> 16) & 0xff, (x >>> 24) & 0xff);
    }
    return this;
  }

  bytes(b: ArrayLike<number>): this {
    for (let i = 0; i < b.length; i++) this.out.push(b[i] & 0xff);
    return this;
  }

  get length(): number {
    return this.out.length;
  }

  done(): Uint8Array {
    return Uint8Array.from(this.out);
  }
}

// --- LZCOMP encoder (test only) ------------------------------------------------------------

/** Run-length pre-pass: the rarest byte value is the escape; runs of 4 … 255 become E, r, v. */
export function rlePack(data: Uint8Array): Uint8Array {
  const counts = new Array<number>(256).fill(0);
  for (const v of data) counts[v]++;
  let escape = 0;
  for (let v = 1; v < 256; v++) if (counts[v] < counts[escape]) escape = v;
  const out: number[] = [escape];
  let i = 0;
  while (i < data.length) {
    const v = data[i];
    let run = 1;
    while (run < 255 && i + run < data.length && data[i + run] === v) run++;
    if (run >= 4) {
      out.push(escape, run, v);
      i += run;
    } else {
      if (v === escape) out.push(escape, 0);
      else out.push(v);
      i += 1;
    }
  }
  return Uint8Array.from(out);
}

export interface LzEncodeOptions {
  /** Apply `rlePack` first and set the run-length flag. */
  runLength?: boolean;
  /** Skip the match search: emit only literals and DUP2/4/6. */
  literalsOnly?: boolean;
}

/** Number of binary digits of x, with bitLength(0) = 1. */
function bitLength(x: number): number {
  let n = 1;
  for (let v = x; v >= 2; v = Math.floor(v / 2)) n++;
  return n;
}

const MATCH_SEARCH_SPAN = 9000;

/** Longest earlier match for the bytes at `pos`, as a copy item the decoder accepts. */
function findMatch(
  win: Uint8Array,
  pos: number,
  end: number,
  maxDistance: number
): { length: number; distance: number } {
  let best = { length: 0, distance: 0 };
  for (let start = pos - 2; start >= Math.max(0, pos - MATCH_SEARCH_SPAN); start--) {
    let length = 0;
    while (start + length < pos && pos + length < end && win[start + length] === win[pos + length])
      length++;
    const distance = pos - start - length + 1;
    if (length < 2 || (length < 3 && distance >= LZ_MAX_2BYTE_DIST) || distance > maxDistance)
      continue;
    if (length > best.length) best = { length, distance };
  }
  return best;
}

function writeCopy(
  w: BitWriter,
  models: { main: AdaptiveHuffman; lengths: AdaptiveHuffman; distances: AdaptiveHuffman },
  length: number,
  distance: number
): void {
  const value = length - (distance < LZ_MAX_2BYTE_DIST ? 2 : 3);
  const groups = Math.ceil(bitLength(distance - 1) / 3);
  const chunkCount = Math.max(1, Math.ceil(bitLength(value) / 2));
  const chunk = (i: number): number => Math.floor(value / 4 ** (chunkCount - 1 - i)) % 4;
  models.main.writeSymbol(w, 256 + 8 * (groups - 1) + chunk(0) + (chunkCount > 1 ? 4 : 0));
  for (let i = 1; i < chunkCount; i++) {
    models.lengths.writeSymbol(w, chunk(i) + (i < chunkCount - 1 ? 4 : 0));
  }
  for (let g = groups - 1; g >= 0; g--) {
    models.distances.writeSymbol(w, Math.floor((distance - 1) / 8 ** g) % 8);
  }
}

/** Greedy LZCOMP encoder (version 3 layout: the run-length flag bit is always written). */
export function lzcompEncode(input: Uint8Array, opts: LzEncodeOptions = {}): Uint8Array {
  const payload = opts.runLength ? rlePack(input) : input;
  const w = new BitWriter();
  w.writeBit(opts.runLength ? 1 : 0);
  w.writeValue(payload.length, 24);
  const layout = lzSymbolLayout(payload.length);
  const models = {
    distances: new AdaptiveHuffman(8),
    lengths: new AdaptiveHuffman(8),
    main: new AdaptiveHuffman(layout.numSyms),
  };
  const maxDistance = 8 ** layout.numDistRanges;
  const end = LZ_PRELOAD_SIZE + payload.length;
  const win = new Uint8Array(end);
  writePreload(win);
  win.set(payload, LZ_PRELOAD_SIZE);
  let pos = LZ_PRELOAD_SIZE;
  while (pos < end) {
    const match = opts.literalsOnly
      ? { length: 0, distance: 0 }
      : findMatch(win, pos, end, maxDistance);
    if (match.length >= 3) {
      writeCopy(w, models, match.length, match.distance);
      pos += match.length;
      continue;
    }
    const x = win[pos];
    if (x === win[pos - 2]) models.main.writeSymbol(w, layout.dup2);
    else if (x === win[pos - 4]) models.main.writeSymbol(w, layout.dup4);
    else if (x === win[pos - 6]) models.main.writeSymbol(w, layout.dup6);
    else models.main.writeSymbol(w, x);
    pos += 1;
  }
  return w.finish();
}

export function enc255UShort(v: number): number[] {
  if (v < 253) return [v];
  if (v < 506) return [255, v - 253];
  if (v < 762) return [254, v - 506];
  return [253, (v >> 8) & 0xff, v & 0xff];
}

export function enc255Short(v: number): number[] {
  const mag = Math.abs(v);
  const sign = v < 0 ? [250] : [];
  if (mag < 250) return [...sign, mag];
  if (mag < 500) return [...sign, 255, mag - 250];
  if (mag < 756) return [...sign, 254, mag - 500];
  return [253, (v >> 8) & 0xff, v & 0xff];
}

/** Encodes one point with the generic 16/16-bit triplet (indices 124..127). */
export function encTriplet16(
  dx: number,
  dy: number,
  onCurve: boolean
): { flag: number; data: number[] } {
  const index = 124 + (dx >= 0 ? 1 : 0) + (dy >= 0 ? 2 : 0);
  const ax = Math.abs(dx);
  const ay = Math.abs(dy);
  return {
    flag: index | (onCurve ? 0 : 0x80),
    data: [(ax >> 8) & 0xff, ax & 0xff, (ay >> 8) & 0xff, ay & 0xff],
  };
}

export function magnitudeBits(values: number[]): Uint8Array {
  const bits: number[] = [];
  for (const v of values) {
    for (let i = 0; i < Math.abs(v); i++) bits.push(1);
    bits.push(0);
    if (v !== 0) bits.push(v < 0 ? 1 : 0);
  }
  const out = new Uint8Array(Math.ceil(bits.length / 8));
  bits.forEach((b, i) => {
    if (b) out[i >> 3] |= 1 << (i & 7);
  });
  return out;
}

// --- synthetic font --------------------------------------------------------------------

export const SQUARE = [
  { x: 0, y: 0, on: true },
  { x: 0, y: 700, on: true },
  { x: 500, y: 700, on: false },
  { x: 500, y: 0, on: true },
  { x: 100, y: 100, on: true },
  { x: 200, y: 100, on: true },
  { x: 150, y: 200, on: true },
];
export const SQUARE_ENDPTS = [3, 6];
export const GLYPH1_PUSH = [5, 7, 5, 9, 5];
export const GLYPH1_CODE = [0x2b, 0x00, 0x01];
export const GLYPH2_PUSH = [-1, 300];
export const GLYPH2_CODE = [0x00];
export const CVT_VALUES = [10, 253, 250, -32768, -2000, 1500, 1000];
export const HDMX_WIDTHS = [6, 8, 7, 7];
export const ADVANCES = [500, 600, 600, 600];
/** Glyph 3 points, hand-encoded with spec triplet indices (see buildCtfGlyf). */
export const GLYPH3_POINTS = [
  { x: 0, y: -5, on: true },
  { x: 3, y: -3, on: true },
  { x: -14, y: -36, on: false },
  { x: 242, y: 476, on: true },
];
export const GLYPH3_BOX = [-20, -40, 260, 480];

function buildCtfGlyf(): { glyf: Uint8Array; push: Uint8Array; code: Uint8Array } {
  const g = new Bytes();
  const push = new Bytes();
  const code = new Bytes();
  // glyph 0: empty
  g.u16(0);
  // glyph 1: simple, computed bbox, two contours
  g.u16(2);
  g.bytes(enc255UShort(SQUARE_ENDPTS[0]));
  g.bytes(enc255UShort(SQUARE_ENDPTS[1] - SQUARE_ENDPTS[0]));
  const flags: number[] = [];
  const coords: number[] = [];
  let px = 0;
  let py = 0;
  for (const p of SQUARE) {
    const t = encTriplet16(p.x - px, p.y - py, p.on);
    flags.push(t.flag);
    coords.push(...t.data);
    px = p.x;
    py = p.y;
  }
  g.bytes(flags).bytes(coords);
  g.bytes(enc255UShort(GLYPH1_PUSH.length)).bytes(enc255UShort(GLYPH1_CODE.length));
  push.bytes([5, 7, 251, 9]); // A, X1, Hop3Code, X2 -> 5,7,5,9,5
  code.bytes(GLYPH1_CODE);
  // glyph 2: composite of glyph 1 with instructions
  g.u16(0xffff).u16(10, 0xffec, 510, 680);
  g.u16(0x0001 | 0x0002 | 0x0100)
    .u16(1)
    .u16(10, 0xffec);
  g.bytes(enc255UShort(GLYPH2_PUSH.length)).bytes(enc255UShort(GLYPH2_CODE.length));
  for (const v of GLYPH2_PUSH) push.bytes(enc255Short(v));
  code.bytes(GLYPH2_CODE);
  // glyph 3: explicit bbox (0x7FFF), one contour, triplets chosen from section 5.11
  g.u16(0x7fff)
    .u16(1)
    .u16(...GLYPH3_BOX.map((v) => v & 0xffff));
  g.bytes(enc255UShort(3));
  g.u8(0x00, 0x17, 0x80 | 84, 127); // indices 0, 23, 84 (off-curve), 127
  g.u8(0x05); // idx 0: dx 0, dy -(5+0)
  g.u8(0x21); // idx 23: dx +(2+1), dy +(1+1)
  g.u8(0x10, 0x20); // idx 84: dx -(16+1), dy -(32+1)
  g.u8(0x01, 0x00, 0x02, 0x00); // idx 127: dx +256, dy +512
  g.bytes(enc255UShort(0)).bytes(enc255UShort(0));
  return { glyf: g.done(), push: push.done(), code: code.done() };
}

function cvtCtf(): Uint8Array {
  const b = new Bytes().u16(CVT_VALUES.length);
  let prev = 0;
  for (const v of CVT_VALUES) {
    const delta = (((v - prev) & 0xffff) << 16) >> 16;
    prev = v;
    const mag = Math.abs(delta);
    if (delta === -32768 || mag >= 238 * 9) b.u8(238).u16(delta & 0xffff);
    else if (delta >= 0 && mag < 238) b.u8(mag);
    else if (delta >= 0) b.u8(248 + Math.floor(mag / 238) - 1, mag % 238);
    else b.u8(239 + Math.floor(mag / 238), mag % 238);
  }
  return b.done();
}

function hdmxCtf(): Uint8Array {
  const predicted = ADVANCES.map((aw) =>
    Math.floor((Math.floor((64 * 12 * aw + 500) / 1000) + 32) / 64)
  );
  const surprises = HDMX_WIDTHS.map((w, i) => w - predicted[i]);
  return new Bytes().u16(0, 1).u32(8).u8(12, 8).bytes(magnitudeBits(surprises)).done();
}

function commonTables(numGlyphs: number): SfntTable[] {
  const head = new Bytes()
    .u32(0x00010000, 0x00010000, 0, 0x5f0f3cf5)
    .u16(0x000b, 1000)
    .u32(0, 0, 0, 0)
    .u16(0xffec, 0xffd8, 510, 700, 0, 8, 2, 0, 0)
    .done();
  const hhea = new Bytes()
    .u32(0x00010000)
    .u16(800, 0xff38, 0, 600, 0, 0, 600, 1, 0, 0, 0, 0, 0, 0, 0, numGlyphs)
    .done();
  const maxp = new Bytes()
    .u32(0x00010000)
    .u16(numGlyphs, 16, 2, 7, 1, 2, 0, 0, 0, 0, 0, 64, 1, 1)
    .done();
  const hmtx = new Bytes();
  for (const aw of ADVANCES.slice(0, numGlyphs)) hmtx.u16(aw, 0);
  // cmap format 4: 'A' -> 1, 'B' -> 2, 'C' -> 3
  const cmap4 = new Bytes()
    .u16(4, 32, 0, 4, 4, 1, 0)
    .u16(0x43, 0xffff)
    .u16(0)
    .u16(0x41, 0xffff)
    .u16((1 - 0x41) & 0xffff, 1)
    .u16(0, 0)
    .done();
  const cmap = new Bytes().u16(0, 1).u16(3, 1).u32(12).bytes(cmap4).done();
  const nameStr = 'Synth';
  const name = new Bytes().u16(0, 1, 18).u16(3, 1, 0x409, 1, nameStr.length * 2, 0);
  for (const ch of nameStr) name.u16(ch.charCodeAt(0));
  const post = new Bytes().u32(0x00030000, 0).u16(0, 0).u32(0, 0, 0, 0, 0).done();
  return [
    { tag: 'head', data: head },
    { tag: 'hhea', data: hhea },
    { tag: 'maxp', data: maxp },
    { tag: 'hmtx', data: hmtx.done() },
    { tag: 'cmap', data: cmap },
    { tag: 'name', data: name.done() },
    { tag: 'post', data: post },
  ];
}

/** CTF block 1: an sfnt whose glyf/cvt/hdmx are CTF-encoded and whose loca is empty. */
export function buildCtf(): { block1: Uint8Array; block2: Uint8Array; block3: Uint8Array } {
  const g = buildCtfGlyf();
  const tables: SfntTable[] = [
    ...commonTables(4),
    { tag: 'glyf', data: g.glyf },
    { tag: 'loca', data: new Uint8Array(0) },
    { tag: 'cvt ', data: cvtCtf() },
    { tag: 'hdmx', data: hdmxCtf() },
  ];
  const block1 = buildSfnt(0x00010000, tables, 1 << 24);
  // CTF stores loca with offset and length zero.
  const n = (block1[4] << 8) | block1[5];
  for (let i = 0; i < n; i++) {
    const rec = 12 + i * 16;
    if (String.fromCharCode(...block1.subarray(rec, rec + 4)) === 'loca') {
      block1.fill(0, rec + 4, rec + 16);
    }
  }
  return { block1, block2: g.push, block3: g.code };
}

export function buildMtx(opts: LzEncodeOptions = {}): Uint8Array {
  const ctf = buildCtf();
  const b1 = lzcompEncode(ctf.block1, opts);
  const b2 = lzcompEncode(ctf.block2, opts);
  const b3 = lzcompEncode(ctf.block3, opts);
  const o2 = 10 + b1.length;
  const o3 = o2 + b2.length;
  const total = Math.max(ctf.block1.length, 50000);
  return new Bytes()
    .u8(3)
    .u8((total >> 16) & 0xff, (total >> 8) & 0xff, total & 0xff)
    .u8((o2 >> 16) & 0xff, (o2 >> 8) & 0xff, o2 & 0xff)
    .u8((o3 >> 16) & 0xff, (o3 >> 8) & 0xff, o3 & 0xff)
    .bytes(b1)
    .bytes(b2)
    .bytes(b3)
    .done();
}

/** Plain (uncompressed) TrueType equivalent, for the non-MTX EOT path. */
export function buildPlainTtf(): Uint8Array {
  // Triangle (0,0) (100,0) (100,100): flags on|x-same|y-same, on|x-short+|y-same, on|y-short+.
  const glyph = new Bytes().u16(1, 0, 0, 100, 100, 2, 0).u8(0x31, 0x33, 0x35).u8(100, 100).done();
  const glyf = new Bytes().bytes(glyph).u8(0).done();
  const loca = new Bytes().u16(0, glyf.length / 2).done();
  const tables = commonTables(1);
  tables.push({ tag: 'glyf', data: glyf }, { tag: 'loca', data: loca });
  return buildSfnt(0x00010000, tables, 1 << 20);
}

export interface EotOptions {
  version?: number;
  flags?: number;
  family?: string;
  style?: string;
  fsType?: number;
  weight?: number;
  italic?: boolean;
  charset?: number;
  rootString?: string;
}

function utf16(s: string): number[] {
  const out: number[] = [];
  for (const ch of s) {
    const c = ch.charCodeAt(0);
    out.push(c & 0xff, c >> 8);
  }
  return out;
}

export function buildEot(fontData: Uint8Array, o: EotOptions = {}): Uint8Array {
  const version = o.version ?? 0x00020002;
  const flags = o.flags ?? 0x4;
  const data = flags & 0x10000000 ? fontData.map((b) => b ^ 0x50) : fontData;
  const body = new Bytes()
    .le32(data.length, version, flags)
    .u8(2, 0, 5, 3, 0, 0, 0, 0, 0, 0)
    .u8(o.charset ?? 1, o.italic ? 1 : 0)
    .le32(o.weight ?? 400)
    .le16(o.fsType ?? 0, 0x504c)
    .le32(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
  const str = (s: string): void => {
    const b = utf16(s);
    body.le16(0, b.length).bytes(b);
  };
  str(o.family ?? 'Synth');
  str(o.style ?? 'Regular');
  str('Version 1.0');
  str(`${o.family ?? 'Synth'} ${o.style ?? 'Regular'}`);
  if (version !== 0x00010000) {
    const root = utf16(o.rootString ?? '');
    body.le16(0, root.length).bytes(root);
    if (version === 0x00020002) {
      let sum = 0;
      for (const b of root) sum += b;
      body
        .le32((sum ^ 0x50475342) >>> 0, 0)
        .le16(0, 0)
        .le32(0, 0);
    }
  }
  body.bytes(data);
  const out = body.done();
  const total = out.length + 4;
  return new Bytes().le32(total).bytes(out).done();
}

/** Deterministic PRNG for fuzz-style tests. */
export function prng(seed: number): () => number {
  let s = seed >>> 0;
  return () => {
    s = (s + 0x6d2b79f5) >>> 0;
    let t = s;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
