// Test-only sfnt inspection helpers (OpenType 'glyf'/'loca' simple + composite parsing).

const u16 = (d: Uint8Array, p: number): number => (d[p] << 8) | d[p + 1];
const i16 = (d: Uint8Array, p: number): number => (u16(d, p) << 16) >> 16;
const u32 = (d: Uint8Array, p: number): number =>
  ((d[p] << 24) >>> 0) + ((d[p + 1] << 16) | (d[p + 2] << 8) | d[p + 3]);

export function getTable(font: Uint8Array, tag: string): Uint8Array | undefined {
  const n = u16(font, 4);
  for (let i = 0; i < n; i++) {
    const r = 12 + i * 16;
    if (String.fromCharCode(...font.subarray(r, r + 4)) === tag) {
      const off = u32(font, r + 8);
      return font.subarray(off, off + u32(font, r + 12));
    }
  }
  return undefined;
}

export function tableTags(font: Uint8Array): string[] {
  const out: string[] = [];
  for (let i = 0; i < u16(font, 4); i++) {
    const r = 12 + i * 16;
    out.push(String.fromCharCode(...font.subarray(r, r + 4)));
  }
  return out;
}

export interface ParsedGlyph {
  numContours: number;
  box: number[];
  endPts: number[];
  instructions: number[];
  points: { x: number; y: number; on: boolean }[];
  components: { flags: number; glyphIndex: number; args: number[] }[];
}

export function glyphBytes(font: Uint8Array, gid: number): Uint8Array {
  const head = getTable(font, 'head');
  const loca = getTable(font, 'loca');
  const glyf = getTable(font, 'glyf');
  if (!head || !loca || !glyf) throw new Error('missing tables');
  const long = u16(head, 50) === 1;
  const off = (g: number): number => (long ? u32(loca, g * 4) : u16(loca, g * 2) * 2);
  return glyf.subarray(off(gid), off(gid + 1));
}

export function parseGlyph(g: Uint8Array): ParsedGlyph | null {
  if (g.length === 0) return null;
  const numContours = i16(g, 0);
  const box = [i16(g, 2), i16(g, 4), i16(g, 6), i16(g, 8)];
  const result: ParsedGlyph = {
    numContours,
    box,
    endPts: [],
    instructions: [],
    points: [],
    components: [],
  };
  let p = 10;
  if (numContours < 0) {
    let flags: number;
    let instr = false;
    do {
      flags = u16(g, p);
      const glyphIndex = u16(g, p + 2);
      p += 4;
      const args = flags & 1 ? [i16(g, p), i16(g, p + 2)] : [g[p], g[p + 1]];
      p += flags & 1 ? 4 : 2;
      if (flags & 0x8) p += 2;
      else if (flags & 0x40) p += 4;
      else if (flags & 0x80) p += 8;
      if (flags & 0x100) instr = true;
      result.components.push({ flags, glyphIndex, args });
    } while (flags & 0x20);
    if (instr) {
      const n = u16(g, p);
      result.instructions = [...g.subarray(p + 2, p + 2 + n)];
    }
    return result;
  }
  for (let c = 0; c < numContours; c++) result.endPts.push(u16(g, p + c * 2));
  p += numContours * 2;
  const n = u16(g, p);
  result.instructions = [...g.subarray(p + 2, p + 2 + n)];
  p += 2 + n;
  const numPoints = numContours ? result.endPts[numContours - 1] + 1 : 0;
  const flags: number[] = [];
  while (flags.length < numPoints) {
    const f = g[p++];
    flags.push(f);
    if (f & 8) for (let k = g[p++]; k > 0; k--) flags.push(f);
  }
  const xs: number[] = [];
  let x = 0;
  for (const f of flags) {
    if (f & 2) x += f & 0x10 ? g[p++] : -g[p++];
    else if (!(f & 0x10)) {
      x += i16(g, p);
      p += 2;
    }
    xs.push(x);
  }
  let y = 0;
  flags.forEach((f, i) => {
    if (f & 4) y += f & 0x20 ? g[p++] : -g[p++];
    else if (!(f & 0x20)) {
      y += i16(g, p);
      p += 2;
    }
    result.points.push({ x: xs[i], y, on: (f & 1) === 1 });
  });
  if (p > g.length) throw new Error('glyph overruns its loca slot');
  return result;
}

export { u16 as readU16, i16 as readI16, u32 as readU32 };
