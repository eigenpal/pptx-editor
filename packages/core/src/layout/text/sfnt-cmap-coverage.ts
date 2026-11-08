// Adapted from DOCX Editor `packages/core/src/layout/sfnt-cmap-coverage.ts` at commit adf0e178.
// Changes: `unicodeCmapSubtable` ranks the encodings as HarfBuzz and fontkit do, and passes
// over a subtable that is cut short or whose segments or groups are out of order. And `cmapMapsCodePoint` finds a format 4 segment with a binary search instead of a scan. A
// deck can embed a font with 32,767 segments, and the fallback search asks every face for
// every character it lacks. `fontMapsCodePoint` is not ported: the HarfBuzz measurer keeps
// each face's subtable, and a coverage cache, with the loaded face.
//
// Whether a face's own cmap maps a code point to a glyph.
//
// Shaping cannot answer this: HarfBuzz draws a missing U+2011 with the face's U+2010, and draws
// a missing precomposed letter as its decomposition. Both hide that the face itself has no
// glyph. Only the Unicode subtables count, in the order of `UNICODE_ENCODINGS`. Every read is bounds-checked, because the bytes come from a file. A face with no
// subtable that passes these checks has no answer here, and its caller asks HarfBuzz.

const u16 = (b: Uint8Array, o: number): number => (b[o]! << 8) | b[o + 1]!;
const u32 = (b: Uint8Array, o: number): number =>
  (b[o]! * 0x1000000 + (b[o + 1]! << 16) + (b[o + 2]! << 8) + b[o + 3]!) >>> 0;

/**
 * The Unicode encodings, as `platform * 0x10000 + encoding`, first choice first: (3,10) and
 * (0,4) for the full repertoire, then (3,1) and (0,3) to (0,0) for the BMP. A face whose
 * (0,3) table is a subset of its (3,1) table answers from (3,1), as HarfBuzz does.
 */
const UNICODE_ENCODINGS = [0x3000a, 0x00004, 0x30001, 0x00003, 0x00002, 0x00001, 0x00000];

/** The chosen Unicode subtable: its format and offset in the font bytes. */
export interface UnicodeSubtable {
  readonly format: 4 | 12;
  readonly offset: number;
}

/**
 * Whether the subtable at `at` fits in `bytes` with its segments (format 4) or groups
 * (format 12) in increasing order, as the binary search in `cmapMapsCodePoint` needs. One pass
 * over the subtable, once for each face.
 */
function wellFormed(bytes: Uint8Array, format: number, at: number): boolean {
  if (format === 12) {
    if (at + 16 > bytes.length) return false;
    const groups = u32(bytes, at + 12);
    if (groups > Math.floor((bytes.length - at - 16) / 12)) return false;
    let previousEnd = -1;
    for (let group = at + 16, last = group + groups * 12; group < last; group += 12) {
      const start = u32(bytes, group);
      const end = u32(bytes, group + 4);
      if (start <= previousEnd || end < start) return false;
      previousEnd = end;
    }
    return true;
  }
  if (format !== 4 || at + 14 > bytes.length) return false;
  const segments = u16(bytes, at + 6) >>> 1;
  if (at + 16 + segments * 8 > bytes.length) return false;
  let previousEnd = -1;
  for (let segment = 0; segment < segments; segment += 1) {
    const end = u16(bytes, at + 14 + segment * 2);
    if (end <= previousEnd) return false;
    previousEnd = end;
  }
  return true;
}

/** The face's Unicode cmap subtable, or null when it has none this lookup reads. */
export function unicodeCmapSubtable(bytes: Uint8Array, faceIndex: number): UnicodeSubtable | null {
  const fits = (offset: number, length: number) => offset >= 0 && offset + length <= bytes.length;
  if (!fits(0, 12)) return null;
  const collection = u32(bytes, 0) === 0x74746366;
  if (collection && !fits(12 + faceIndex * 4, 4)) return null;
  const base = collection ? u32(bytes, 12 + faceIndex * 4) : 0;
  if (!fits(base, 12)) return null;
  const tables = u16(bytes, base + 4);
  for (let index = 0; index < tables; index += 1) {
    const record = base + 12 + index * 16;
    if (!fits(record, 16)) return null;
    if (u32(bytes, record) !== 0x636d6170) continue;
    const cmap = u32(bytes, record + 8);
    if (!fits(cmap, 4)) return null;
    // The best-ranked subtable of format 4 or 12 that is well formed; a broken one is passed.
    let best: UnicodeSubtable | null = null;
    let bestRank = UNICODE_ENCODINGS.length;
    const encodings = u16(bytes, cmap + 2);
    for (let entry = 0; entry < encodings; entry += 1) {
      const at = cmap + 4 + entry * 8;
      if (!fits(at, 8)) break;
      const rank = UNICODE_ENCODINGS.indexOf(u16(bytes, at) * 0x10000 + u16(bytes, at + 2));
      if (rank < 0 || rank >= bestRank) continue;
      const offset = cmap + u32(bytes, at + 4);
      if (!fits(offset, 2)) continue;
      const format = u16(bytes, offset);
      if ((format !== 4 && format !== 12) || !wellFormed(bytes, format, offset)) continue;
      best = { format, offset };
      bestRank = rank;
    }
    return best;
  }
  return null;
}

/** Whether `subtable` maps `codePoint` to a glyph other than notdef. */
export function cmapMapsCodePoint(
  bytes: Uint8Array,
  subtable: UnicodeSubtable,
  codePoint: number
): boolean {
  const fits = (offset: number, length: number) => offset >= 0 && offset + length <= bytes.length;
  const at = subtable.offset;
  if (subtable.format === 12) {
    if (!fits(at, 16)) return false;
    const groups = u32(bytes, at + 12);
    // Groups are sorted by start: a binary search reads at most log2(groups) records.
    let low = 0;
    let high = Math.min(groups, Math.floor((bytes.length - at - 16) / 12)) - 1;
    while (low <= high) {
      const middle = (low + high) >>> 1;
      const group = at + 16 + middle * 12;
      const start = u32(bytes, group);
      const end = u32(bytes, group + 4);
      if (codePoint < start) high = middle - 1;
      else if (codePoint > end) low = middle + 1;
      else return u32(bytes, group + 8) + (codePoint - start) !== 0;
    }
    return false;
  }
  if (codePoint > 0xffff || !fits(at, 14)) return false;
  const segments = u16(bytes, at + 6) >>> 1;
  const ends = at + 14;
  const starts = ends + segments * 2 + 2;
  const deltas = starts + segments * 2;
  const rangeOffsets = deltas + segments * 2;
  if (!fits(rangeOffsets, segments * 2)) return false;
  // Segments are sorted by end code (`unicodeCmapSubtable` refuses a table that is not): the
  // first segment that ends at or after `codePoint` is the only one that can hold it.
  let low = 0;
  let high = segments;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (u16(bytes, ends + middle * 2) < codePoint) low = middle + 1;
    else high = middle;
  }
  if (low < segments) {
    const segment = low;
    const start = u16(bytes, starts + segment * 2);
    if (codePoint < start) return false;
    const delta = u16(bytes, deltas + segment * 2);
    const rangeOffsetAt = rangeOffsets + segment * 2;
    const rangeOffset = u16(bytes, rangeOffsetAt);
    if (rangeOffset === 0) return ((codePoint + delta) & 0xffff) !== 0;
    const glyphAt = rangeOffsetAt + rangeOffset + (codePoint - start) * 2;
    if (!fits(glyphAt, 2)) return false;
    const glyph = u16(bytes, glyphAt);
    return glyph !== 0 && ((glyph + delta) & 0xffff) !== 0;
  }
  return false;
}
