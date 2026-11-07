// Recognize EMF and WMF pictures and read their picture frame from the header. Every read is
// bounds-checked, because the bytes come from the file (AGENTS.md "Security"). Nothing is
// decoded: a host converts the records, for example to a raster, outside this package.

/** The two metafile formats. */
export type MetafileFormat = 'emf' | 'wmf';

/** EMU per 0.01 mm, the unit of an EMF picture frame. */
const EMU_PER_HUNDREDTH_MM = 360;
/** EMU per inch. */
const EMU_PER_INCH = 914_400;
/** The size of an EMF header record with no extensions, in bytes. */
const EMF_HEADER_MIN = 88;
/** The ` EMF` record signature at byte 40 of an EMF header, little endian. */
const EMF_SIGNATURE = 0x464d4520;
/** The key of a placeable WMF header, little endian. */
const WMF_PLACEABLE_KEY = 0x9ac6cdd7;
/** The size of a placeable WMF header, in bytes. */
const WMF_PLACEABLE_SIZE = 22;
/** The largest frame side accepted: 10 m, far past any slide. */
const MAX_FRAME_EMU = 10 * 1000 * 36_000;

const u16le = (b: Uint8Array, i: number): number => b[i]! | (b[i + 1]! << 8);
const i16le = (b: Uint8Array, i: number): number => (u16le(b, i) << 16) >> 16;
const u32le = (b: Uint8Array, i: number): number =>
  (b[i]! | (b[i + 1]! << 8) | (b[i + 2]! << 16) | (b[i + 3]! << 24)) >>> 0;
const i32le = (b: Uint8Array, i: number): number => u32le(b, i) | 0;

/** Whether `b` starts with an EMF header record: type 1, a plausible size, and ` EMF`. */
function isEmf(b: Uint8Array): boolean {
  if (b.length < EMF_HEADER_MIN || u32le(b, 0) !== 1) return false;
  const size = u32le(b, 4);
  return size >= EMF_HEADER_MIN && size % 4 === 0 && u32le(b, 40) === EMF_SIGNATURE;
}

/** Whether `b` starts with a WMF header: placeable, or a memory or disk metafile header. */
function isWmf(b: Uint8Array): boolean {
  if (b.length >= WMF_PLACEABLE_SIZE && u32le(b, 0) === WMF_PLACEABLE_KEY) return true;
  return wmfHeaderAt(b, 0);
}

/** Whether a WMF header (type 1 or 2, 9 words, version 1 or 3) starts at `at`. */
function wmfHeaderAt(b: Uint8Array, at: number): boolean {
  if (b.length < at + 18) return false;
  const type = u16le(b, at);
  const version = u16le(b, at + 4);
  return (
    (type === 1 || type === 2) && u16le(b, at + 2) === 9 && (version === 0x100 || version === 0x300)
  );
}

/** The metafile format of `bytes` from its signature, or undefined for anything else. */
export function metafileFormat(bytes: Uint8Array): MetafileFormat | undefined {
  if (isEmf(bytes)) return 'emf';
  if (isWmf(bytes)) return 'wmf';
  return undefined;
}

/** A plausible frame side in EMU, or undefined. */
const side = (emu: number): number | undefined =>
  Number.isFinite(emu) && emu > 0 && emu <= MAX_FRAME_EMU ? Math.round(emu) : undefined;

/**
 * The picture frame of a metafile in EMU: the EMF header frame (in 0.01 mm), or the bounding
 * box of a placeable WMF over its units per inch. Undefined for a WMF with no placeable
 * header, which states no physical size, and for an empty or implausible frame.
 */
export function metafileSizeEmu(bytes: Uint8Array): { cx: number; cy: number } | undefined {
  let cx: number | undefined;
  let cy: number | undefined;
  if (isEmf(bytes)) {
    // `rclFrame` at byte 24: left, top, right, bottom.
    cx = side((i32le(bytes, 32) - i32le(bytes, 24)) * EMU_PER_HUNDREDTH_MM);
    cy = side((i32le(bytes, 36) - i32le(bytes, 28)) * EMU_PER_HUNDREDTH_MM);
  } else if (bytes.length >= WMF_PLACEABLE_SIZE && u32le(bytes, 0) === WMF_PLACEABLE_KEY) {
    // Bounding box at byte 6: left, top, right, bottom; units per inch at byte 14.
    const inch = u16le(bytes, 14);
    if (inch === 0) return undefined;
    cx = side((Math.abs(i16le(bytes, 10) - i16le(bytes, 6)) * EMU_PER_INCH) / inch);
    cy = side((Math.abs(i16le(bytes, 12) - i16le(bytes, 8)) * EMU_PER_INCH) / inch);
  }
  return cx && cy ? { cx, cy } : undefined;
}
