// Copied from docx-editor/packages/core/src/store/package/zip.ts at 7cd3776a (EigenPal).
// It also holds fixes that DOCX Editor does not have yet: S1, S6, K1, P6, W1.
// Keep it equal to that file plus those fixes.
// Bounded ZIP read/write over fflate (document-engine task 2.3 / design D14).
// Reading enforces entry-count, total-decompressed-size and archive-ratio ceilings
// (zip-bomb guard) and normalizes every entry name through the OPC profile
// (path-traversal guard) BEFORE the bytes are handed on. Writing produces a
// deterministic archive.

import { unzipSync } from 'fflate';
import { normalizePartName, partNameKey, zipItemName } from './opc-names.ts';
import { writeZipEntriesSteps } from './zip-write.ts';

/**
 * Why a zip was refused at the trust boundary.
 *
 * `too-large` and `too-many-entries` are the zip-bomb guards; `bad-name` catches path traversal —
 * an entry with `..` or a leading `/` is rejected rather than normalized. `overlapping-entries`
 * refuses an archive whose entries share bytes (many directory records over one stored entry
 * is a zip bomb that no size field reveals). `invalid-limits` refuses a limit that is not a
 * finite, non-negative number, so a bad configuration can never turn the caps off.
 */
export type ZipRejection =
  | 'too-many-entries'
  | 'too-large'
  | 'bad-name'
  | 'inflate-error'
  | 'overlapping-entries'
  | 'invalid-limits';

/**
 * Declared uncompressed bytes an archive may reach before `maxRatio` applies: 8 MiB.
 *
 * Small parts are where legitimate documents carry extreme ratios: a few kilobytes of zero
 * bytes, or a metafile drawn from long uniform runs, stores in a few dozen bytes. Above this
 * size the ratio applies to the whole archive, so splitting a bomb across many small entries
 * does not evade it. The most any archive may expand to is therefore the larger of 8 MiB and
 * `maxRatio` times its own length, and never more than `maxTotalBytes`: a tiny upload costs
 * the reader at most 8 MiB.
 */
export const ZIP_RATIO_EXEMPT_BYTES = 8 * 1024 * 1024;

/** Archive caps: entry count, total decompressed bytes, and the decompression ratio. */
export interface ZipLimits {
  readonly maxEntries: number;
  /** Max total UNCOMPRESSED bytes across the archive. */
  readonly maxTotalBytes: number;
  /**
   * Max ratio of the archive's total declared uncompressed bytes to the archive's own byte
   * length (zip-bomb guard). Applies once that total exceeds 8 MiB, so a small archive is
   * never refused for compressing well.
   */
  readonly maxRatio?: number;
}

/** The archive caps in force when a host configures none. Conservative and finite. */
export const DEFAULT_ZIP_LIMITS: ZipLimits = Object.freeze({
  maxEntries: 10_000,
  maxTotalBytes: 512 * 1024 * 1024,
  maxRatio: 200,
});

/** The ceilings no host configuration can raise. */
export const ZIP_HARD_LIMITS: Readonly<{ maxEntries: number; maxTotalBytes: number }> =
  Object.freeze({ maxEntries: 100_000, maxTotalBytes: 2 * 1024 * 1024 * 1024 });

/** The read entries, or the typed refusal. Never throws — the bytes are untrusted. */
export type ZipReadResult =
  | { readonly ok: true; readonly entries: ReadonlyMap<string, Uint8Array> }
  | {
      readonly ok: false;
      readonly reason: ZipRejection;
      readonly detail?: string;
      readonly limit?: 'zip.maxTotalBytes' | 'zip.maxRatio';
    };

class ZipViolation extends Error {
  constructor(
    readonly reason: ZipRejection,
    readonly limit?: 'zip.maxTotalBytes' | 'zip.maxRatio'
  ) {
    super(reason);
  }
}

interface ResolvedLimits {
  readonly maxEntries: number;
  readonly maxTotalBytes: number;
  readonly maxRatio: number;
}

const isCount = (v: unknown): v is number =>
  typeof v === 'number' && Number.isSafeInteger(v) && v >= 0;

/** Validate and clamp host limits. A missing field takes the default; a bad one refuses. */
function resolveZipLimits(limits: ZipLimits): ResolvedLimits | null {
  if (limits === null || typeof limits !== 'object') return null;
  const { maxEntries, maxTotalBytes, maxRatio } = limits as Partial<ZipLimits>;
  const entries = maxEntries ?? DEFAULT_ZIP_LIMITS.maxEntries;
  const total = maxTotalBytes ?? DEFAULT_ZIP_LIMITS.maxTotalBytes;
  const ratio = maxRatio ?? DEFAULT_ZIP_LIMITS.maxRatio ?? 200;
  if (!isCount(entries) || !isCount(total)) return null;
  if (typeof ratio !== 'number' || !Number.isFinite(ratio) || ratio < 1) return null;
  return {
    maxEntries: Math.min(entries, ZIP_HARD_LIMITS.maxEntries),
    maxTotalBytes: Math.min(total, ZIP_HARD_LIMITS.maxTotalBytes),
    maxRatio: ratio,
  };
}

const u16 = (d: Uint8Array, p: number): number => d[p]! | (d[p + 1]! << 8);
const u32 = (d: Uint8Array, p: number): number =>
  (d[p]! | (d[p + 1]! << 8) | (d[p + 2]! << 16)) + d[p + 3]! * 0x1000000;
const u64 = (d: Uint8Array, p: number): number => u32(d, p) + u32(d, p + 4) * 0x100000000;

/**
 * Check the layout of the archive before fflate reads it: every entry must lie inside the input,
 * and no two entries may share bytes. fflate copies a stored entry by its compressed size and
 * ignores the declared size, so directory records that all point at one stored entry would
 * inflate the same bytes again and again. This walks the central directory exactly as fflate
 * does (same end record, same zip64 fields) and needs at least 46 bytes per record, so it is
 * linear in the input.
 */
function checkLayout(d: Uint8Array): ZipRejection | null {
  const n = d.length;
  if (n < 22) return 'inflate-error';
  let e = n - 22;
  while (u32(d, e) !== 0x06054b50) {
    e -= 1;
    if (e < 0 || n - e > 65558) return 'inflate-error';
  }
  let count = u16(d, e + 8);
  let offset = u32(d, e + 16);
  let zip64 = e >= 20 && u32(d, e - 20) === 0x07064b50;
  if (zip64) {
    const ze = u32(d, e - 12);
    zip64 = ze + 56 <= n && u32(d, ze) === 0x06064b50;
    if (zip64) {
      count = u32(d, ze + 32);
      offset = u32(d, ze + 48);
    }
  }
  const spans: [number, number][] = [];
  let compressedTotal = 0;
  for (let i = 0; i < count; i += 1) {
    if (offset + 46 > n || u32(d, offset) !== 0x02014b50) return 'inflate-error';
    const nameLength = u16(d, offset + 28);
    const extraLength = u16(d, offset + 30);
    const extraStart = offset + 46 + nameLength;
    let size = u32(d, offset + 20);
    let local = u32(d, offset + 42);
    if (
      zip64 &&
      (size === 0xffffffff || local === 0xffffffff || u32(d, offset + 24) === 0xffffffff)
    ) {
      const wide = zip64Fields(d, extraStart, extraLength, offset);
      if (!wide) return 'inflate-error';
      size = wide.size;
      local = wide.local;
    }
    offset = extraStart + extraLength + u16(d, offset + 32);
    if (local + 30 > n) return 'inflate-error';
    const dataStart = local + 30 + u16(d, local + 26) + u16(d, local + 28);
    if (dataStart + size > n) return 'inflate-error';
    compressedTotal += size;
    spans.push([local, dataStart + size]);
  }
  if (compressedTotal > n) return 'overlapping-entries';
  spans.sort((a, b) => a[0] - b[0]);
  for (let i = 1; i < spans.length; i += 1)
    if (spans[i]![0] < spans[i - 1]![1]) return 'overlapping-entries';
  return null;
}

/** The zip64 extra field values that fflate substitutes for 0xFFFFFFFF markers. */
function zip64Fields(
  d: Uint8Array,
  start: number,
  length: number,
  record: number
): { size: number; local: number } | null {
  const end = start + length;
  const sizeWide = u32(d, record + 20) === 0xffffffff;
  const originalWide = u32(d, record + 24) === 0xffffffff;
  const localWide = u32(d, record + 42) === 0xffffffff;
  for (let b = start; b + 4 < end; b += 4 + u16(d, b + 2)) {
    if (u16(d, b) !== 1) continue;
    const at = (slot: number): number => b + 4 + 8 * slot;
    const size = sizeWide ? u64(d, at(Number(originalWide))) : u32(d, record + 20);
    const local = localWide
      ? u64(d, at(Number(originalWide) + Number(sizeWide)))
      : u32(d, record + 42);
    return { size, local };
  }
  return null;
}

/**
 * Inflate a ZIP archive with bounds + OPC name normalization. Entry name, count, archive
 * ratio, and total-uncompressed-size limits are enforced BEFORE each entry is decompressed
 * (via fflate's pre-inflation filter), so a zip bomb or a traversal name is rejected without
 * ever being inflated. Keys are canonical part names.
 *
 * The budgets count each entry's DECLARED sizes, and the declared size is also a hard cap on
 * what the entry may produce: fflate inflates into a buffer of exactly `originalSize` bytes
 * and never grows it, so an entry whose stream holds more than it declares is truncated (or
 * refused as `inflate-error`), never expanded past its budget. A stored entry is copied at
 * its compressed size, so the larger of its two declared sizes is what counts. Before any of
 * this, the layout check refuses entries that lie outside the input or share bytes, so the
 * compressed sizes the budgets count are real and add up to at most the input length.
 */
export function readZip(bytes: Uint8Array, limits: ZipLimits = DEFAULT_ZIP_LIMITS): ZipReadResult {
  const caps = resolveZipLimits(limits);
  if (!caps) return { ok: false, reason: 'invalid-limits' };
  if (!(bytes instanceof Uint8Array)) return { ok: false, reason: 'inflate-error' };
  const layout = checkLayout(bytes);
  if (layout) return { ok: false, reason: layout };
  let entryCount = 0;
  let totalUncompressed = 0;
  let badDetail: string | undefined;
  const seenNorms = new Set<string>();

  let raw: Record<string, Uint8Array>;
  try {
    raw = unzipSync(bytes, {
      // `size` = compressed, `originalSize` = uncompressed. This runs BEFORE inflation.
      filter: (file) => {
        if (file.name.endsWith('/')) return false; // directory entry — never inflated
        entryCount += 1;
        if (entryCount > caps.maxEntries) throw new ZipViolation('too-many-entries');
        const norm = normalizePartName(file.name);
        if (!norm.ok) {
          badDetail = `${file.name}: ${norm.reason}`;
          throw new ZipViolation('bad-name');
        }
        // Reject two entries whose names are OPC-equivalent (normalized AND case-folded,
        // e.g. `word/document.xml` vs `word/%64ocument.xml` vs `Word/Document.xml`) —
        // last-wins would smuggle/omit a part. Checked BEFORE inflation.
        const key = partNameKey(norm.partName);
        if (seenNorms.has(key)) {
          badDetail = `${file.name}: OPC-equivalent duplicate of ${norm.partName}`;
          throw new ZipViolation('bad-name');
        }
        seenNorms.add(key);
        // Both budgets are checked before this entry is decompressed. A stored entry is
        // copied at `size` bytes whatever it declares as `originalSize`.
        totalUncompressed += Math.max(file.originalSize, file.size);
        if (totalUncompressed > caps.maxTotalBytes)
          throw new ZipViolation('too-large', 'zip.maxTotalBytes');
        if (
          totalUncompressed > ZIP_RATIO_EXEMPT_BYTES &&
          totalUncompressed / Math.max(1, bytes.byteLength) > caps.maxRatio
        )
          throw new ZipViolation('too-large', 'zip.maxRatio');
        return true;
      },
    });
  } catch (e) {
    if (e instanceof ZipViolation)
      return {
        ok: false,
        reason: e.reason,
        detail: badDetail,
        ...(e.limit === undefined ? {} : { limit: e.limit }),
      };
    return { ok: false, reason: 'inflate-error' };
  }

  const entries = new Map<string, Uint8Array>();
  for (const [name, data] of Object.entries(raw)) {
    const norm = normalizePartName(name);
    if (norm.ok) entries.set(norm.partName, data);
  }
  return { ok: true, entries };
}

/** The modification time stamped on every entry written.
 *
 *  Without one, a writer stamps `Date.now()`, which makes the same document save to
 *  different bytes depending on when it was saved: a ZIP entry carries its time in the DOS
 *  format, whose resolution is two seconds, so two saves that land either side of a tick
 *  differ in that field and nowhere else. Save/reopen/save then stops being a fixed point,
 *  and any digest a caller takes over saved bytes stops being stable.
 *
 *  The writer reads the DOS fields through the local-time getters (`getFullYear`, `getHours`),
 *  so the time is built in local time too: every time zone then writes the same fields, and an
 *  archive saved on one machine has the same bytes as on another. Mid-day on 1980-01-02 keeps
 *  the year inside the DOS range. */
const FIXED_ENTRY_MTIME = new Date(1980, 0, 2, 12);

const startsWith = (d: Uint8Array, at: number, sig: readonly number[]): boolean =>
  d.length >= at + sig.length && sig.every((b, i) => d[at + i] === b);

/**
 * Whether bytes are already compressed, by their signature: PNG, JPEG, GIF, JPEG XR, WebP, ZIP
 * (embedded OOXML packages), the ISO base media family (MP4, MOV, M4A), ASF (WMV, WMA), Ogg,
 * and MP3. Deflating these again costs time and saves nothing, so they are stored.
 */
export function isPrecompressed(d: Uint8Array): boolean {
  return (
    startsWith(d, 0, [0x89, 0x50, 0x4e, 0x47]) ||
    startsWith(d, 0, [0xff, 0xd8, 0xff]) ||
    startsWith(d, 0, [0x47, 0x49, 0x46, 0x38]) ||
    startsWith(d, 0, [0x49, 0x49, 0xbc]) ||
    (startsWith(d, 0, [0x52, 0x49, 0x46, 0x46]) && startsWith(d, 8, [0x57, 0x45, 0x42, 0x50])) ||
    startsWith(d, 0, [0x50, 0x4b, 0x03, 0x04]) ||
    startsWith(d, 4, [0x66, 0x74, 0x79, 0x70]) ||
    startsWith(d, 0, [0x30, 0x26, 0xb2, 0x75]) ||
    startsWith(d, 0, [0x4f, 0x67, 0x67, 0x53]) ||
    startsWith(d, 0, [0x49, 0x44, 0x33])
  );
}

/** Deflate a set of canonical-part-name -> bytes into a ZIP archive. Every part name
 *  is re-validated through the OPC normalization profile before writing, so a
 *  traversal/encoded/normalized-alias name from an untrusted serialized model can
 *  never be smuggled into a ZIP entry (write-side path-traversal guard). The name is
 *  encoded, never decoded again: the entry reads back as exactly the same part name.
 *  Already-compressed media is stored, not deflated a second time. The archive stays within
 *  the ratio cap of {@link DEFAULT_ZIP_LIMITS}, so a document that opened opens again. */
export function writeZip(entries: ReadonlyMap<string, Uint8Array>): Uint8Array {
  const steps = writeZipSteps(entries);
  for (;;) {
    const next = steps.next();
    if (next.done) return next.value;
  }
}

/**
 * {@link writeZip}, yielding between slices of a long entry. The archive is the same either
 * way.
 *
 * @internal
 */
export function writeZipSteps(
  entries: ReadonlyMap<string, Uint8Array>
): Generator<void, Uint8Array> {
  const named = new Map<string, Uint8Array>();
  const seen = new Set<string>();
  for (const [partName, data] of entries) {
    const item = zipItemName(partName);
    if (!item.ok) throw new Error(`unsafe part name on write: ${partName} (${item.reason})`);
    const fold = partNameKey(partName); // OPC case-folded equivalence
    if (seen.has(fold)) throw new Error(`OPC-equivalent duplicate part name on write: ${partName}`);
    seen.add(fold);
    // Canonical part names carry a leading slash; ZIP entry names do not.
    named.set(item.itemName, data);
  }
  return writeReadable(named);
}

/** The compressed and uncompressed size of each entry of an archive this module wrote. */
function writtenSizes(zip: Uint8Array): [compressed: number, size: number][] {
  const end = zip.length - 22;
  let at = u32(zip, end + 16);
  const sizes: [number, number][] = [];
  for (let i = u16(zip, end + 10); i > 0; i -= 1) {
    sizes.push([u32(zip, at + 20), u32(zip, at + 24)]);
    at += 46 + u16(zip, at + 28) + u16(zip, at + 30) + u16(zip, at + 32);
  }
  return sizes;
}

/**
 * Write `named` within the ratio cap of {@link DEFAULT_ZIP_LIMITS}, so that {@link readZip}
 * reads it back as it read the document that was opened.
 *
 * Deflate shrinks a part of one repeated byte about a thousand times, so a document that opened
 * with such a part stored (a blank bitmap, say) would save to an archive past the ratio cap and
 * never open again. When the archive is past the cap, the entries that shrank the most are
 * stored instead, as few as bring it under the cap, and the archive is written once more.
 */
function* writeReadable(entries: ReadonlyMap<string, Uint8Array>): Generator<void, Uint8Array> {
  // One view object per entry, so storing one entry never stores another that shares its bytes.
  const named = new Map<string, Uint8Array>();
  const views = new Set<Uint8Array>();
  for (const [name, data] of entries) {
    const view = views.has(data) ? data.subarray() : data;
    views.add(view);
    named.set(name, view);
  }
  const forced = new Set<Uint8Array>();
  const store = (data: Uint8Array): boolean => forced.has(data) || isPrecompressed(data);
  const maxRatio = DEFAULT_ZIP_LIMITS.maxRatio ?? 200;
  for (;;) {
    const zip = yield* writeZipEntriesSteps(named, FIXED_ENTRY_MTIME, undefined, store);
    const sizes = writtenSizes(zip);
    let counted = 0;
    for (const [compressed, size] of sizes) counted += Math.max(compressed, size);
    if (counted <= ZIP_RATIO_EXEMPT_BYTES || counted / Math.max(1, zip.length) <= maxRatio)
      return zip;
    // Storing an entry grows the archive by what deflate saved on it; take the biggest savings.
    const savings: [gain: number, data: Uint8Array][] = [];
    let index = 0;
    for (const data of named.values()) {
      const [compressed, size] = sizes[index]!;
      index += 1;
      if (!store(data) && size > compressed) savings.push([size - compressed, data]);
    }
    if (savings.length === 0) return zip;
    savings.sort((a, b) => b[0] - a[0]);
    let missing = counted / maxRatio - zip.length;
    for (const [gain, data] of savings) {
      forced.add(data);
      missing -= gain;
      if (missing < 0) break;
    }
  }
}

export { strToU8, strFromU8 } from 'fflate';
