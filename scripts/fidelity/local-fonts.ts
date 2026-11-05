// Harness-only font discovery: the reference fonts, which are the font folders the harness may
// load, from `PPTX_HARNESS_FONT_DIRS` (separated by `:`, in priority order; a file name is read
// once), plus macOS system Arial. Read at run time from this machine; NEVER copied into the repo.
// A font collection (`.ttc`, `.otc`) gives one font for each of its faces.

import { strFromU8, unzipSync } from 'fflate';
import { createHash } from 'node:crypto';
import { existsSync, readdirSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';

export interface LocalFont {
  readonly family: string;
  readonly bold: boolean;
  readonly italic: boolean;
  /** The font file. The harness page fetches it from the harness server (`openDeck`). */
  readonly file: string;
  /**
   * The face in a font collection. The harness server then gives the page that face alone, as
   * a single font (`collectionFace`). Unset for a single font file.
   */
  readonly faceIndex?: number;
}

/** The most faces the harness reads from one font collection. */
export const MAX_COLLECTION_FACES = 64;
/** The most tables that one face may list. */
const MAX_TABLES = 256;

/**
 * The offset of each face's table directory: `[0]` for a single font, one for each face of a
 * collection (OpenType, "Font collections": `ttcf`, `numFonts`, `tableDirectoryOffsets`), at
 * most `MAX_COLLECTION_FACES`. Empty when the header is cut short.
 */
export function faceOffsets(bytes: Uint8Array): number[] {
  if (bytes.length < 12) return [];
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (v.getUint32(0) !== 0x74746366) return [0];
  const count = Math.min(v.getUint32(8), MAX_COLLECTION_FACES);
  const out: number[] = [];
  for (let i = 0; i < count && 16 + i * 4 <= bytes.length; i += 1)
    out.push(v.getUint32(12 + i * 4));
  return out;
}

/**
 * One face of a font collection as a single sfnt font: its table directory and its tables,
 * copied to new offsets. A browser `FontFace` and the HarfBuzz measurer read the first face of
 * the bytes they get, so the page gets the face alone. Undefined when the face or one of its
 * tables lies outside the file.
 */
export function collectionFace(bytes: Uint8Array, faceIndex: number): Uint8Array | undefined {
  const base = faceOffsets(bytes)[faceIndex];
  if (base === undefined || base + 12 > bytes.length) return undefined;
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const numTables = v.getUint16(base + 4);
  if (numTables > MAX_TABLES || base + 12 + numTables * 16 > bytes.length) return undefined;
  const tables: { record: number; offset: number; length: number }[] = [];
  let size = 12 + numTables * 16;
  for (let i = 0; i < numTables; i += 1) {
    const record = base + 12 + i * 16;
    const offset = v.getUint32(record + 8);
    const length = v.getUint32(record + 12);
    if (offset + length > bytes.length) return undefined;
    tables.push({ record, offset, length });
    size += (length + 3) & ~3;
  }
  const out = new Uint8Array(size);
  const w = new DataView(out.buffer);
  out.set(bytes.subarray(base, base + 12), 0);
  let at = 12 + numTables * 16;
  tables.forEach((t, i) => {
    const record = 12 + i * 16;
    out.set(bytes.subarray(t.record, t.record + 8), record);
    w.setUint32(record + 8, at);
    w.setUint32(record + 12, t.length);
    out.set(bytes.subarray(t.offset, t.offset + t.length), at);
    at += (t.length + 3) & ~3;
  });
  return out;
}

/** The harness font folders, from `PPTX_HARNESS_FONT_DIRS`. Empty on a machine without them. */
const DIRS = (process.env.PPTX_HARNESS_FONT_DIRS ?? '')
  .split(':')
  .map((dir) => dir.trim())
  .filter(Boolean);
const SYSTEM_ARIAL = '/System/Library/Fonts/Supplemental';

/** Bold and italic bits of `head.macStyle`, for the face whose table directory is at `base`. */
function macStyle(bytes: Uint8Array, base = 0): { bold: boolean; italic: boolean } {
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const numTables = bytes.length >= base + 12 ? v.getUint16(base + 4) : 0;
  for (let i = 0; i < numTables; i += 1) {
    const rec = base + 12 + i * 16;
    if (rec + 16 > bytes.length) break;
    const tag = String.fromCharCode(bytes[rec], bytes[rec + 1], bytes[rec + 2], bytes[rec + 3]);
    if (tag !== 'head') continue;
    const style = v.getUint16(v.getUint32(rec + 8) + 44);
    return { bold: (style & 1) !== 0, italic: (style & 2) !== 0 };
  }
  return { bold: false, italic: false };
}

/** Family (name ID 1, US English) from the `name` table of the face at `base`. */
function names(bytes: Uint8Array, base = 0): { family?: string; sub?: string } {
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.length < base + 12) return {};
  const numTables = v.getUint16(base + 4);
  for (let i = 0; i < numTables; i += 1) {
    const rec = base + 12 + i * 16;
    if (rec + 16 > bytes.length) break;
    const tag = String.fromCharCode(bytes[rec], bytes[rec + 1], bytes[rec + 2], bytes[rec + 3]);
    if (tag !== 'name') continue;
    const off = v.getUint32(rec + 8);
    const count = v.getUint16(off + 2);
    const strings = off + v.getUint16(off + 4);
    const out: { family?: string; sub?: string } = {};
    for (let j = 0; j < count; j += 1) {
      const r = off + 6 + j * 12;
      const platform = v.getUint16(r);
      const id = v.getUint16(r + 6);
      const len = v.getUint16(r + 8);
      const start = strings + v.getUint16(r + 10);
      const language = v.getUint16(r + 4);
      if (platform !== 3 || language !== 0x409 || (id !== 1 && id !== 2)) continue;
      let s = '';
      for (let k = 0; k + 1 < len; k += 2) s += String.fromCharCode(v.getUint16(start + k));
      if (id === 1) out.family ??= s;
      else out.sub ??= s;
    }
    return out;
  }
  return {};
}

interface FontFile {
  readonly file: string;
  readonly faceIndex?: number;
  readonly family: string;
  readonly bold: boolean;
  readonly italic: boolean;
}

let index: FontFile[] | undefined;

/** The local font files the harness may load: each font file name once, and the system Arial. */
function fontFiles(): string[] {
  // The first two folders can hold the same font files (about 500 MB each): read each file name
  // once.
  const files: string[] = [];
  const seen = new Set<string>();
  for (const dir of DIRS)
    if (existsSync(dir))
      for (const f of readdirSync(dir)) {
        if (seen.has(f.toLowerCase())) continue;
        seen.add(f.toLowerCase());
        files.push(join(dir, f));
      }
  if (existsSync(SYSTEM_ARIAL))
    for (const f of readdirSync(SYSTEM_ARIAL))
      if (/^Arial( Bold| Italic| Bold Italic)?\.ttf$/.test(f)) files.push(join(SYSTEM_ARIAL, f));
  return files.filter((file) => /\.(ttf|otf|ttc|otc)$/i.test(file));
}

/** Family and style of every local font file, read once per process. */
function fontIndex(): FontFile[] {
  if (index) return index;
  index = [];
  for (const file of fontFiles()) {
    const bytes = new Uint8Array(readFileSync(file));
    const collection = strFromU8(bytes.subarray(0, 4)) === 'ttcf';
    faceOffsets(bytes).forEach((base, faceIndex) => {
      const { family } = names(bytes, base);
      if (family)
        index?.push({
          file,
          ...(collection ? { faceIndex } : {}),
          family,
          ...macStyle(bytes, base),
        });
    });
  }
  return index;
}

/** Whether the reference fonts are installed (the first two `PPTX_HARNESS_FONT_DIRS` folders). */
export function referenceFontsInstalled(): boolean {
  return DIRS.slice(0, 2).some((dir) => existsSync(dir));
}

/**
 * A SHA-256 over the name, size, and time of every local font file the harness may load. It
 * changes when a font file is added, removed, or replaced, so a corpus resume cannot mix two
 * font sets. It reads no font bytes, so it is fast.
 */
export function localFontsIdentity(): string {
  const entries = fontFiles()
    .map((file) => {
      const stat = statSync(file);
      return `${file}\0${stat.size}\0${Math.round(stat.mtimeMs)}`;
    })
    .sort();
  return createHash('sha256').update(entries.join('\n')).digest('hex');
}

/** Every local font the harness may use. Families in `only` limit the set (by lower-case name). */
export function localFonts(only?: ReadonlySet<string>): LocalFont[] {
  return fontIndex()
    .filter((f) => !only || only.has(f.family.toLowerCase()))
    .map((f) => ({
      family: f.family,
      bold: f.bold,
      italic: f.italic,
      file: f.file,
      ...(f.faceIndex === undefined ? {} : { faceIndex: f.faceIndex }),
    }));
}

/**
 * The font families a deck names (lower case): every `typeface` attribute in its XML parts,
 * theme fonts included, plus Arial for fallback. The harness sends only these to the page,
 * because the full reference font set is about 750 MB.
 */
export function deckFontFamilies(deck: Uint8Array): Set<string> {
  const out = new Set<string>(['arial']);
  for (const [name, bytes] of Object.entries(unzipSync(deck))) {
    if (!name.endsWith('.xml')) continue;
    for (const m of strFromU8(bytes).matchAll(/typeface="([^"]+)"/g)) {
      const family = m[1];
      if (!family.startsWith('+')) out.add(family.toLowerCase());
    }
  }
  return out;
}
