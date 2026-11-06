// Harness-only font discovery. The fixture set loads only `fixtures/fonts/`: open fonts that are
// in the repository, so a fixture renders with no system font (only these and the fonts
// that ship with the engine). The corpus set loads the font folders of `PPTX_HARNESS_FONT_DIRS`
// (separated by `:`, in priority order; a file name is read once): the reference fonts of
// private decks, read at run time and NEVER copied into the repo. Arial, the fallback that every
// deck gets (`deckFontFamilies`), must be in one of these folders. A font collection (`.ttc`,
// `.otc`) gives one font for each of its faces.

import { strFromU8, unzipSync } from 'fflate';
import { createHash } from 'node:crypto';
import { existsSync, readdirSync, readFileSync, statSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';

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

/** The corpus font folders, from `PPTX_HARNESS_FONT_DIRS`. Empty on a machine without them. */
const DIRS = (process.env.PPTX_HARNESS_FONT_DIRS ?? '')
  .split(':')
  .map((dir) => dir.trim())
  .filter(Boolean);

/** The open fonts of the fixture set, in the repository (`fixtures/fonts/`, with their licenses). */
export const FIXTURE_FONTS_DIR = resolve(import.meta.dir, '../../fixtures/fonts');

/** The fonts that ship with the engine (`@docx-editor.dev/fonts`): the page always has them. */
export const ENGINE_FONTS_DIR = join(
  dirname(
    Bun.resolveSync(
      '@docx-editor.dev/fonts/package.json',
      resolve(import.meta.dir, '../../packages/core')
    )
  ),
  'assets'
);

/**
 * The faces that the harness page draws with where the painted SVG names the generic
 * `sans-serif` family: an open text face, then an open symbol face. So a glyph that a deck's face
 * lacks never takes a system font, as long as these hold it (`examples/vite/src/harness.ts`).
 */
export const FALLBACK_FONTS: readonly string[] = [
  join(ENGINE_FONTS_DIR, 'LiberationSans-Regular.ttf'),
  resolve(import.meta.dir, '../../fixtures/fonts/NotoSansSymbols2-Regular.ttf'),
];

/**
 * The fixtures whose decks name fonts that are not open, on purpose, with the reason. Their
 * reference PDFs may embed any font. Every other fixture names only open fonts, and its
 * reference PDF embeds only the open fonts that its deck names (`foreignReferenceFonts`).
 */
export const NON_OPEN_FONT_DECKS: Readonly<Record<string, string>> = {
  'font-substitution':
    'it checks which open face the engine draws for each face that is not open, symbol faces included',
};

/** Noto Sans CJK JP of `@docx-editor.dev/fonts-cjk`, part of the fixture set. */
const CJK_FACE = join(
  dirname(
    Bun.resolveSync('@docx-editor.dev/fonts-cjk/package.json', resolve(import.meta.dir, '../..'))
  ),
  'assets',
  'NotoSansCJKjp-Regular.otf'
);

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

/**
 * Family (name ID 1), subfamily (2), and PostScript name (6), US English, from the `name` table
 * of the face at `base`.
 */
function names(
  bytes: Uint8Array,
  base = 0
): { family?: string; sub?: string; postScript?: string } {
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
    const out: { family?: string; sub?: string; postScript?: string } = {};
    for (let j = 0; j < count; j += 1) {
      const r = off + 6 + j * 12;
      const platform = v.getUint16(r);
      const id = v.getUint16(r + 6);
      const len = v.getUint16(r + 8);
      const start = strings + v.getUint16(r + 10);
      const language = v.getUint16(r + 4);
      if (platform !== 3 || language !== 0x409 || (id !== 1 && id !== 2 && id !== 6)) continue;
      let s = '';
      for (let k = 0; k + 1 < len; k += 2) s += String.fromCharCode(v.getUint16(start + k));
      if (id === 1) out.family ??= s;
      else if (id === 2) out.sub ??= s;
      else out.postScript ??= s;
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

/** The font sets: `corpus` is this machine's reference fonts, `fixtures` the repository's. */
export type FontSet = 'corpus' | 'fixtures';

const indexes = new Map<FontSet, FontFile[]>();

/**
 * The font files of a set: each font file name once. The fixture set reads `FIXTURE_FONTS_DIR`
 * and the CJK face; the corpus set reads the `PPTX_HARNESS_FONT_DIRS` folders.
 */
function fontFiles(set: FontSet = 'corpus'): string[] {
  // The first two corpus folders can hold the same font files (about 500 MB each): read each
  // file name once.
  const files: string[] = [];
  const seen = new Set<string>();
  for (const dir of set === 'fixtures' ? [FIXTURE_FONTS_DIR] : DIRS)
    if (existsSync(dir))
      for (const f of readdirSync(dir)) {
        if (seen.has(f.toLowerCase())) continue;
        seen.add(f.toLowerCase());
        files.push(join(dir, f));
      }
  // The CJK face that a host supplies (`@docx-editor.dev/fonts-cjk`): without it, East Asian
  // text paints with whatever CJK font the system has, and scores differ between machines.
  if (set === 'fixtures') files.push(CJK_FACE);
  return files.filter((file) => /\.(ttf|otf|ttc|otc)$/i.test(file));
}

/** Family and style of every font file of a set, read once per process. */
function fontIndex(set: FontSet): FontFile[] {
  let index = indexes.get(set);
  if (index) return index;
  index = [];
  indexes.set(set, index);
  for (const file of fontFiles(set)) {
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

/**
 * Every font of a set that the harness may use (the corpus set by default). Families in `only`
 * limit the set (by lower-case name).
 */
export function localFonts(only?: ReadonlySet<string>, set: FontSet = 'corpus'): LocalFont[] {
  return fontIndex(set)
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
 * theme fonts included, plus Arial for fallback (unless `withFallback` is false). The harness
 * sends only these to the page, because the full reference font set is about 750 MB.
 */
export function deckFontFamilies(deck: Uint8Array, withFallback = true): Set<string> {
  const out = new Set<string>(withFallback ? ['arial'] : []);
  for (const [name, bytes] of Object.entries(unzipSync(deck))) {
    if (!name.endsWith('.xml')) continue;
    for (const m of strFromU8(bytes).matchAll(/typeface="([^"]+)"/g)) {
      const family = m[1];
      if (!family.startsWith('+')) out.add(family.toLowerCase());
    }
  }
  return out;
}

/**
 * The fixture fonts, for `FIXTURE_KNOWN_GAPS.fonts`: the versions of the font packages that ship
 * with the engine, and a SHA-256 over the files of `FIXTURE_FONTS_DIR`. Scores depend on these
 * bytes, so a change to them measures the list again.
 */
export function fixtureFontsIdentity(): string {
  const version = (pkg: string): string =>
    (
      JSON.parse(
        readFileSync(
          Bun.resolveSync(`${pkg}/package.json`, resolve(import.meta.dir, '../../packages/core')),
          'utf8'
        )
      ) as { version: string }
    ).version;
  const hash = createHash('sha256');
  for (const file of readdirSync(FIXTURE_FONTS_DIR).sort())
    hash
      .update(file)
      .update('\0')
      .update(readFileSync(join(FIXTURE_FONTS_DIR, file)));
  return `@docx-editor.dev/fonts ${version('@docx-editor.dev/fonts')}, @docx-editor.dev/fonts-cjk ${version('@docx-editor.dev/fonts-cjk')}, fixtures/fonts ${hash.digest('hex').slice(0, 12)}`;
}

/** The family (lower case) of each face of each file, by PostScript name; read once per file. */
const postScriptCache = new Map<string, readonly (readonly [string, string])[]>();

/** The family (lower case) of every face of the font files, by PostScript name. */
export function postScriptFamilies(files: readonly string[]): Map<string, string> {
  const out = new Map<string, string>();
  for (const file of files) {
    let faces = postScriptCache.get(file);
    if (!faces) {
      const bytes = new Uint8Array(readFileSync(file));
      faces = faceOffsets(bytes).flatMap((base) => {
        const { family, postScript } = names(bytes, base);
        return family && postScript ? [[postScript, family.toLowerCase()] as const] : [];
      });
      postScriptCache.set(file, faces);
    }
    for (const [postScript, family] of faces) out.set(postScript, family);
  }
  return out;
}

/** The engine's fonts, the fallback fonts, and `files`: every font the harness page can draw with. */
export function harnessFontFiles(files: readonly string[]): string[] {
  const engine = readdirSync(ENGINE_FONTS_DIR)
    .filter((f) => /\.(ttf|otf)$/i.test(f))
    .map((f) => join(ENGINE_FONTS_DIR, f));
  return [...new Set([...engine, ...FALLBACK_FONTS, ...files])];
}

/** A font of a PDF, as `pdf.ts` lists it. */
interface PdfFontName {
  readonly name: string;
  readonly embedded: boolean;
}

/**
 * The fonts of our PDF that are not faces the harness page loaded (`harnessFontFiles`), or not
 * embedded: a system font took part in the paint. A font is known by its PostScript name only:
 * a system font with the PostScript name of a loaded face passes unseen.
 */
export function foreignOurFonts(
  fonts: readonly PdfFontName[],
  loaded: ReadonlyMap<string, string>
): string[] {
  return fonts
    .filter((f) => !f.embedded || !loaded.has(f.name))
    .map((f) => (f.embedded ? f.name : `${f.name} (not embedded)`));
}

/**
 * The fonts of a fixture's reference PDF that are not open fonts that its deck names: each one
 * must be a face of `open` whose family is in `named`, and embedded. Empty for a deck in
 * `NON_OPEN_FONT_DECKS`.
 */
export function foreignReferenceFonts(
  deck: string,
  fonts: readonly PdfFontName[],
  named: ReadonlySet<string>,
  open: ReadonlyMap<string, string>
): string[] {
  if (Object.hasOwn(NON_OPEN_FONT_DECKS, deck)) return [];
  return fonts
    .filter((f) => {
      const family = open.get(f.name);
      return !f.embedded || family === undefined || !named.has(family);
    })
    .map((f) => (f.embedded ? f.name : `${f.name} (not embedded)`));
}
