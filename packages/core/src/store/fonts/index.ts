// Public API for decoding embedded deck fonts (.fntdata) into plain sfnt bytes.
// Pipeline: EOT header (W3C EOT Submission section 3) -> XOR de-obfuscation (4.4) ->
// MicroType Express decompression (W3C MTX Submission sections 2-7) -> sfnt validation.
// Embedding permissions follow OpenType OS/2 fsType and EOT Submission section 4.1.

import { parseEot } from './eot.ts';
import { decodeMtx } from './mtx.ts';
import { readSfntName, readTableDirectory, readU16, validateSfnt } from './sfnt.ts';
import {
  FontFormatError,
  reject,
  resolveLimits,
  type DecodeLimits,
  type FontRejection,
} from './types.ts';

export type { DecodeLimits, FontRejection, FontRejectReason } from './types.ts';
export { DEFAULT_LIMITS } from './types.ts';
export { validateSfnt, type SfntValidation } from './sfnt.ts';
export { parseEot, type EotHeader, type EotParseResult } from './eot.ts';

export type EmbeddingPermission = 'installable' | 'editable' | 'preview-print' | 'restricted';

/**
 * Maps OS/2 fsType to an embedding level. Per EOT section 4.1, when legacy fonts set several
 * of bits 0-3 the least restrictive one wins (editable > preview & print > restricted).
 * Bit 9 (bitmap embedding only) means no outlines may be used, which this engine treats as
 * 'restricted' because it renders outlines only. Reserved bits are ignored.
 */
export function fontEmbeddingPermission(fsType: number): EmbeddingPermission {
  const v = fsType & 0xffff;
  if (v & 0x0200) return 'restricted';
  if (v & 0x0008) return 'editable';
  if (v & 0x0004) return 'preview-print';
  if (v & 0x0002) return 'restricted';
  return 'installable';
}

export interface EmbeddedFontInfo {
  familyName: string;
  styleName: string;
  fullName: string;
  versionName: string;
  weight: number;
  italic: boolean;
  /** OS/2 fsType copy from the EOT header. */
  fsType: number;
  permission: EmbeddingPermission;
  /** EOT processing flags. */
  flags: number;
  charset: number;
  eotVersion: number;
  compressed: boolean;
  subset: boolean;
  xorEncrypted: boolean;
  numGlyphs: number;
  /** Non-fatal observations (e.g. dropped optional tables, unconsumed stream bytes). */
  warnings: string[];
}

export type DecodeEmbeddedFontResult =
  { ok: true; sfnt: Uint8Array; info: EmbeddedFontInfo } | FontRejection;

/**
 * Decodes an EOT (as found in PPTX `ppt/fonts/*.fntdata`) into TrueType/OpenType sfnt bytes.
 * Never throws: malformed or hostile input yields `{ ok: false, reason, detail }`.
 */
export function decodeEmbeddedFont(
  bytes: Uint8Array,
  limits?: Partial<DecodeLimits>
): DecodeEmbeddedFontResult {
  const lim = resolveLimits(limits);
  try {
    if (!(bytes instanceof Uint8Array)) return reject('truncated', 'input is not a Uint8Array');
    if (bytes.length > lim.maxInputBytes) return reject('limit-exceeded', 'input too large');
    const eot = parseEot(bytes);
    if (!eot.ok) return eot;
    const { header, fontData } = eot;
    const warnings: string[] = [];
    let sfnt: Uint8Array;
    let numGlyphs: number;

    if (header.compressed) {
      const mtx = decodeMtx(fontData, lim);
      sfnt = mtx.sfnt;
      numGlyphs = mtx.numGlyphs;
      for (const w of mtx.warnings) warnings.push(w);
      const check = validateSfnt(sfnt);
      if (!check.ok) return reject('invalid-sfnt', check.errors.join('; '));
    } else {
      if (fontData.length > lim.maxOutputBytes) return reject('limit-exceeded', 'font too large');
      const dir = readTableDirectory(fontData, 0, 'invalid-sfnt');
      for (const t of dir.tables) {
        if (t.offset > fontData.length || t.length > fontData.length - t.offset) {
          return reject('invalid-sfnt', `table '${t.tag}' out of bounds`);
        }
      }
      const maxp = dir.tables.find((t) => t.tag === 'maxp');
      numGlyphs = maxp && maxp.length >= 6 ? readU16(fontData, maxp.offset + 4) : 0;
      if (numGlyphs > lim.maxGlyphs) return reject('limit-exceeded', 'too many glyphs');
      const check = validateSfnt(fontData);
      if (!check.ok) for (const e of check.errors) warnings.push(e);
      sfnt = fontData.slice();
    }
    if (header.rootStringChecksumValid === false) warnings.push('RootString checksum mismatch');

    // Some writers leave the EOT name fields empty; fall back to the font's own name table.
    const nameOr = (value: string, nameId: number): string =>
      value.trim() !== '' ? value : readSfntName(sfnt, nameId);
    return {
      ok: true,
      sfnt,
      info: {
        familyName: nameOr(nameOr(header.familyName, 1), 16),
        styleName: nameOr(header.styleName, 2),
        fullName: nameOr(header.fullName, 4),
        versionName: header.versionName,
        weight: header.weight,
        italic: header.italic,
        fsType: header.fsType,
        permission: fontEmbeddingPermission(header.fsType),
        flags: header.flags,
        charset: header.charset,
        eotVersion: header.version,
        compressed: header.compressed,
        subset: header.subset,
        xorEncrypted: header.xorEncrypted,
        numGlyphs,
        warnings,
      },
    };
  } catch (err) {
    if (err instanceof FontFormatError) return reject(err.reason, err.message);
    if (err instanceof RangeError) return reject('limit-exceeded', 'allocation failed');
    return reject('invalid-sfnt', 'unexpected decoder failure');
  }
}
