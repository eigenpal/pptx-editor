// Embedded OpenType (EOT) header parsing.
// Implements W3C Member Submission "Embedded OpenType (EOT) File Format" (2008):
//   3.1-3.3 EMBEDDEDFONT structure versions 0x00010000, 0x00020001, 0x00020002
//   4.2     processing flags (TTEMBED_TTCOMPRESSED, TTEMBED_XORENCRYPTDATA, ...)
//   4.3.2   RootString checksum (reported, not enforced)
//   4.4     XOR de-obfuscation (key 0x50) applied before MTX decompression
// All EMBEDDEDFONT fields are little-endian; FontData is opaque here.

import { FontFormatError, reject, type FontRejection } from './types.ts';

export const EOT_MAGIC = 0x504c;
export const EOT_VERSION_1 = 0x00010000;
export const EOT_VERSION_2_1 = 0x00020001;
export const EOT_VERSION_2_2 = 0x00020002;

export const TTEMBED_SUBSET = 0x00000001;
export const TTEMBED_TTCOMPRESSED = 0x00000004;
export const TTEMBED_EMBEDEUDC = 0x00000020;
export const TTEMBED_XORENCRYPTDATA = 0x10000000;

const XOR_KEY = 0x50;
const ROOT_STRING_XOR = 0x50475342;

export interface EotHeader {
  version: number;
  eotSize: number;
  flags: number;
  panose: Uint8Array;
  charset: number;
  italic: boolean;
  weight: number;
  fsType: number;
  checkSumAdjustment: number;
  familyName: string;
  styleName: string;
  versionName: string;
  fullName: string;
  rootString: string;
  /** null when the version has no RootStringCheckSum field. */
  rootStringChecksumValid: boolean | null;
  eudcCodePage: number;
  /** TTEMBED_TTCOMPRESSED: FontData is MicroType Express compressed. */
  compressed: boolean;
  /** TTEMBED_SUBSET: font was subsetted when embedded. */
  subset: boolean;
  /** TTEMBED_XORENCRYPTDATA: FontData was XOR-obfuscated (already undone in fontData). */
  xorEncrypted: boolean;
}

export interface EotParseSuccess {
  ok: true;
  header: EotHeader;
  /** FontData with XOR obfuscation removed; still MTX-compressed if header.compressed. */
  fontData: Uint8Array;
}

export type EotParseResult = EotParseSuccess | FontRejection;

class LeReader {
  private readonly d: Uint8Array;
  pos = 0;
  private readonly end: number;

  constructor(d: Uint8Array, end: number) {
    this.d = d;
    this.end = end;
  }

  need(n: number): void {
    if (n < 0 || this.pos + n > this.end) {
      throw new FontFormatError('truncated', `EOT header field at ${this.pos} overruns data`);
    }
  }

  u8(): number {
    this.need(1);
    return this.d[this.pos++];
  }

  u16(): number {
    this.need(2);
    const v = this.d[this.pos] | (this.d[this.pos + 1] << 8);
    this.pos += 2;
    return v;
  }

  u32(): number {
    this.need(4);
    const p = this.pos;
    this.pos += 4;
    return (this.d[p] | (this.d[p + 1] << 8) | (this.d[p + 2] << 16)) + this.d[p + 3] * 0x1000000;
  }

  bytes(n: number): Uint8Array {
    this.need(n);
    const out = this.d.subarray(this.pos, this.pos + n);
    this.pos += n;
    return out;
  }

  /** Padding u16 + size u16 + UTF-16LE bytes. Padding is not enforced (writers vary). */
  sizedString(): Uint8Array {
    this.u16();
    return this.bytes(this.u16());
  }
}

function utf16le(bytes: Uint8Array): string {
  let s = '';
  for (let i = 0; i + 1 < bytes.length; i += 2) {
    const c = bytes[i] | (bytes[i + 1] << 8);
    if (c === 0) break;
    s += String.fromCharCode(c);
  }
  return s;
}

function rootChecksum(bytes: Uint8Array): number {
  let sum = 0;
  for (const b of bytes) sum = (sum + b) >>> 0;
  return (sum ^ ROOT_STRING_XOR) >>> 0;
}

/** Parses an EOT (EMBEDDEDFONT) structure. Never throws. */
export function parseEot(bytes: Uint8Array): EotParseResult {
  try {
    return parseEotUnsafe(bytes);
  } catch (err) {
    if (err instanceof FontFormatError) return reject(err.reason, err.message);
    return reject('truncated', 'unexpected EOT parse failure');
  }
}

function parseEotUnsafe(bytes: Uint8Array): EotParseResult {
  if (bytes.length < 82) return reject('truncated', 'shorter than the fixed EOT header');
  const head = new LeReader(bytes, bytes.length);
  const eotSize = head.u32();
  const fontDataSize = head.u32();
  const version = head.u32();
  const flags = head.u32();
  if (version !== EOT_VERSION_1 && version !== EOT_VERSION_2_1 && version !== EOT_VERSION_2_2) {
    return reject('unsupported-version', `EOT version 0x${version.toString(16)}`);
  }
  if (eotSize > bytes.length) return reject('truncated', 'EOTSize exceeds buffer length');
  if (eotSize < 82) return reject('size-mismatch', 'EOTSize smaller than header');
  // Re-bound every subsequent read by the declared structure size.
  const r = new LeReader(bytes, eotSize);
  r.pos = 16;
  const panose = r.bytes(10).slice();
  const charset = r.u8();
  const italic = r.u8() === 1;
  const weight = r.u32();
  const fsType = r.u16();
  const magic = r.u16();
  if (magic !== EOT_MAGIC) return reject('bad-magic', `magic 0x${magic.toString(16)}`);
  r.bytes(16); // UnicodeRange1-4
  r.bytes(8); // CodePageRange1-2
  const checkSumAdjustment = r.u32();
  r.bytes(16); // Reserved1-4
  const familyName = utf16le(r.sizedString());
  const styleName = utf16le(r.sizedString());
  const versionName = utf16le(r.sizedString());
  const fullName = utf16le(r.sizedString());

  let rootString = '';
  let rootStringChecksumValid: boolean | null = null;
  let eudcCodePage = 0;
  if (version !== EOT_VERSION_1) {
    const rootBytes = r.sizedString();
    rootString = utf16le(rootBytes).replace(/\0/g, ' ').trim();
    if (version === EOT_VERSION_2_2) {
      const stored = r.u32();
      rootStringChecksumValid = stored === rootChecksum(rootBytes);
      eudcCodePage = r.u32();
      r.sizedString(); // Padding6 + Signature (reserved)
      r.u32(); // EUDCFlags
      const eudcSize = r.u32();
      r.bytes(eudcSize); // EUDC font data is not used for rendering.
    }
  }

  const fontStart = r.pos;
  if (fontDataSize > eotSize - fontStart) {
    return reject('size-mismatch', 'FontDataSize exceeds EOTSize');
  }
  let fontData = bytes.subarray(fontStart, fontStart + fontDataSize);
  const xorEncrypted = (flags & TTEMBED_XORENCRYPTDATA) !== 0;
  if (xorEncrypted) {
    const plain = new Uint8Array(fontData.length);
    for (let i = 0; i < fontData.length; i++) plain[i] = fontData[i] ^ XOR_KEY;
    fontData = plain;
  }
  return {
    ok: true,
    header: {
      version,
      eotSize,
      flags,
      panose,
      charset,
      italic,
      weight,
      fsType,
      checkSumAdjustment,
      familyName,
      styleName,
      versionName,
      fullName,
      rootString,
      rootStringChecksumValid,
      eudcCodePage,
      compressed: (flags & TTEMBED_TTCOMPRESSED) !== 0,
      subset: (flags & TTEMBED_SUBSET) !== 0,
      xorEncrypted,
    },
    fontData,
  };
}
