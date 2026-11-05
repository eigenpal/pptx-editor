// Shared result/rejection types for the embedded-font decoder.
// Rejections follow the "never throw on untrusted input" rule: every public entry point
// returns `{ ok: false, reason, detail }` instead of raising.

export type FontRejectReason =
  /** Input ends before a structure it declares. */
  | 'truncated'
  /** EOT MagicNumber is not 0x504C. */
  | 'bad-magic'
  /** EOT version or MTX version is not one we implement. */
  | 'unsupported-version'
  /** Declared sizes disagree with each other or with the buffer. */
  | 'size-mismatch'
  /** MicroType Express container header is invalid (MTX spec section 2.1). */
  | 'invalid-mtx-header'
  /** An LZCOMP block failed to decompress (MTX spec section 3.3). */
  | 'lzcomp-corrupt'
  /** The Compact Table Format payload is malformed (MTX spec sections 4-7). */
  | 'ctf-corrupt'
  /** The resulting / embedded sfnt is structurally invalid. */
  | 'invalid-sfnt'
  /** A resource limit (output size, glyph count, ...) would be exceeded. */
  | 'limit-exceeded';

export interface FontRejection {
  ok: false;
  reason: FontRejectReason;
  detail: string;
}

export function reject(reason: FontRejectReason, detail: string): FontRejection {
  return { ok: false, reason, detail };
}

/** Internal exception used inside the decoders; converted to FontRejection at the edge. */
export class FontFormatError extends Error {
  readonly reason: FontRejectReason;

  constructor(reason: FontRejectReason, detail: string) {
    super(detail);
    this.name = 'FontFormatError';
    this.reason = reason;
  }
}

export interface DecodeLimits {
  /** Maximum size of any decompressed buffer and of the final sfnt. Default 32 MiB. */
  maxOutputBytes: number;
  /** Maximum glyph count accepted from maxp. Default (and hard ceiling) 65535. */
  maxGlyphs: number;
  /** Maximum size of the EOT input itself. Default 64 MiB. */
  maxInputBytes: number;
}

export const DEFAULT_LIMITS: DecodeLimits = {
  maxOutputBytes: 32 * 1024 * 1024,
  maxGlyphs: 65535,
  maxInputBytes: 64 * 1024 * 1024,
};

export function resolveLimits(limits?: Partial<DecodeLimits>): DecodeLimits {
  const pick = (value: number | undefined, fallback: number, ceiling: number): number => {
    if (value === undefined || !Number.isFinite(value) || value <= 0) return fallback;
    return Math.min(Math.floor(value), ceiling);
  };
  return {
    maxOutputBytes: pick(limits?.maxOutputBytes, DEFAULT_LIMITS.maxOutputBytes, 256 * 1024 * 1024),
    maxGlyphs: pick(limits?.maxGlyphs, DEFAULT_LIMITS.maxGlyphs, 65535),
    maxInputBytes: pick(limits?.maxInputBytes, DEFAULT_LIMITS.maxInputBytes, 1024 * 1024 * 1024),
  };
}
