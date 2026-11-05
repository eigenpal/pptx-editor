// MicroType Express container decoding.
// Implements W3C Member Submission "MicroType Express (MTX) Font Format" (2008),
// section 2.1 "MicroType Express Font Header" (version, copy limit, block offsets) and
// Appendix A validity checks, then drives LZCOMP (section 3, mtx-lzcomp.ts) over the three
// blocks and CTF reconstruction (sections 4-7, mtx-ctf.ts / mtx-glyf.ts).

import { ctfToSfnt } from './mtx-ctf.ts';
import { LzcompError, lzcompDecompress } from './mtx-lzcomp.ts';
import { FontFormatError, type DecodeLimits } from './types.ts';

export interface MtxHeader {
  version: number;
  copyLimit: number;
  offsetData2: number;
  offsetData3: number;
}

const u24 = (d: Uint8Array, p: number): number => (d[p] << 16) | (d[p + 1] << 8) | d[p + 2];

/** Parses and validates the 10-byte MTX header (offsets are from the start of the MTX data). */
export function parseMtxHeader(data: Uint8Array): MtxHeader {
  if (data.length < 10) throw new FontFormatError('truncated', 'MTX data shorter than header');
  const header: MtxHeader = {
    version: data[0],
    copyLimit: u24(data, 1),
    offsetData2: u24(data, 4),
    offsetData3: u24(data, 7),
  };
  // Version 3 is "1.0"; Appendix C's decoder also handles version 1 (no run-length flag bit).
  if (header.version !== 3 && header.version !== 1) {
    throw new FontFormatError('unsupported-version', `MTX version ${header.version}`);
  }
  if (
    header.offsetData2 < 10 ||
    header.offsetData3 < header.offsetData2 ||
    header.offsetData3 > data.length
  ) {
    throw new FontFormatError('invalid-mtx-header', 'MTX block offsets out of order or bounds');
  }
  return header;
}

function inflateBlock(
  block: Uint8Array,
  name: string,
  version: number,
  limits: DecodeLimits
): Uint8Array {
  try {
    return lzcompDecompress(block, { maxOutputBytes: limits.maxOutputBytes, version });
  } catch (err) {
    if (err instanceof LzcompError) {
      const reason = err.kind === 'limit' ? 'limit-exceeded' : 'lzcomp-corrupt';
      throw new FontFormatError(reason, `${name}: ${err.message}`);
    }
    throw err;
  }
}

export interface MtxDecodeResult {
  sfnt: Uint8Array;
  numGlyphs: number;
  warnings: string[];
  header: MtxHeader;
}

/** Decompresses an MTX payload into a TrueType sfnt. Throws FontFormatError on bad input. */
export function decodeMtx(data: Uint8Array, limits: DecodeLimits): MtxDecodeResult {
  const header = parseMtxHeader(data);
  const block1 = inflateBlock(
    data.subarray(10, header.offsetData2),
    'block1',
    header.version,
    limits
  );
  const block2 = inflateBlock(
    data.subarray(header.offsetData2, header.offsetData3),
    'block2',
    header.version,
    limits
  );
  const block3 = inflateBlock(data.subarray(header.offsetData3), 'block3', header.version, limits);
  const total = block1.length + block2.length + block3.length;
  if (total > limits.maxOutputBytes) {
    throw new FontFormatError('limit-exceeded', 'decompressed MTX blocks exceed limit');
  }
  const ctf = ctfToSfnt({
    block1,
    block2,
    block3,
    maxOutputBytes: limits.maxOutputBytes,
    maxGlyphs: limits.maxGlyphs,
  });
  return { ...ctf, header };
}
