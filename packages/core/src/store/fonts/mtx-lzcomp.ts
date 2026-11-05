// LZCOMP block decompression for MicroType Express embedded fonts.
// Implements the LZCOMP compression of the W3C Member Submission "MicroType Express (MTX)
// Font Format" (2008), section 3: an adaptive Huffman coded LZ77 variant with a pre-loaded
// history window and an optional run-length stage.
//
// This is an independent implementation of section 3. It contains no code from the sample
// code in Appendix C of the submission or from any other LZCOMP or MTX decoder.
//
// Every input is untrusted: each bit read is bounds-checked, the declared length and the
// run-length expansion are held to `maxOutputBytes`, every copy is checked against the history
// before a byte moves, and each loop step consumes input, so a hostile stream cannot hang.
// All failures are `LzcompError`: kind 'limit' for the output cap, 'corrupt' for the rest.

/** Failure while decoding an LZCOMP block. `kind` is 'limit' when the output cap is hit. */
export class LzcompError extends Error {
  readonly kind: 'corrupt' | 'limit';

  constructor(message: string, kind: 'corrupt' | 'limit' = 'corrupt') {
    super(message);
    this.name = 'LzcompError';
    this.kind = kind;
  }
}

// --- bit input and output ------------------------------------------------------------------

/** Reads bits most significant first from each byte, in byte order. */
export class BitReader {
  private readonly data: Uint8Array;
  private byteIndex = 0;
  private mask = 0x80;

  constructor(data: Uint8Array) {
    this.data = data;
  }

  readBit(): number {
    if (this.byteIndex >= this.data.length) {
      throw new LzcompError('LZCOMP stream ends before the data it declares');
    }
    const bit = (this.data[this.byteIndex] & this.mask) === 0 ? 0 : 1;
    if (this.mask === 1) {
      this.mask = 0x80;
      this.byteIndex++;
    } else {
      this.mask >>= 1;
    }
    return bit;
  }

  /** Reads an unsigned field, first bit most significant. Exact up to 53 bits. */
  readValue(numberOfBits: number): number {
    let value = 0;
    for (let i = 0; i < numberOfBits; i++) value = value * 2 + this.readBit();
    return value;
  }
}

/** Writes bits in the order that `BitReader` reads them. Used by the test-side encoder. */
export class BitWriter {
  private readonly bytes: number[] = [];
  private pending = 0;
  private pendingCount = 0;

  writeBit(bit: number): void {
    this.pending = (this.pending << 1) | (bit ? 1 : 0);
    this.pendingCount++;
    if (this.pendingCount === 8) {
      this.bytes.push(this.pending);
      this.pending = 0;
      this.pendingCount = 0;
    }
  }

  /** Writes the `numberOfBits` low-order bits of `value`, most significant first. */
  writeValue(value: number, numberOfBits: number): void {
    for (let k = numberOfBits - 1; k >= 0; k--) {
      const shifted = Math.floor(value / 2 ** k);
      this.writeBit(((shifted % 2) + 2) % 2);
    }
  }

  /** Pads a partial last byte with zero bits and returns every byte written so far. */
  finish(): Uint8Array {
    if (this.pendingCount > 0) {
      this.bytes.push((this.pending << (8 - this.pendingCount)) & 0xff);
      this.pending = 0;
      this.pendingCount = 0;
    }
    return Uint8Array.from(this.bytes);
  }
}

// --- adaptive Huffman model ----------------------------------------------------------------

const MAX_MODEL_SYMBOLS = 4096;

/**
 * Adaptive Huffman code over the symbols 0 … range − 1.
 *
 * The tree lives in numbered slots 1 … 2·range − 1; slot 1 is the root. Weights never
 * increase with the slot number, so the slots of one weight form one contiguous run. An
 * increment moves the node to the first slot of its run, then adds 1, then goes to the parent.
 * Contents move between slots; the parent link stays with the slot.
 */
export class AdaptiveHuffman {
  private readonly size: number;
  /** Per slot: the left child slot (> 0) of an internal node, or −(symbol + 1) for a leaf.
   *  Child pairs never split, so the right child is always the left child + 1. */
  private readonly content: Int32Array;
  private readonly parent: Int32Array;
  /** Float64 keeps weights exact to 2^53; a hostile stream can grow them without bound. */
  private readonly weight: Float64Array;
  private readonly slotOfSymbol: Int32Array;

  constructor(range: number) {
    if (
      !Number.isInteger(range) ||
      range < 2 ||
      range > MAX_MODEL_SYMBOLS ||
      // The warm-up for 256 < range < 512 names symbol 257, which a 257-symbol model lacks.
      range === 257
    ) {
      throw new LzcompError(`unsupported adaptive Huffman alphabet size ${range}`);
    }
    this.size = range;
    const slots = 2 * range;
    this.content = new Int32Array(slots);
    this.parent = new Int32Array(slots);
    this.weight = new Float64Array(slots);
    this.slotOfSymbol = new Int32Array(range);

    for (let slot = range; slot < slots; slot++) {
      const symbol = slot - range;
      this.content[slot] = -(symbol + 1);
      this.slotOfSymbol[symbol] = slot;
      this.weight[slot] = 1;
    }
    for (let slot = range - 1; slot >= 1; slot--) {
      this.content[slot] = 2 * slot;
      this.weight[slot] = this.weight[2 * slot] + this.weight[2 * slot + 1];
    }
    for (let slot = 2; slot < slots; slot++) this.parent[slot] = slot >> 1;

    this.warmUp();
  }

  private warmUp(): void {
    const s = this.size;
    if (s > 256 && s < 512) {
      this.increment(256);
      this.increment(257);
      for (let i = 0; i < 12; i++) this.increment(s - 3);
      for (let i = 0; i < 6; i++) this.increment(s - 2);
      return;
    }
    for (let pass = 0; pass < 2; pass++) {
      for (let symbol = 0; symbol < s; symbol++) this.increment(symbol);
    }
  }

  readSymbol(bits: BitReader): number {
    let slot = 1;
    let steps = 0;
    while (this.content[slot] > 0) {
      if (++steps > 2 * this.size) throw new LzcompError('adaptive Huffman walk too deep');
      slot = this.content[slot] + bits.readBit();
    }
    const symbol = -this.content[slot] - 1;
    this.increment(symbol);
    return symbol;
  }

  writeSymbol(bits: BitWriter, symbol: number): void {
    if (!Number.isInteger(symbol) || symbol < 0 || symbol >= this.size) {
      throw new LzcompError(`symbol ${symbol} outside the alphabet of ${this.size}`);
    }
    const path: number[] = [];
    for (let slot = this.slotOfSymbol[symbol]; slot !== 1; slot = this.parent[slot]) {
      path.push(slot - this.content[this.parent[slot]]);
    }
    for (let i = path.length - 1; i >= 0; i--) bits.writeBit(path[i]);
    this.increment(symbol);
  }

  /** First slot of the run of slots whose weight equals the weight of `slot`. */
  private runStart(slot: number): number {
    // Weights do not increase with the slot number, so a binary search over 2 … slot finds the
    // lowest slot whose weight is not larger. The root (slot 1) is always strictly heavier.
    const w = this.weight[slot];
    let lo = 2;
    let hi = slot;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.weight[mid] > w) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  }

  private increment(symbol: number): void {
    let slot = this.slotOfSymbol[symbol];
    while (slot !== 1) {
      const leader = this.runStart(slot);
      if (leader !== slot) {
        this.exchange(slot, leader);
        slot = leader;
      }
      this.weight[slot]++;
      slot = this.parent[slot];
    }
    this.weight[1]++;
  }

  /** Swaps the subtrees in two slots of equal weight. Each slot keeps its own parent link. */
  private exchange(a: number, b: number): void {
    const held = this.content[a];
    this.content[a] = this.content[b];
    this.content[b] = held;
    this.relink(a);
    this.relink(b);
  }

  private relink(slot: number): void {
    const c = this.content[slot];
    if (c > 0) {
      this.parent[c] = slot;
      this.parent[c + 1] = slot;
    } else {
      this.slotOfSymbol[-c - 1] = slot;
    }
  }
}

// --- block layout --------------------------------------------------------------------------

/** Bytes of fixed history before the first decoded byte of every block. */
export const LZ_PRELOAD_SIZE = 7168;
/** From this distance on, a copy item is one byte longer than its length field says. */
export const LZ_MAX_2BYTE_DIST = 512;

const LENGTH_FIELD_BITS = 24;
const LITERAL_SYMBOLS = 256;
const SMALL_MODEL_SYMBOLS = 8;

/** Main-model symbol numbers for a block that declares `length` bytes. */
export interface LzSymbolLayout {
  /** Most distance groups a copy item can have: the smallest n ≥ 1 with 8^n ≥ length. */
  numDistRanges: number;
  /** Single-byte copy from 2 bytes back; DUP4 and DUP6 follow it. */
  dup2: number;
  dup4: number;
  dup6: number;
  /** Alphabet size of the main model. */
  numSyms: number;
}

export function lzSymbolLayout(length: number): LzSymbolLayout {
  let groups = 1;
  let reach = SMALL_MODEL_SYMBOLS;
  while (reach < length) {
    groups++;
    reach *= SMALL_MODEL_SYMBOLS;
  }
  const dup2 = LITERAL_SYMBOLS + SMALL_MODEL_SYMBOLS * groups;
  return { numDistRanges: groups, dup2, dup4: dup2 + 1, dup6: dup2 + 2, numSyms: dup2 + 3 };
}

const PAIR_REGION_END = 6144;
const PAIR_SECOND_VALUES = 96;

/** Writes the fixed history into `target[0 … 7167]`, or as much of it as `target` holds. */
export function writePreload(target: Uint8Array): void {
  const end = Math.min(target.length, LZ_PRELOAD_SIZE);
  for (let i = 0; i < end; i++) {
    if (i < PAIR_REGION_END) {
      // Byte pairs (k, j) for k = 0 … 31 and j = 0 … 95.
      const pair = i >> 1;
      target[i] = (i & 1) === 0 ? Math.floor(pair / PAIR_SECOND_VALUES) : pair % PAIR_SECOND_VALUES;
    } else {
      // Each byte value four times, in increasing order.
      target[i] = (i - PAIR_REGION_END) >> 2;
    }
  }
}

// --- run-length stage ----------------------------------------------------------------------

const RL_START = 0;
const RL_NORMAL = 1;
const RL_ESCAPE = 2;
const RL_COUNT = 3;

/** Expands the run-length escapes in the LZ bytes, one byte at a time, under the output cap. */
class RunLengthExpander {
  private readonly cap: number;
  private out: Uint8Array;
  private length = 0;
  private state = RL_START;
  private escape = 0;
  private count = 0;

  constructor(cap: number, sizeHint: number) {
    this.cap = cap;
    this.out = new Uint8Array(Math.max(16, Math.min(cap, sizeHint)));
  }

  push(value: number): void {
    switch (this.state) {
      case RL_START:
        this.escape = value;
        this.state = RL_NORMAL;
        return;
      case RL_NORMAL:
        if (value === this.escape) this.state = RL_ESCAPE;
        else this.emit(value, 1);
        return;
      case RL_ESCAPE:
        if (value === 0) {
          this.emit(this.escape, 1);
          this.state = RL_NORMAL;
        } else {
          this.count = value;
          this.state = RL_COUNT;
        }
        return;
      default:
        this.emit(value, this.count);
        this.state = RL_NORMAL;
    }
  }

  /** The expanded bytes. A pending escape or count at the end is dropped. */
  result(): Uint8Array {
    return this.out.slice(0, this.length);
  }

  private emit(value: number, times: number): void {
    const needed = this.length + times;
    if (needed > this.cap) {
      throw new LzcompError('run-length expansion exceeds the output limit', 'limit');
    }
    if (needed > this.out.length) {
      let capacity = this.out.length;
      while (capacity < needed) capacity *= 2;
      const grown = new Uint8Array(Math.min(capacity, this.cap));
      grown.set(this.out.subarray(0, this.length));
      this.out = grown;
    }
    this.out.fill(value, this.length, needed);
    this.length = needed;
  }
}

// --- block decoder -------------------------------------------------------------------------

export interface LzcompOptions {
  /** Cap on the declared LZ length and on the final output length. */
  maxOutputBytes: number;
  /** MTX header version. Version 1 blocks have no run-length flag bit. */
  version?: number;
}

/** Decodes one LZCOMP block. Throws `LzcompError` on corrupt input or when a cap is hit. */
export function lzcompDecompress(data: Uint8Array, options: LzcompOptions): Uint8Array {
  const bits = new BitReader(data);
  const runLength = options.version === 1 ? false : bits.readBit() === 1;
  const declared = bits.readValue(LENGTH_FIELD_BITS);
  if (declared > options.maxOutputBytes) {
    throw new LzcompError(`declared length ${declared} exceeds the output limit`, 'limit');
  }

  const layout = lzSymbolLayout(declared);
  const main = new AdaptiveHuffman(layout.numSyms);
  const lengths = new AdaptiveHuffman(SMALL_MODEL_SYMBOLS);
  const distances = new AdaptiveHuffman(SMALL_MODEL_SYMBOLS);
  const firstCopySymbol = LITERAL_SYMBOLS;
  const end = LZ_PRELOAD_SIZE + declared;
  const history = new Uint8Array(end);
  writePreload(history);
  const expander = runLength ? new RunLengthExpander(options.maxOutputBytes, declared) : null;

  let pos = LZ_PRELOAD_SIZE;
  while (pos < end) {
    const symbol = main.readSymbol(bits);
    let count = 1;
    if (symbol < firstCopySymbol) {
      history[pos] = symbol;
    } else if (symbol >= layout.dup2) {
      const back = 2 * (symbol - layout.dup2 + 1);
      history[pos] = history[pos - back];
    } else {
      // Copy item: length chunks (first one in the main symbol), then distance groups.
      const item = symbol - firstCopySymbol;
      const groups = Math.floor(item / SMALL_MODEL_SYMBOLS) + 1;
      let chunk = item % SMALL_MODEL_SYMBOLS;
      let extra = 0;
      for (;;) {
        extra = extra * 4 + (chunk & 3);
        if (extra > end) throw new LzcompError('copy length out of range');
        if ((chunk & 4) === 0) break;
        chunk = lengths.readSymbol(bits);
      }
      let distance = 0;
      for (let g = 0; g < groups; g++) {
        distance = distance * SMALL_MODEL_SYMBOLS + distances.readSymbol(bits);
      }
      distance += 1;
      count = extra + (distance < LZ_MAX_2BYTE_DIST ? 2 : 3);
      const from = pos - distance - count + 1;
      if (from < 0) throw new LzcompError('copy source starts before the history window');
      if (pos + count > end) throw new LzcompError('copy runs past the declared length');
      // The source ends at least one byte before `pos`, so the ranges never overlap.
      history.copyWithin(pos, from, from + count);
    }
    if (expander) {
      for (let i = 0; i < count; i++) expander.push(history[pos + i]);
    }
    pos += count;
  }

  return expander ? expander.result() : history.slice(LZ_PRELOAD_SIZE);
}
