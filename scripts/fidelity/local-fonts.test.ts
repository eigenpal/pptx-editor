import { describe, expect, test } from 'bun:test';
import {
  MAX_COLLECTION_FACES,
  collectionFace,
  faceOffsets,
  localFontsIdentity,
  referenceFontsInstalled,
} from './local-fonts.ts';

/** Write a face of `tables` (tag, bytes, offset) with its table directory at `base`. */
function writeFace(
  out: Uint8Array,
  base: number,
  tables: readonly (readonly [string, Uint8Array, number])[]
): void {
  const v = new DataView(out.buffer);
  v.setUint32(base, 0x00010000);
  v.setUint16(base + 4, tables.length);
  tables.forEach(([tag, bytes, offset], i) => {
    const record = base + 12 + i * 16;
    for (let k = 0; k < 4; k += 1) out[record + k] = tag.charCodeAt(k);
    v.setUint32(record + 8, offset);
    v.setUint32(record + 12, bytes.length);
    out.set(bytes, offset);
  });
}

/** A collection of two faces that share the table `glyf` and differ in `name`. */
function collection(): Uint8Array {
  const out = new Uint8Array(256);
  const v = new DataView(out.buffer);
  v.setUint32(0, 0x74746366);
  v.setUint32(4, 0x00010000);
  v.setUint32(8, 2);
  v.setUint32(12, 20);
  v.setUint32(16, 80);
  const glyf = new Uint8Array([1, 2, 3, 4, 5]);
  writeFace(out, 20, [
    ['glyf', glyf, 200],
    ['name', new Uint8Array([10, 11]), 220],
  ]);
  writeFace(out, 80, [
    ['glyf', glyf, 200],
    ['name', new Uint8Array([20, 21, 22]), 230],
  ]);
  return out;
}

describe('local fonts', () => {
  test('the identity is a stable SHA-256 of the font files', () => {
    const a = localFontsIdentity();
    expect(a).toMatch(/^[0-9a-f]{64}$/);
    expect(localFontsIdentity()).toBe(a);
    expect(typeof referenceFontsInstalled()).toBe('boolean');
  });

  test('a collection gives each face as a single font with its own tables', () => {
    const ttc = collection();
    expect(faceOffsets(ttc)).toEqual([20, 80]);
    const face = collectionFace(ttc, 1);
    expect(face).toBeDefined();
    const v = new DataView(face!.buffer);
    expect(v.getUint32(0)).toBe(0x00010000);
    expect(v.getUint16(4)).toBe(2);
    const table = (i: number) => {
      const record = 12 + i * 16;
      const offset = v.getUint32(record + 8);
      return [
        String.fromCharCode(...face!.subarray(record, record + 4)),
        [...face!.subarray(offset, offset + v.getUint32(record + 12))],
        offset % 4,
      ];
    };
    expect(table(0)).toEqual(['glyf', [1, 2, 3, 4, 5], 0]);
    expect(table(1)).toEqual(['name', [20, 21, 22], 0]);
    // A single font is its own face 0.
    expect(faceOffsets(face!)).toEqual([0]);
  });

  test('a face outside the file, or a table past its end, gives no font', () => {
    const ttc = collection();
    expect(collectionFace(ttc, 2)).toBeUndefined();
    const cut = ttc.slice(0, 225);
    expect(collectionFace(cut, 0)).toBeDefined();
    expect(collectionFace(cut, 1)).toBeUndefined();
    // The face count is capped.
    const many = collection();
    new DataView(many.buffer).setUint32(8, 1_000_000);
    expect(faceOffsets(many).length).toBeLessThanOrEqual(MAX_COLLECTION_FACES);
  });
});
