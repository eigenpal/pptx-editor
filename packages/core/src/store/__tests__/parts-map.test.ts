// A derived part map reads exactly as a copied and edited `Map`: the same entries, size, and
// order. Each test runs the same edits on a plain `Map` model and compares.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { DerivedParts } from '../package/parts-map.ts';
import {
  applyOps,
  changedPartNames,
  readPptx,
  type PackagePart,
  type PptxPackage,
} from '../index.ts';

const part = (name: string, tag = 0): PackagePart => ({
  kind: 'binary',
  partName: name,
  bytes: new Uint8Array([tag]),
});

/** A deterministic random source (mulberry32). */
function random(seed: number): () => number {
  let a = seed;
  return () => {
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function same(a: ReadonlyMap<string, PackagePart>, b: ReadonlyMap<string, PackagePart>): void {
  expect([...a.keys()]).toEqual([...b.keys()]);
  expect([...a.values()]).toEqual([...b.values()]);
  expect([...a]).toEqual([...b.entries()]);
  expect(a.size).toBe(b.size);
  for (const name of b.keys()) expect(a.get(name)).toBe(b.get(name));
}

describe('DerivedParts', () => {
  test('random edits over many derivations read as a copied Map, in order', () => {
    const next = random(7);
    const names = Array.from({ length: 40 }, (_, i) => `/p/${i}`);
    let model = new Map(names.slice(0, 20).map((n) => [n, part(n)]));
    let derived: ReadonlyMap<string, PackagePart> = new Map(model);
    for (let round = 0; round < 300; round += 1) {
      const copy = new Map(model);
      const edit = DerivedParts.from(derived);
      for (let k = 0; k < 1 + Math.floor(next() * 4); k += 1) {
        const name = names[Math.floor(next() * names.length)];
        if (next() < 0.35) {
          expect(edit.delete(name)).toBe(copy.delete(name));
        } else {
          const p = part(name, round);
          edit.set(name, p);
          copy.set(name, p);
        }
      }
      same(edit, copy);
      for (const name of names) expect(edit.has(name)).toBe(copy.has(name));
      // Changed names: exactly the names whose entry differs.
      const expected = names.filter((n) => model.get(n) !== copy.get(n)).sort();
      const before: PptxPackage = { parts: derived };
      const after: PptxPackage = { parts: edit.seal() };
      expect([...(changedPartNames(before, after, Infinity) ?? [])].sort()).toEqual(expected);
      model = copy;
      derived = edit;
    }
  });

  test('a sealed map refuses writes', () => {
    const sealed = DerivedParts.from(new Map([['/a', part('/a')]])).seal();
    expect(() => sealed.set('/b', part('/b'))).toThrow();
    expect(() => sealed.delete('/a')).toThrow();
  });

  test('past the change bound, a derivation starts a new root and stays equal', () => {
    const base = new Map(Array.from({ length: 10 }, (_, i) => [`/r/${i}`, part(`/r/${i}`)]));
    let derived: ReadonlyMap<string, PackagePart> = base;
    const model = new Map(base);
    for (let i = 0; i < 1200; i += 1) {
      const edit = DerivedParts.from(derived);
      const p = part(`/n/${i}`);
      edit.set(`/n/${i}`, p);
      model.set(`/n/${i}`, p);
      derived = edit.seal();
      expect(edit.changeCount()).toBeLessThanOrEqual(513);
    }
    same(derived, model);
  });

  test('a commit on a long deck copies no part map, and its changed parts are its own', () => {
    const r = readPptx(
      buildPptx({
        slides: Array.from({ length: 30 }, (_, i) => ({
          shapes: shape({ id: 2, x: 0, y: 0, cx: 100, cy: 100, text: para(run(`S${i}`)) }),
        })),
      })
    );
    if (!r.ok) throw new Error('deck');
    const slide = '/ppt/slides/slide7.xml';
    const xml = r.pkg.parts.get(slide);
    if (xml?.kind !== 'xml') throw new Error('slide');
    let pkg = r.pkg;
    for (let i = 0; i < 5; i += 1) {
      const out = applyOps(pkg, [
        {
          op: 'setAttribute',
          part: slide,
          node: xml.xml.root.id,
          name: { ns: '', local: 'show' },
          value: `${i % 2}`,
        },
      ]);
      if (!out.ok) throw new Error(JSON.stringify(out.rejection));
      expect(out.pkg.parts).toBeInstanceOf(DerivedParts);
      expect(changedPartNames(pkg, out.pkg, 8)).toEqual([slide]);
      pkg = out.pkg;
    }
    // Every other part is the object the read gave.
    for (const [name, p] of r.pkg.parts) if (name !== slide) expect(pkg.parts.get(name)).toBe(p);
    expect([...pkg.parts.keys()]).toEqual([...r.pkg.parts.keys()]);
  });
});
