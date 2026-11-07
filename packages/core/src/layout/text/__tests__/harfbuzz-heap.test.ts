// Font tables are read in JavaScript, not through `Face.referenceTable`: each table that
// HarfBuzz returned held a blob with a copy of the whole font that was never freed, so every
// font change grew the WebAssembly heap until shaping failed with "out of bounds".

import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { createHarfBuzzMeasurer, initializeHarfBuzz } from '../../index.ts';
import { sfntTable } from '../harfbuzz-measurer.ts';

const repo = join(import.meta.dir, '../../../../../..');
const substitutes = join(
  dirname(Bun.resolveSync('@docx-editor.dev/fonts/package.json', import.meta.dir)),
  'assets'
);
const fonts = [
  join(substitutes, 'LiberationSans-Regular.ttf'),
  join(substitutes, 'LiberationSerif-Bold.ttf'),
  join(repo, 'fixtures/fonts/NotoSansArabic-Regular.ttf'),
  join(repo, 'fixtures/fonts/NotoSansSymbols2-Regular.ttf'),
].map((path) => new Uint8Array(readFileSync(path)));

let hb: typeof import('harfbuzzjs');
beforeAll(async () => {
  await initializeHarfBuzz();
  hb = await import('harfbuzzjs');
});

test('the JavaScript table reader returns the bytes HarfBuzz reads', () => {
  for (const bytes of fonts) {
    const face = new hb.Face(new hb.Blob(bytes), 0);
    for (const tag of ['OS/2', 'post', 'kern', 'GPOS', 'zzzz']) {
      const ours = sfntTable(bytes, tag);
      const theirs = face.referenceTable(tag);
      expect(ours ? [...ours] : undefined).toEqual(theirs?.length ? [...theirs] : undefined);
    }
  }
});

test('a truncated or empty font has no tables, and reading them does not throw', () => {
  expect(sfntTable(new Uint8Array(0), 'OS/2')).toBeUndefined();
  expect(sfntTable(fonts[0]!.subarray(0, 40), 'OS/2')).toBeUndefined();
  const collection = new Uint8Array(16);
  collection.set([0x74, 0x74, 0x63, 0x66]);
  new DataView(collection.buffer).setUint32(12, 1_000_000);
  expect(sfntTable(collection, 'OS/2')).toBeUndefined();
});

test('new measurers over the same fonts do not grow the memory', async () => {
  const sources = fonts.map((bytes, i) => ({
    family: `F${i}`,
    bold: false,
    italic: false,
    bytes,
    origin: 'system',
  }));
  // HarfBuzz frees a measurer's objects from a finalizer, after a collection, as between edits.
  const churn = async (times: number): Promise<number> => {
    for (let i = 0; i < times; i += 1) {
      const m = createHarfBuzzMeasurer(sources);
      for (let f = 0; f < sources.length; f += 1)
        m.advances(m.resolve({ family: `F${f}`, bold: false, italic: false }), 'Hello');
      if (i % 10 === 9) {
        Bun.gc(true);
        await new Promise((resolve) => setTimeout(resolve, 0));
      }
    }
    Bun.gc(true);
    await new Promise((resolve) => setTimeout(resolve, 0));
    return process.memoryUsage().rss;
  };
  const start = await churn(50);
  // About 2.6 MB of font bytes for each measurer leaked before: 300 of them made 780 MB.
  const grown = (await churn(300)) - start;
  expect(grown).toBeLessThan(200 * 1024 * 1024);
}, 60_000);
