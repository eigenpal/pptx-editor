// The one-pass XML reader from DOCX Editor (`@eigenpal/ooxml-core` `xml-fast-read.ts`) must
// build exactly the tree the bounded reader builds on PresentationML parts too, or give the part
// to it. This reads every XML part of every synthetic fixture both ways.

import { expect, test } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, unzipSync } from 'fflate';
import {
  XML_DEFAULT_MAX_ELEMENTS,
  fastReadXml,
  readXml,
  readXmlBounded,
} from '@eigenpal/ooxml-core';

const FIXTURES = join(import.meta.dir, '../../../../../fixtures');

test('every XML part of every fixture reads to the same tree both ways', () => {
  let parts = 0;
  let fast = 0;
  for (const file of readdirSync(FIXTURES).filter((name) => name.endsWith('.pptx'))) {
    const entries = unzipSync(new Uint8Array(readFileSync(join(FIXTURES, file))));
    for (const [name, bytes] of Object.entries(entries)) {
      if (!/\.(xml|rels)$/.test(name)) continue;
      const xml = strFromU8(bytes);
      parts += 1;
      const bounded = readXmlBounded(xml);
      const quick = fastReadXml(xml, XML_DEFAULT_MAX_ELEMENTS);
      if (quick !== null) {
        fast += 1;
        expect(bounded.ok, `${file} ${name}`).toBe(true);
        // JSON text keeps attribute order, which `toEqual` ignores; a save writes that order.
        if (bounded.ok)
          expect(JSON.stringify(quick), `${file} ${name}`).toBe(JSON.stringify(bounded.nodes));
      }
      expect(JSON.stringify(readXml(xml)), `${file} ${name}`).toBe(JSON.stringify(bounded));
    }
  }
  expect(parts).toBeGreaterThan(100);
  // The comparison means something only when the one-pass path is taken.
  expect(fast).toBeGreaterThan(0);
});
