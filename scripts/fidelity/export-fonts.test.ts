import { describe, expect, test } from 'bun:test';
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';

// The export lane checks only that each font of the exported PDF is embedded (`score-deck.ts`):
// `exportPdf` names every face `PptxFont`, so the lane cannot match faces by name. That is sound
// only while the exporter cannot reach a system font. It draws with the faces that the caller
// gives, the packaged substitutes, and the optional CJK package (`cjk-font.ts`), and it reads no
// other file. This test keeps it that way.

const SRC = 'packages/pptx-to-pdf/src';

/** The source files of the exporter, without its tests. */
function sources(dir: string): string[] {
  return readdirSync(dir, { withFileTypes: true }).flatMap((e) => {
    const path = join(dir, e.name);
    if (e.isDirectory()) return e.name === '__tests__' ? [] : sources(path);
    return e.name.endsWith('.ts') && !e.name.endsWith('.test.ts') ? [path] : [];
  });
}

/** An import of a module that reads files from the host. */
const FILE_ACCESS = /from ['"](?:node:)?fs(?:\/promises)?['"]|require\(['"](?:node:)?fs/;

describe('the export lane font check', () => {
  test('the exporter reads files only to load the CJK package', () => {
    const readers = sources(SRC).filter((f) => FILE_ACCESS.test(readFileSync(f, 'utf8')));
    expect(readers).toEqual([join(SRC, 'cjk-font.ts')]);
  });
});
