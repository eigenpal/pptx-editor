import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { copyFileSync, mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { recordReferences } from './record-reference.ts';
import { readManifest, sha256File } from './reference-identity.ts';

const FIXTURES = join(import.meta.dir, '../../fixtures');

let dir: string;
beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'pptx-record-'));
});
afterEach(() => rmSync(dir, { recursive: true, force: true }));

const place = (deck: string, pdf: string): void => {
  copyFileSync(join(FIXTURES, `${deck}.pptx`), join(dir, 'f.pptx'));
  copyFileSync(join(FIXTURES, `${pdf}.pdf`), join(dir, 'f.pdf'));
};

describe('record-reference', () => {
  test('records the hashes, page count, and page size of a reference', async () => {
    place('bullets-gallery', 'bullets-gallery');
    expect(await recordReferences(dir, ['f'])).toEqual(['f: 1 pages of 960 × 540 pt']);
    expect(readManifest(dir).references.f).toEqual({
      deck: sha256File(join(dir, 'f.pptx')),
      pdf: sha256File(join(dir, 'f.pdf')),
      pages: 1,
      pageSizePt: [960, 540],
    });
  });

  test('refuses a deck that changed while its PDF did not', async () => {
    place('bullets-gallery', 'bullets-gallery');
    await recordReferences(dir, ['f']);
    copyFileSync(join(FIXTURES, 'colors.pptx'), join(dir, 'f.pptx'));
    await expect(recordReferences(dir, ['f'])).rejects.toThrow(
      'the deck changed, and the PDF did not: export the PDF again'
    );
  });

  test('refuses a PDF with another page count, and keeps the manifest', async () => {
    place('bullets-gallery', 'slide-structure');
    await expect(recordReferences(dir, ['f'])).rejects.toThrow(
      'the PDF has 4 pages, more than allowed'
    );
    expect(readManifest(dir).references).toEqual({});
  });

  test('refuses a PDF that embeds fonts that the deck does not name as open fonts', async () => {
    place('colors', 'bullets-gallery');
    await expect(recordReferences(dir, ['f'])).rejects.toThrow(
      'the PDF embeds fonts that are not open fonts of the deck'
    );
  });
});
