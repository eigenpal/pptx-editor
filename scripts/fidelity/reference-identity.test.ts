import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { copyFileSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { buildPptx } from '../fixtures/builder.ts';
import {
  deckFacts,
  readManifest,
  referenceProblems,
  referenceState,
  sha256File,
  slidePages,
  writeManifest,
  type DeckFacts,
  type ReferenceIdentity,
} from './reference-identity.ts';
import { openPdfFile } from './pdf.ts';

const FIXTURE = join(import.meta.dir, '../../fixtures/slide-structure.pptx');

let dir: string;
beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'pptx-refs-'));
});
afterEach(() => rmSync(dir, { recursive: true, force: true }));

function recorded(deck: string, pdf: string): ReferenceIdentity {
  return {
    deck: sha256File(deck),
    pdf: sha256File(pdf),
    pages: 4,
    pageSizePt: [960, 540],
  };
}

describe('referenceState', () => {
  test('missing without a PDF, even with an identity', () => {
    const deck = join(dir, 'a.pptx');
    copyFileSync(FIXTURE, deck);
    expect(referenceState(deck, readManifest(dir))).toEqual({ status: 'missing' });
  });

  test('stale when the PDF has no recorded identity', () => {
    const deck = join(dir, 'a.pptx');
    copyFileSync(FIXTURE, deck);
    writeFileSync(join(dir, 'a.pdf'), '%PDF-1.4 hand-made');
    const state = referenceState(deck, readManifest(dir));
    expect(state.status).toBe('stale');
    expect(state.status === 'stale' && state.reasons).toEqual(['no identity recorded for the PDF']);
  });

  test('current while the deck and the PDF are the recorded bytes', () => {
    const deck = join(dir, 'a.pptx');
    const pdf = join(dir, 'a.pdf');
    copyFileSync(FIXTURE, deck);
    writeFileSync(pdf, '%PDF-1.4 reference');
    writeManifest(dir, { version: 1, references: { a: recorded(deck, pdf) } });
    const state = referenceState(deck, readManifest(dir));
    expect(state.status).toBe('current');
  });

  test('stale when the deck or the PDF changed after they were recorded', () => {
    const deck = join(dir, 'a.pptx');
    const pdf = join(dir, 'a.pdf');
    copyFileSync(FIXTURE, deck);
    writeFileSync(pdf, '%PDF-1.4 reference');
    writeManifest(dir, { version: 1, references: { a: recorded(deck, pdf) } });
    writeFileSync(deck, Buffer.concat([readFileSync(deck), Buffer.from([0])]));
    writeFileSync(pdf, '%PDF-1.4 replaced by hand');
    const state = referenceState(deck, readManifest(dir));
    expect(state.status === 'stale' && state.reasons).toEqual([
      'the deck changed after its reference was recorded',
      'the PDF changed after it was recorded',
    ]);
  });

  test('a manifest key from the prototype never counts as an identity', () => {
    const deck = join(dir, 'constructor.pptx');
    copyFileSync(FIXTURE, deck);
    writeFileSync(join(dir, 'constructor.pdf'), '%PDF');
    expect(referenceState(deck, readManifest(dir)).status).toBe('stale');
  });
});

describe('manifest', () => {
  test('writes sorted keys and reads them back', () => {
    const deck = join(dir, 'a.pptx');
    const pdf = join(dir, 'a.pdf');
    copyFileSync(FIXTURE, deck);
    writeFileSync(pdf, '%PDF');
    writeManifest(dir, {
      version: 1,
      references: { z: recorded(deck, pdf), a: recorded(deck, pdf) },
    });
    const text = readFileSync(join(dir, 'references.json'), 'utf8');
    expect(text.indexOf('"a"')).toBeLessThan(text.indexOf('"z"'));
    expect(text.endsWith('}\n')).toBe(true);
    expect(Object.keys(readManifest(dir).references)).toEqual(['a', 'z']);
  });

  test('refuses an unknown version', () => {
    writeFileSync(join(dir, 'references.json'), '{"version":2,"references":{}}');
    expect(() => readManifest(dir)).toThrow('unknown version');
  });
});

/** Deck facts of a 960 × 540 pt deck, with `over` applied. */
const facts16x9 = (over: Partial<DeckFacts> = {}): DeckFacts => ({
  slides: 1,
  hidden: [],
  widthPt: 960,
  heightPt: 540,
  widthEmu: 12_192_000,
  heightEmu: 6_858_000,
  ...over,
});

describe('deck and PDF facts', () => {
  test('deckFacts reads slide count, hidden slides, and slide size', () => {
    const bytes = buildPptx({
      width: 9_144_000,
      height: 6_858_000,
      slides: [{ shapes: '' }, { shapes: '', hidden: true }, { shapes: '' }],
    });
    expect(deckFacts(bytes)).toEqual({
      slides: 3,
      hidden: [2],
      widthPt: 720,
      heightPt: 540,
      widthEmu: 9_144_000,
      heightEmu: 6_858_000,
    });
  });

  test('slidePages maps visible slides to pages in order and skips hidden ones', () => {
    expect(slidePages(facts16x9({ slides: 4, hidden: [1, 3] }))).toEqual([
      undefined,
      0,
      undefined,
      1,
    ]);
  });

  test('openPdf reads every page size and the producer of a reference PDF', async () => {
    const pdf = await openPdfFile(join(import.meta.dir, '../../fixtures/slide-structure.pdf'));
    try {
      expect(pdf.pages.length).toBe(4);
      expect(pdf.pages[0]).toEqual([960, 540]);
      expect(referenceProblems(deckFacts(new Uint8Array(readFileSync(FIXTURE))), pdf)).toEqual([]);
    } finally {
      await pdf.close();
    }
  });

  test('referenceProblems refuses a PDF that another presentation tool made', () => {
    const deck = facts16x9();
    const page = [[960, 540]] as const;
    expect(
      referenceProblems(deck, { pages: page, producer: 'macOS Version 27.0.1 Quartz PDFContext' })
    ).toEqual([]);
    expect(
      referenceProblems(deck, {
        pages: page,
        producer: 'Keynote; macOS Version 27.0.1 Quartz PDFContext',
      })
    ).toEqual([
      'the PDF was made by Keynote; macOS Version 27.0.1 Quartz PDFContext, which never makes a reference PDF',
    ]);
    expect(referenceProblems(deck, { pages: page, producer: 'LibreOffice 25.2' })).toHaveLength(1);
    expect(referenceProblems(deck, { pages: page, producer: '' })).toEqual([]);
  });

  test('referenceProblems checks the visible slide count and the page size within 0.625 pt', () => {
    const deck = facts16x9({ slides: 3, hidden: [2] });
    expect(
      referenceProblems(deck, {
        pages: [
          [960, 540],
          [960.6, 539.4],
        ],
      })
    ).toEqual([]);
    expect(referenceProblems(deck, { pages: [[960, 540]] })).toEqual([
      'the PDF has 1 pages, the deck has 2 visible slides',
    ]);
    expect(
      referenceProblems(deck, {
        pages: [
          [960, 540],
          [720, 540],
        ],
      })
    ).toEqual(['page 2 is 720 × 540 pt, the slide is 960 × 540 pt']);
    // A reference PDF rounds a side to a whole point, so 1 pt off is already wrong.
    expect(
      referenceProblems(deck, {
        pages: [
          [960, 540],
          [961, 540],
        ],
      })
    ).toHaveLength(1);
    // The new fixture's page: 400.25 × 300.5 pt rounds to 400 × 301 pt.
    expect(
      referenceProblems(
        facts16x9({ widthPt: 400.25, heightPt: 300.5, widthEmu: 5_083_175, heightEmu: 3_816_350 }),
        { pages: [[400, 301]] }
      )
    ).toEqual([]);
  });
});
