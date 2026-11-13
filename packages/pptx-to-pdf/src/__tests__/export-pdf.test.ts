/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import {
  presentationInfo,
  readPptx,
  textOf,
  xmlPartOf,
  type Element,
} from '@pptx-editor.dev/core/store';
import { fontEmbeddingDecision } from '../pdf-font-embedding.ts';
import {
  exportPdf,
  PdfDocumentOpenError,
  PdfFidelityError,
  PdfPageLimitError,
  type PdfExportResult,
} from '../index.ts';

const FIXTURES = join(import.meta.dir, '../../../../fixtures');
const deck = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(join(FIXTURES, `${name}.pptx`)));

/** Slide count, hidden flags, slide size in points, and the `a:t` text of each slide. */
function expected(bytes: Uint8Array) {
  const read = readPptx(bytes);
  if (!read.ok) throw new Error('fixture');
  const info = presentationInfo(read.pkg);
  const texts = info.slides.map((s) => {
    const out: string[] = [];
    const walk = (el: Element): void => {
      if (el.local === 't' && el.kind === 'text') out.push(textOf(el));
      for (const c of el.children) if (c.type === 'element') walk(c);
    };
    walk(xmlPartOf(read.pkg, s.partName)!.root);
    return out.join(' ');
  });
  return { info, texts, widthPt: info.width / 12_700, heightPt: info.height / 12_700 };
}

/** Page sizes and text of each PDF page, through pdf.js (as DOCX Editor's tests read PDFs). */
async function readPdf(result: PdfExportResult) {
  const pdf = await getDocument({ data: result.bytes.slice(), useSystemFonts: false }).promise;
  const pages = [];
  for (let i = 1; i <= pdf.numPages; i += 1) {
    const page = await pdf.getPage(i);
    const [, , w, h] = page.view;
    const content = await page.getTextContent();
    const text = content.items.map((item) => ('str' in item ? item.str : '')).join('');
    pages.push({ width: w, height: h, text });
  }
  return pages;
}

// Case-insensitive: `cap="all"` runs show as capitals (ST_TextCapsType).
const squash = (s: string): string => s.replace(/\s+/g, '').toLowerCase();

describe('exportPdf', () => {
  test('one page per slide, at the slide size, with searchable text', async () => {
    const bytes = deck('text-runs');
    const want = expected(bytes);
    const result = await exportPdf(bytes, { fidelityPolicy: 'best-effort' });
    const pages = await readPdf(result);
    expect(result.pageCount).toBe(want.info.slides.length);
    expect(pages.length).toBe(want.info.slides.length);
    for (const page of pages) {
      expect(page.width).toBeCloseTo(want.widthPt, 3);
      expect(page.height).toBeCloseTo(want.heightPt, 3);
    }
    // Every character of the slide's runs is in the PDF text layer.
    const firstWords = want.texts[0].split(/\s+/).filter((w) => w.length > 2);
    expect(firstWords.length).toBeGreaterThan(0);
    for (const word of firstWords) expect(squash(pages[0].text)).toContain(squash(word));
  });

  test('the same deck gives the same bytes', async () => {
    const bytes = deck('preset-shapes');
    const a = await exportPdf(bytes, { fidelityPolicy: 'best-effort' });
    const b = await exportPdf(bytes, { fidelityPolicy: 'best-effort' });
    expect(Buffer.from(a.bytes).equals(Buffer.from(b.bytes))).toBe(true);
  });

  test('hidden slides are left out unless asked for', async () => {
    const bytes = deck('slide-structure');
    const { info } = expected(bytes);
    const hidden = info.slides.filter((s) => s.hidden).length;
    const shown = await exportPdf(bytes, { fidelityPolicy: 'best-effort' });
    expect(shown.pageCount).toBe(info.slides.length - hidden);
    const all = await exportPdf(bytes, {
      fidelityPolicy: 'best-effort',
      includeHiddenSlides: true,
    });
    expect(all.pageCount).toBe(info.slides.length);
  });

  test('pictures, pattern fills, and compound lines draw; diagnostics are per slide', async () => {
    const result = await exportPdf(deck('fills-lines'), { fidelityPolicy: 'best-effort' });
    const codes = result.diagnostics.map((d) => d.code);
    expect(codes).not.toContain('fill.pattern');
    // Every compound line draws its bands, the dashed one too.
    expect(codes).not.toContain('line.compound');
    for (const d of result.diagnostics) {
      expect(d.pageNumber).toBe((d.pageIndex ?? 0) + 1);
      expect(['unsupported', 'approximation', 'information']).toContain(d.severity);
    }
  });

  test('strict policy refuses a deck with content it cannot draw', async () => {
    const run = exportPdf(deck('fills-lines'));
    await expect(run).rejects.toBeInstanceOf(PdfFidelityError);
  });

  test('an invalid deck is refused with a typed error', async () => {
    const empty = exportPdf(new Uint8Array());
    await expect(empty).rejects.toBeInstanceOf(PdfDocumentOpenError);
    const garbage = exportPdf(new Uint8Array([1, 2, 3, 4]));
    await expect(garbage).rejects.toMatchObject({ code: 'documentOpenFailed', reason: 'zip' });
  });

  test('limits and options are checked', async () => {
    const bytes = deck('slide-structure');
    await expect(
      exportPdf(bytes, { fidelityPolicy: 'best-effort', maxPages: 1 })
    ).rejects.toBeInstanceOf(PdfPageLimitError);
    await expect(exportPdf(bytes, { maxPages: 0 })).rejects.toBeInstanceOf(RangeError);
    await expect(
      exportPdf(bytes, { fidelityPolicy: 'lenient' as 'strict' })
    ).rejects.toBeInstanceOf(RangeError);
  });
});

describe('font embedding (DOCX Editor embedder)', () => {
  /** A minimal sfnt with one OS/2 table holding `fsType`. */
  const sfnt = (fsType: number): Uint8Array => {
    const b = new Uint8Array(12 + 16 + 10);
    const v = new DataView(b.buffer);
    v.setUint16(4, 1);
    b.set([0x4f, 0x53, 0x2f, 0x32], 12); // 'OS/2'
    v.setUint32(12 + 8, 28);
    v.setUint32(12 + 12, 10);
    v.setUint16(28 + 8, fsType);
    return b;
  };
  const decide = (fsType: number) =>
    fontEmbeddingDecision({
      identity: 'f',
      bytes: sfnt(fsType),
      faceIndex: 0,
      request: { family: 'F' },
    }).kind;

  test('honors the OS/2 embedding permissions', () => {
    expect(decide(0)).toBe('embed'); // installable
    expect(decide(0x0004)).toBe('embed'); // preview and print
    expect(decide(0x0008)).toBe('embed'); // editable
    expect(decide(0x0002)).toBe('refuse'); // restricted
    expect(decide(0x0100)).toBe('refuse'); // no subsetting
    expect(decide(0x0200)).toBe('refuse'); // bitmap only
  });

  test('fonts are written as subsets, and the text stays searchable', async () => {
    const bytes = deck('text-runs');
    const result = await exportPdf(bytes, { fidelityPolicy: 'best-effort' });
    // Whole faces made this one-slide PDF 944 KB; with subsets it is about 25 KB.
    expect(result.bytes.byteLength).toBeLessThan(60_000);
    const pdf = new TextDecoder('latin1').decode(result.bytes);
    expect(pdf).toMatch(/\/BaseFont \/[A-Z]{6}\+PptxFont/);
    const [page] = await readPdf(result);
    expect(squash(page.text)).toContain(squash('The quick brown fox jumps over the lazy dog'));
  });
});
