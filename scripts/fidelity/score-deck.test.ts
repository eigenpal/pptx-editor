import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { Page } from '@playwright/test';
import { buildPptx } from '../fixtures/builder.ts';
import { deliverUpload, ResetFailed } from './harness-page.ts';
import { scoreDeck, scoreSafely } from './score-deck.ts';

/**
 * A PDF of two 960 × 540 pt pages whose text uses Helvetica without embedding it, built by hand
 * with a correct cross-reference table.
 */
function pdfWithoutEmbeddedFont(): Uint8Array {
  const page = (content: number) =>
    `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 960 540] /Resources << /Font << /F1 5 0 R >> >> /Contents ${content} 0 R >>`;
  const text = 'BT /F1 24 Tf 72 400 Td (Not embedded) Tj ET';
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >>',
    page(6),
    page(6),
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
    `<< /Length ${text.length} >>\nstream\n${text}\nendstream`,
  ];
  let out = '%PDF-1.4\n';
  const offsets: number[] = [];
  objects.forEach((body, i) => {
    offsets.push(out.length);
    out += `${i + 1} 0 obj\n${body}\nendobj\n`;
  });
  const xref = out.length;
  out += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const o of offsets) out += `${String(o).padStart(10, '0')} 00000 n \n`;
  out += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return new TextEncoder().encode(out);
}

/** A real PDF of two 960 × 540 pt pages in the open fonts that ship with the engine. */
const TWO_PAGES = join(import.meta.dir, '../../fixtures/groups-transforms.pdf');

/**
 * A harness page without Chromium. It answers the calls of `harness-page.ts` to the page's
 * `__pptxHarness` API (open, render, print, endPrint, reopenSaved, rasterPdf) and `page.pdf`.
 * Each slide paints `markup(reopened)`; `page.pdf` gives `printed`; every PDF page draws white.
 */
function fakePage(
  slides: number,
  evaluate?: (code: string) => Promise<unknown>,
  markup: (reopened: boolean) => string = () => 'same',
  printed: Uint8Array = new Uint8Array(readFileSync(TWO_PAGES))
): Page {
  let reopened = false;
  return {
    async evaluate(fn: unknown, arg?: unknown) {
      const code = String(fn);
      if (evaluate) return evaluate(code);
      if (code.includes('.rasterPdf(')) {
        const a = arg as { width: number; height: number; pages: { rgba: string; png: string }[] };
        const white = Bun.gzipSync(new Uint8Array(a.width * a.height * 4).fill(255));
        for (const p of a.pages) {
          deliverUpload(p.rgba, new Uint8Array(white));
          deliverUpload(p.png, new Uint8Array([0x89]));
        }
        return undefined;
      }
      if (code.includes('.open('))
        return { slides, received: { deck: { bytes: 0, sha256: '' }, fonts: [] } };
      if (code.includes('.reopenSaved(')) {
        reopened = true;
        return { slides };
      }
      if (code.includes('.render('))
        return { diagnostics: [], fonts: [], markup: markup(reopened) };
      if (code.includes('.print(') || code.includes('.endPrint(')) return undefined;
      throw new Error(`unexpected harness call: ${code}`);
    },
    pdf: async () => Buffer.from(printed),
  } as unknown as Page;
}

let dir: string;
let deck: string;
beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'pptx-score-'));
  deck = join(dir, 'deck.pptx');
  // Three slides of 960 × 540 pt; slide 2 is hidden.
  writeFileSync(
    deck,
    buildPptx({
      width: 12_192_000,
      height: 6_858_000,
      slides: [{ shapes: '' }, { shapes: '', hidden: true }, { shapes: '' }],
    })
  );
});
afterEach(() => rmSync(dir, { recursive: true, force: true }));

describe('scoreDeck', () => {
  test('without a PDF, every slide is a gap; a hidden slide has no page and says why', async () => {
    const r = await scoreDeck(fakePage(3), { file: deck, name: 'd' });
    expect(r).toMatchObject({
      name: 'd',
      reference: 'none',
      referenceIssues: ['no reference PDF'],
    });
    expect(r.slides.map((s) => [s.index, s.page, s.hidden, s.status, s.cause])).toEqual([
      [1, 1, false, 'no-reference', 'no reference PDF'],
      [2, undefined, true, 'no-reference', 'hidden slide: the reference PDF leaves it out'],
      [3, 2, false, 'no-reference', 'no reference PDF'],
    ]);
    expect(r.slides.every((s) => s.selfConsistent)).toBe(true);
  });

  test('a slide is save-consistent only when it paints the same markup after reopen', async () => {
    const r = await scoreDeck(
      fakePage(3, undefined, (again) => (again ? 'after' : 'before')),
      {
        file: deck,
        name: 'd',
      }
    );
    expect(r.slides.map((s) => s.selfConsistent)).toEqual([false, false, false]);
  });

  test('a PDF that does not open is stale, with a fixed reason and never its path', async () => {
    const pdf = join(dir, 'secret client name.pdf');
    writeFileSync(pdf, 'not a PDF');
    const r = await scoreDeck(fakePage(3), { file: deck, name: 'd', pdf });
    expect(r.reference).toBe('stale');
    expect(r.referenceIssues).toEqual(['pdf.js cannot read the PDF']);
    expect(JSON.stringify(r)).not.toContain('secret');
  });

  test('a PDF with another page count is stale', async () => {
    const pdf = join(import.meta.dir, '../../fixtures/slide-structure.pdf'); // 4 pages
    const r = await scoreDeck(fakePage(3), { file: deck, name: 'd', pdf });
    expect(r.referenceIssues).toEqual(['the PDF has 4 pages, the deck has 2 visible slides']);
    expect(r.slides.map((s) => s.status)).toEqual(['stale', 'no-reference', 'stale']);
  });

  test('with real PDFs, the visible slides compare page by page and the hidden one has no page', async () => {
    const r = await scoreDeck(fakePage(3), { file: deck, name: 'd', pdf: TWO_PAGES });
    expect(r.reference).toBe('pdf');
    expect(r.foreignFonts).toEqual([]);
    expect(r.slides.map((s) => [s.index, s.page, s.hidden, s.status])).toEqual([
      [1, 1, false, 'pass'],
      [2, undefined, true, 'no-reference'],
      [3, 2, false, 'pass'],
    ]);
    expect(r.slides[0].linesTotal).toBeGreaterThan(0);
    expect(r.slides[0].pixelMismatch).toBe(0);
  });

  test('fonts of our PDF that the harness did not load are listed', async () => {
    // Open Arabic fixture fonts are absent from this empty deck and its loaded fonts.
    const pdf = join(import.meta.dir, '../../fixtures/rtl-text.pdf'); // four pages
    const one = join(dir, 'one.pptx');
    writeFileSync(
      one,
      buildPptx({
        width: 12_192_000,
        height: 6_858_000,
        slides: Array.from({ length: 4 }, () => ({ shapes: '' })),
      })
    );
    const page = fakePage(4, undefined, undefined, new Uint8Array(readFileSync(pdf)));
    const r = await scoreDeck(page, { file: one, name: 'd', pdf });
    expect(r.foreignFonts).toEqual(['NotoSansArabic-Bold', 'NotoSansArabic-Regular']);
  });

  test('a reference PDF with a font that it does not embed is stale', async () => {
    const pdf = join(dir, 'reference.pdf');
    writeFileSync(pdf, pdfWithoutEmbeddedFont());
    const r = await scoreDeck(fakePage(3), { file: deck, name: 'd', pdf });
    expect(r.reference).toBe('stale');
    expect(r.referenceIssues).toEqual(['the PDF does not embed the font Helvetica']);
  });

  test('a font of our PDF that is not embedded is listed', async () => {
    const page = fakePage(3, undefined, undefined, pdfWithoutEmbeddedFont());
    const r = await scoreDeck(page, { file: deck, name: 'd', pdf: TWO_PAGES });
    expect(r.foreignFonts).toEqual(['Helvetica (not embedded)']);
  });

  test('our PDF must have one page of the slide size for each visible slide', async () => {
    const ours = join(import.meta.dir, '../../fixtures/slide-size.pdf'); // A4, one page
    const page = fakePage(3, undefined, undefined, new Uint8Array(readFileSync(ours)));
    await expect(scoreDeck(page, { file: deck, name: 'd', pdf: TWO_PAGES })).rejects.toThrow(
      'our PDF has 1 pages for 2 slides'
    );
  });

  test('the editor and the package reader must count the same slides', async () => {
    const r = await scoreDeck(fakePage(5), { file: deck, name: 'd' });
    expect(r).toEqual({
      name: 'd',
      error: 'the editor opens 5 slides, the package has 3',
      slides: [],
    });
  });
});

describe('scoreSafely', () => {
  const reset = (next: Page) => {
    let calls = 0;
    return {
      control: {
        async reset() {
          calls += 1;
          return { page: next };
        },
      },
      calls: () => calls,
    };
  };

  test('an error becomes the deck error, with no path, and the page is replaced', async () => {
    const next = fakePage(3);
    const r = reset(next);
    const broken = fakePage(3, async () => {
      throw new Error('Target crashed while reading /Users/me/client deck.pptx');
    });
    const scored = await scoreSafely(
      broken,
      r.control,
      { file: deck, name: 'd' },
      { redact: true }
    );
    expect(scored.result).toEqual({
      name: 'd',
      error: 'the harness failed: Target crashed while reading <path>',
      slides: [],
    });
    expect(scored.page).toBe(next);
    expect(r.calls()).toBe(1);
  });

  test('a deck that hangs stops at the time limit', async () => {
    const r = reset(fakePage(3));
    const hung = fakePage(3, () => new Promise(() => {}));
    const scored = await scoreSafely(
      hung,
      r.control,
      { file: deck, name: 'd' },
      { redact: false, timeoutMs: 50 }
    );
    expect(scored.result.error).toBe('the harness failed: the deck did not finish within 0.05 s');
    expect(r.calls()).toBe(1);
  });

  test('when no new page opens, the error is still returned with a fatal reason', async () => {
    const broken = fakePage(3, async () => {
      throw new Error('Browser closed');
    });
    const scored = await scoreSafely(
      broken,
      {
        async reset() {
          throw new Error('the browser is gone');
        },
      },
      { file: deck, name: 'd' },
      { redact: false }
    );
    expect(scored.result.error).toBe('the harness failed: Browser closed');
    expect(scored.fatal?.message).toBe('the browser is gone');
  });

  for (const [newBrowser, cause] of [
    ['exited', 'the browser crashed while rendering the deck'],
    ['hung', 'the browser stopped answering while rendering the deck'],
  ] as const) {
    test(`a browser that ${newBrowser} is the deck's cause, and the next page is new`, async () => {
      const next = fakePage(3);
      const broken = fakePage(3, async () => {
        throw new Error('page.evaluate: Target page, context or browser has been closed');
      });
      const scored = await scoreSafely(
        broken,
        { reset: async () => ({ page: next, newBrowser }) },
        { file: deck, name: 'd' },
        { redact: true }
      );
      expect(scored.result.error).toBe(
        `${cause} (the next deck gets a new browser): page.evaluate: Target page, context or browser has been closed`
      );
      expect(scored.page).toBe(next);
      expect(scored.fatal).toBeUndefined();
    });
  }

  test('when no new browser opens, the deck error keeps the crash cause', async () => {
    const broken = fakePage(3, async () => {
      throw new Error('page.evaluate: Target page, context or browser has been closed');
    });
    const scored = await scoreSafely(
      broken,
      {
        async reset() {
          throw new ResetFailed(
            'a new browser did not open the harness page: launch failed',
            'exited'
          );
        },
      },
      { file: deck, name: 'd' },
      { redact: false }
    );
    expect(scored.result.error).toBe(
      'the browser crashed while rendering the deck (no new browser opened, so the run stops): page.evaluate: Target page, context or browser has been closed'
    );
    expect(scored.fatal?.message).toBe(
      'a new browser did not open the harness page: launch failed'
    );
  });

  test('a deck that scores is returned as it is', async () => {
    const page = fakePage(3);
    const scored = await scoreSafely(
      page,
      reset(page).control,
      { file: deck, name: 'd' },
      { redact: true }
    );
    expect(scored.page).toBe(page);
    expect(scored.result.slides).toHaveLength(3);
  });
});
