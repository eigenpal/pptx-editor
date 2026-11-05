import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { Page } from '@playwright/test';
import { PNG } from 'pngjs';
import { buildPptx } from '../fixtures/builder.ts';
import { hasPoppler } from './references.ts';
import { ResetFailed } from './harness-page.ts';
import { scoreDeck, scoreSafely } from './score-deck.ts';
import { THRESHOLDS } from './thresholds.ts';

/** A white PNG of `width` × `height` px. */
function white(width: number, height: number): Buffer {
  const p = new PNG({ width, height });
  p.data.fill(255);
  return PNG.sync.write(p);
}
const WHITE = white(4, 3);

/** A PNG made `scale` times larger, each pixel repeated, as Chromium paints at a device scale. */
function upscale(png: Buffer, scale: number): Buffer {
  const src = PNG.sync.read(png);
  const out = new PNG({ width: src.width * scale, height: src.height * scale });
  for (let y = 0; y < out.height; y += 1)
    for (let x = 0; x < out.width; x += 1)
      src.data.copy(
        out.data,
        (y * out.width + x) * 4,
        (((y / scale) | 0) * src.width + ((x / scale) | 0)) * 4,
        (((y / scale) | 0) * src.width + ((x / scale) | 0)) * 4 + 4
      );
  return PNG.sync.write(out);
}

/**
 * A harness page without Chromium. It answers the calls of `harness-page.ts` to the page's
 * `__pptxHarness` API: open, render, and reopenSaved. Its screenshot of `render` is at the raster
 * scale, as the real page paints it (`THRESHOLDS.rasterScale`).
 */
function fakePage(
  slides: number,
  evaluate?: (code: string) => Promise<unknown>,
  render: Buffer = WHITE
): Page {
  const shot = upscale(render, THRESHOLDS.rasterScale);
  return {
    async evaluate(fn: unknown) {
      const code = String(fn);
      // The capture waits for fonts, as Playwright's screenshot does.
      if (code.includes('document.fonts.ready')) return undefined;
      if (evaluate) return evaluate(code);
      if (code.includes('.open('))
        return { slides, received: { deck: { bytes: 0, sha256: '' }, fonts: [] } };
      if (code.includes('.reopenSaved(')) return { slides };
      if (code.includes('.render(')) return { lines: [], diagnostics: [] };
      throw new Error(`unexpected harness call: ${code}`);
    },
    viewportSize: () => null,
    locator: () => ({ boundingBox: async () => null, screenshot: async () => shot }),
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

  test('a PDF that does not open is stale, with a fixed reason and never its path', async () => {
    const pdf = join(dir, 'secret client name.pdf');
    writeFileSync(pdf, 'not a PDF');
    const r = await scoreDeck(fakePage(3), { file: deck, name: 'd', pdf });
    expect(r.reference).toBe('stale');
    expect(r.referenceIssues).toEqual(['the PDF does not open (pdfinfo failed)']);
    expect(JSON.stringify(r)).not.toContain('secret');
  });

  test.skipIf(!hasPoppler())('a PDF with another page count is stale', async () => {
    const pdf = join(import.meta.dir, '../../fixtures/slide-structure.pdf'); // 4 pages
    const r = await scoreDeck(fakePage(3), { file: deck, name: 'd', pdf });
    expect(r.referenceIssues).toEqual(['the PDF has 4 pages, the deck has 2 visible slides']);
    expect(r.slides.map((s) => s.status)).toEqual(['stale', 'no-reference', 'stale']);
  });

  test.skipIf(!hasPoppler())(
    'an A4 slide is compared with the slide area of its 842 × 596 pt page, and its render size is checked',
    async () => {
      // The render of the 842 × 595.5 pt `slide-size` fixture is 1920 × 1358 px. Scaled to
      // 1920 px wide, the whole page would be 1360 px high.
      const fixtures = join(import.meta.dir, '../../fixtures');
      const r = await scoreDeck(fakePage(1, undefined, white(1920, 1358)), {
        file: join(fixtures, 'slide-size.pptx'),
        name: 'slide-size',
        pdf: join(fixtures, 'slide-size.pdf'),
      });
      expect(r.reference).toBe('pdf');
      expect(r.slides[0].sizeMismatch).toBe(false);
      // A white render misses the fixture's shapes: the pixels are compared, not refused.
      expect(r.slides[0].pixelMismatch).toBeGreaterThan(0);
      expect(r.slides[0].pixelMismatch).toBeLessThan(1);
      // A render that is cut short fails on size, with the sizes.
      const cut = await scoreDeck(fakePage(1, undefined, white(1920, 1300)), {
        file: join(fixtures, 'slide-size.pptx'),
        name: 'slide-size',
        pdf: join(fixtures, 'slide-size.pdf'),
      });
      expect(cut.slides[0]).toMatchObject({
        status: 'fail',
        sizeMismatch: true,
        renderSizePx: [1920, 1300],
        referenceSizePx: [1920, 1358],
      });
      // No pixel score: the images were not compared.
      expect(cut.slides[0].pixelMismatch).toBeUndefined();
      expect(cut.slides[0].cause).toBe(
        'slide size differs: the render does not have its expected size'
      );
      expect(cut.slides[0].sizeDetail).toBe(
        'the render is 1920 × 1300 px, 1920 × 1358 px expected, and the PDF raster is 1920 × 1358 px, as expected'
      );
      // A size failure adds only measured line causes; the diagnostics cannot explain a size.
      const noisy = fakePage(
        1,
        async (code) => {
          if (code.includes('.open(')) return { slides: 1, substitutions: [] };
          if (code.includes('.reopenSaved(')) return { slides: 1 };
          if (code.includes('.render('))
            return { lines: [], diagnostics: ['font substituted: Calibri', 'approximation: blur'] };
          throw new Error(code);
        },
        white(1920, 1300)
      );
      const sized = await scoreDeck(noisy, {
        file: join(fixtures, 'slide-size.pptx'),
        name: 'slide-size',
        pdf: join(fixtures, 'slide-size.pdf'),
      });
      expect(sized.slides[0].cause).toBe(
        'slide size differs: the render does not have its expected size'
      );
      // A measured line cause follows the size cause.
      const withLine = fakePage(
        1,
        async (code) => {
          if (code.includes('.open(')) return { slides: 1, substitutions: [] };
          if (code.includes('.reopenSaved(')) return { slides: 1 };
          if (code.includes('.render('))
            return {
              lines: [{ node: 'n', text: 'A line', baselinePt: 100, leftPt: 0 }],
              diagnostics: ['font substituted: Calibri'],
            };
          throw new Error(code);
        },
        white(1920, 1300)
      );
      const both = await scoreDeck(withLine, {
        file: join(fixtures, 'slide-size.pptx'),
        name: 'slide-size',
        pdf: join(fixtures, 'slide-size.pdf'),
      });
      expect(both.slides[0].cause).toBe(
        'slide size differs: the render does not have its expected size; the PDF has no text for the laid-out line'
      );
    },
    // Four real Poppler scorings at the supersampled raster size: about 22 s alone, and more
    // than 30 s while the parallel suite runs every other file beside it.
    120_000
  );

  test.skipIf(!hasPoppler())(
    'a slide whose page the reference PDF rounded down expects the shorter raster',
    async () => {
      // 400.25 × 300.5 pt on a 400 × 301 pt page: the render is 1920 px wide, the raster 1919 px.
      const fixtures = join(import.meta.dir, '../../fixtures');
      const r = await scoreDeck(fakePage(1, undefined, white(1920, 1441)), {
        file: join(fixtures, 'slide-size-rounding.pptx'),
        name: 'slide-size-rounding',
        pdf: join(fixtures, 'slide-size-rounding.pdf'),
      });
      expect(r.slides[0]).toMatchObject({
        sizeMismatch: false,
        renderSizePx: [1920, 1441],
        referenceSizePx: [1919, 1441],
      });
    },
    // A real Poppler scoring; under load it takes more than the default 5 s.
    30_000
  );

  test.skipIf(!hasPoppler())(
    'a PDF that does not rasterize in time is not scored, and is no stale reference',
    async () => {
      const pdf = join(import.meta.dir, '../../fixtures/slide-structure.pdf');
      const four = join(dir, 'four.pptx');
      writeFileSync(
        four,
        buildPptx({
          width: 12_192_000,
          height: 6_858_000,
          slides: [{ shapes: '' }, { shapes: '' }, { shapes: '' }, { shapes: '' }],
        })
      );
      await expect(
        scoreDeck(fakePage(4), { file: four, name: 'd', pdf, rasterLimitMs: 1 })
      ).rejects.toThrow('Poppler did not finish within 0.001 s, so the deck is not scored');
    }
  );

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
