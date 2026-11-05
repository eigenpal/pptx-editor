// Score one deck in the harness page: render every slide with the editor's own pipeline, compare
// it with the deck's reference PDF (pixels, line breaks, baselines), and check save → reopen →
// re-render consistency. The reference PDF leaves hidden slides out, so PDF pages map to the
// visible slides in order. A missing or stale reference is a gap, never a pass. `run.ts` drives
// it for the fixture gate, the corpus report, and corpus mode.

import type { Page } from '@playwright/test';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { THRESHOLDS } from './thresholds.ts';
import {
  openDeck,
  renderSlide,
  reopenSaved,
  ResetFailed,
  type HarnessControl,
} from './harness-page.ts';
import {
  comparePixels,
  compareLines,
  missingFonts,
  sizeMismatchCause,
  sizeMismatchDetail,
  slidePasses,
} from './compare.ts';
import { redactPaths } from './corpus-list.ts';
import {
  expectedRasterSizePx,
  glyphOrigins,
  jpegPages,
  pdfLines,
  pdfPageFonts,
  POPPLER,
  rasterizePages,
  rasterLimitMs,
  renderSizePx,
  timedOut,
  withBaselines,
  type PdfLine,
  type ReferencePage,
} from './references.ts';
import { deckFacts, pdfFacts, referenceProblems, slidePages } from './reference-identity.ts';
import { measuredLineCauses, probableCause, type DeckResult, type SlideResult } from './report.ts';

/** One deck to score, with the PDF to compare it with, or the reason it has none. */
export interface DeckInput {
  readonly file: string;
  readonly name: string;
  /** A PDF to compare with. Without it, every slide is a gap. */
  readonly pdf?: string;
  /** Why the deck has no usable reference (missing, or the stale reasons). */
  readonly gap?: { readonly status: 'no-reference' | 'stale'; readonly reasons: readonly string[] };
  /** The time limit to rasterize the PDF (default: `rasterLimitMs` of its page count). */
  readonly rasterLimitMs?: number;
  /** A folder that keeps the PDF's rasters between runs (`rasterizePages`). */
  readonly rasterCache?: string;
}

/** A deck that takes longer than this is stopped and reported as an error. */
export const DECK_TIMEOUT_MS = 10 * 60_000;

/**
 * Render and score one deck. With `imagesDir`, the renders, the PDF rasters, and the diff
 * overlays are written there; without it, nothing stays on disk. Throws when the page fails.
 */
export async function scoreDeck(
  page: Page,
  input: DeckInput,
  imagesDir?: string
): Promise<DeckResult> {
  const started = Date.now();
  let facts;
  try {
    facts = deckFacts(new Uint8Array(readFileSync(input.file)));
  } catch (e) {
    return { name: input.name, error: (e as Error).message, slides: [] };
  }
  const opened = await openDeck(page, input.file);
  if (opened.error)
    return { name: input.name, error: `the deck does not open: ${opened.error}`, slides: [] };
  if (opened.slides !== facts.slides)
    return {
      name: input.name,
      error: `the editor opens ${opened.slides} slides, the package has ${facts.slides}`,
      slides: [],
    };

  let gap = input.gap;
  if (!gap && !input.pdf) gap = { status: 'no-reference', reasons: ['no reference PDF'] };
  const work = imagesDir ?? mkdtempSync(join(tmpdir(), 'pptx-fidelity-'));
  const slides: SlideResult[] = [];
  try {
    mkdirSync(work, { recursive: true });
    let referencePages: ReferencePage[] = [];
    let referenceLines: PdfLine[][] = [];
    let jpeg = new Set<number>();
    let pdfPages: readonly (readonly [number, number])[] = [];
    if (!gap && input.pdf) {
      // Fixed reasons, never a tool's message: it names the PDF's path.
      let step = 'the PDF does not open (pdfinfo failed)';
      let limitMs: number = POPPLER.timeout;
      try {
        const info = pdfFacts(input.pdf);
        pdfPages = info.pages;
        const issues = referenceProblems(facts, info);
        if (issues.length) gap = { status: 'stale', reasons: issues };
        else {
          step = 'the PDF does not rasterize (pdftoppm failed)';
          limitMs = input.rasterLimitMs ?? rasterLimitMs(info.pages.length);
          referencePages = rasterizePages(
            input.pdf,
            join(work, 'reference'),
            facts,
            THRESHOLDS.renderWidthPx,
            limitMs,
            THRESHOLDS.rasterScale,
            input.rasterCache
          );
          limitMs = POPPLER.timeout;
          step = 'the PDF text does not extract (pdftotext failed)';
          // Baselines are glyph origins from the content stream. When pdf.js cannot read the
          // PDF, they fall back to estimates from the line boxes, and the report counts them.
          const origins = await glyphOrigins(input.pdf).catch(() => {
            console.warn(`${input.name}: pdf.js cannot read the PDF; baselines are estimates.`);
            return [];
          });
          referenceLines = withBaselines(pdfLines(input.pdf, work), origins);
          step = 'the PDF images do not list (pdfimages failed)';
          jpeg = jpegPages(input.pdf);
        }
      } catch (e) {
        // A time-out says nothing about the reference: the deck is not scored, and is no gap.
        if (timedOut(e))
          throw new Error(
            `Poppler did not finish within ${limitMs / 1000} s, so the deck is not scored. The machine may be busy: run the deck again`
          );
        gap = { status: 'stale', reasons: [step] };
      }
    }
    const pages = slidePages(facts);
    const firstRenders: Buffer[] = [];
    for (let i = 0; i < opened.slides; i += 1) {
      const r = await renderSlide(page, i, THRESHOLDS.renderWidthPx);
      firstRenders.push(r.devicePng);
      const render = imagesDir ? join(imagesDir, `slide${i + 1}.png`) : undefined;
      if (render) writeFileSync(render, r.png);
      const pageIndex = pages[i];
      const result: SlideResult = {
        index: i + 1,
        hidden: pageIndex === undefined,
        page: pageIndex === undefined ? undefined : pageIndex + 1,
        render,
        diagnostics: r.diagnostics,
        status: gap?.status ?? 'no-reference',
        cause: gap ? gap.reasons.join('; ') : '',
      };
      if (pageIndex === undefined) {
        result.status = 'no-reference';
        result.cause = 'hidden slide: the reference PDF leaves it out';
      } else if (!gap && !referencePages[pageIndex]) {
        result.cause = `the PDF raster has no page ${pageIndex + 1}`;
      } else if (!gap) {
        const reference = referencePages[pageIndex];
        const px = comparePixels(r.image, reference.image, {
          diff: Boolean(imagesDir),
          // Anti-aliasing is judged on the unfiltered images (`supersample.ts`).
          device:
            THRESHOLDS.rasterScale > 1 && THRESHOLDS.antialiasing === 'device'
              ? {
                  render: r.device,
                  reference: reference.device(),
                  scale: THRESHOLDS.rasterScale,
                }
              : undefined,
          expected: {
            render: renderSizePx(facts, THRESHOLDS.renderWidthPx),
            reference: expectedRasterSizePx(facts, pdfPages[pageIndex], THRESHOLDS.renderWidthPx),
          },
        });
        const lines = compareLines(r.lines, referenceLines[pageIndex] ?? []);
        if (imagesDir && px.diff) {
          result.diff = join(imagesDir, `slide${i + 1}-diff.png`);
          writeFileSync(result.diff, px.diff);
        }
        result.referenceImage = imagesDir ? referencePages[pageIndex].path : undefined;
        result.referenceJpeg = jpeg.has(pageIndex + 1);
        // A size mismatch compares no pixels: no score, never a 100 % mismatch.
        result.pixelMismatch = px.sizeMismatch ? undefined : px.mismatch;
        // No checked line, no line score: `undefined`, never a perfect 1 or 0 pt.
        result.lineMatch = lines.total ? lines.matched / lines.total : undefined;
        result.linesMatched = lines.matched;
        result.linesTotal = lines.total;
        result.pdfLinesUnmatched = lines.pdfUnmatched;
        result.baselineDelta = lines.matched ? lines.worstBaseline : undefined;
        result.estimatedBaselines = lines.estimated;
        result.skippedLines = lines.skipped;
        result.unmatchedLines = lines.unmatched.length;
        result.unmatchedTexts = lines.unmatched;
        result.pdfUnmatchedTexts = lines.pdfUnmatchedTexts;
        result.lineSplitOnly = lines.splitOnly;
        result.lineMarkersOnly = lines.markersOnly;
        result.lineSameText = lines.sameText;
        result.sizeMismatch = px.sizeMismatch;
        result.renderSizePx = [px.sizes.render.width, px.sizes.render.height];
        result.referenceSizePx = [px.sizes.reference.width, px.sizes.reference.height];
        const want = px.sizes.expected;
        if (want) {
          result.expectedRenderSizePx = [want.render.width, want.render.height];
          result.expectedReferenceSizePx = [want.reference.width, want.reference.height];
        }
        if (px.sizeMismatch) result.sizeDetail = sizeMismatchDetail(px.sizes);
        if (px.diffBox) {
          // Pixels to points: the render is `renderWidthPx` wide for the slide's width.
          const pt = facts.widthPt / THRESHOLDS.renderWidthPx;
          result.diffBoxPt = px.diffBox.map((v) => Math.round(v * pt * 10) / 10) as [
            number,
            number,
            number,
            number,
          ];
        }
        // Fonts: what the reference PDF drew with (the PDF page) against what the slide asks for,
        // and what the editor substituted. Either side's substitution is reported apart from
        // layout.
        let pdfFonts: string[] = [];
        try {
          pdfFonts = pdfPageFonts(input.pdf!, pageIndex + 1);
        } catch {
          // pdffonts failed: the fonts are unknown, not substituted.
        }
        result.fontsReference = pdfFonts;
        result.fontsSubstituted = {
          reference: pdfFonts.length ? missingFonts(r.fonts, pdfFonts) : [],
          editor: r.diagnostics
            .filter((d) => d.startsWith('font substituted: '))
            .map((d) => d.slice('font substituted: '.length)),
        };
        const pass = slidePasses(px, lines);
        result.status = pass ? 'pass' : 'fail';
        result.cause = pass
          ? ''
          : px.sizeMismatch
            ? // A size failure still measures lines: add only their measured causes, never the
              // diagnostics, which cannot explain a size.
              [sizeMismatchCause(px.sizes), ...measuredLineCauses(result)].join('; ')
            : probableCause(result, r.diagnostics);
      }
      slides.push(result);
    }
    // Save → reopen → re-render: must be pixel-identical to the first render, at the device
    // scale, before the filter can hide a change (`supersample.ts`).
    const reopened = await reopenSaved(page);
    if (reopened.error || reopened.slides !== opened.slides) {
      for (const s of slides) s.selfConsistent = false;
    } else {
      for (let i = 0; i < opened.slides; i += 1) {
        const again = (await renderSlide(page, i, THRESHOLDS.renderWidthPx)).devicePng;
        // The same PNG bytes are the same pixels: no decode is needed.
        slides[i].selfConsistent =
          firstRenders[i].equals(again) ||
          comparePixels(firstRenders[i], again).mismatch <= THRESHOLDS.maxSelfConsistencyMismatch;
      }
    }
  } finally {
    if (!imagesDir) rmSync(work, { recursive: true, force: true });
  }
  return {
    name: input.name,
    slides,
    reference: gap ? (gap.status === 'stale' ? 'stale' : 'none') : 'pdf',
    referenceIssues: gap?.reasons ?? [],
    milliseconds: Date.now() - started,
  };
}

/** The result of `scoreSafely`. */
export interface SafeScore {
  readonly result: DeckResult;
  /** The page for the next deck: the same one, or a new one after a failure. */
  readonly page: Page;
  /**
   * Not even a new browser could open the harness page after a failure. Record `result` first,
   * then stop the run with this error: no later deck can be scored.
   */
  readonly fatal?: Error;
}

/** The start of a deck's error when the browser process exited during the deck. */
export const BROWSER_CRASHED = 'the browser crashed while rendering the deck';
/** The start of a deck's error when the browser stopped answering during the deck. */
export const BROWSER_HUNG = 'the browser stopped answering while rendering the deck';

/**
 * Score one deck within `DECK_TIMEOUT_MS`, so that one deck cannot stop a run. An error or a
 * timeout becomes the deck's `error`, and the page is replaced by a new one, in a new browser
 * when the old one crashed or hung. With `redact`, the error keeps no path (corpus output).
 */
export async function scoreSafely(
  page: Page,
  control: Pick<HarnessControl, 'reset'>,
  input: DeckInput,
  options: {
    readonly imagesDir?: string;
    readonly redact: boolean;
    /** Tests only: a shorter time limit. */
    readonly timeoutMs?: number;
  }
): Promise<SafeScore> {
  const timeoutMs = options.timeoutMs ?? DECK_TIMEOUT_MS;
  let timer: ReturnType<typeof setTimeout> | undefined;
  const timeout = new Promise<never>((_, reject) => {
    timer = setTimeout(
      () => reject(new Error(`the deck did not finish within ${timeoutMs / 1000} s`)),
      timeoutMs
    );
  });
  try {
    const result = await Promise.race([scoreDeck(page, input, options.imagesDir), timeout]);
    return {
      result:
        options.redact && result.error ? { ...result, error: redactPaths(result.error) } : result,
      page,
    };
  } catch (e) {
    const message = (e as Error).message ?? String(e);
    const error = options.redact ? redactPaths(message) : (message.split('\n')[0] ?? message);
    const failed = (cause = 'the harness failed'): DeckResult => ({
      name: input.name,
      error: `${cause}: ${error}`,
      slides: [],
    });
    // Playwright's message ("browser has been closed") does not say that Chromium died, so the
    // cause comes first.
    const cause = (newBrowser: 'exited' | 'hung' | undefined, next: string): string | undefined =>
      newBrowser === undefined
        ? undefined
        : `${newBrowser === 'exited' ? BROWSER_CRASHED : BROWSER_HUNG} (${next})`;
    try {
      const reset = await control.reset();
      return {
        result: failed(cause(reset.newBrowser, 'the next deck gets a new browser')),
        page: reset.page,
      };
    } catch (resetError) {
      const known = resetError instanceof ResetFailed ? resetError.newBrowser : undefined;
      return {
        result: failed(cause(known, 'no new browser opened, so the run stops')),
        page,
        fatal: resetError as Error,
      };
    }
  } finally {
    clearTimeout(timer);
  }
}
