// Score one deck: print it to PDF with the editor's own paint, and compare our PDF with the
// reference PDF page by page (pixels, line breaks, baselines). pdf.js draws and reads both PDFs
// (`pdf.ts`). Then save, reopen, and paint again: every slide must paint the same markup. The
// export lane scores the PDF of `exportPdf` (`@pptx-editor.dev/pptx-to-pdf`) instead, with the
// same fonts and the same comparison; it has no save check. The reference PDF leaves hidden
// slides out, and so does our PDF. A missing or stale reference is a gap, never a pass. `run.ts`
// drives it.

import type { Page } from '@playwright/test';
import { copyFileSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { PAGE_SIZE_TOLERANCE_PT, THRESHOLDS } from './thresholds.ts';
import {
  openDeck,
  printDeck,
  rasterPdf,
  renderSlide,
  reopenSaved,
  ResetFailed,
  type HarnessControl,
  type HarnessRender,
} from './harness-page.ts';
import { comparePixels, compareLines, slidePasses } from './compare.ts';
import { redactPaths } from './corpus-list.ts';
import {
  deckFontFamilies,
  foreignOurFonts,
  foreignReferenceFonts,
  harnessFontFiles,
  localFonts,
  postScriptFamilies,
  type FontSet,
} from './local-fonts.ts';
import { collectionFace, type LocalFont } from './local-fonts.ts';
import { openPdf, openPdfFile, PdfPageLimit, rasterSizePx, type Pdf } from './pdf.ts';
import { deckFacts, referenceProblems, slidePages } from './reference-identity.ts';
import { probableCause, type DeckResult, type SlideResult } from './report.ts';
import type { Lane } from './run-args.ts';

/** One deck to score, with the PDF to compare it with, or the reason it has none. */
export interface DeckInput {
  readonly file: string;
  readonly name: string;
  /** A PDF to compare with. Without it, every slide is a gap. */
  readonly pdf?: string;
  /** Why the deck has no usable reference (missing, or the stale reasons). */
  readonly gap?: { readonly status: 'no-reference' | 'stale'; readonly reasons: readonly string[] };
  /**
   * The fonts the page gets besides the engine's own: the repository's (`fixtures`), or this
   * machine's reference fonts that the deck names (`corpus`, the default).
   */
  readonly fontSet?: FontSet;
  /**
   * Which PDF is ours: the editor's print (`editor`, the default), or the PDF of `exportPdf`
   * (`export`, the export lane).
   */
  readonly lane?: Lane;
}

/** What one lane gives the comparison: our PDF, and each slide's diagnostics and save check. */
interface OurPdf {
  readonly bytes?: Uint8Array;
  readonly diagnostics: readonly (readonly string[])[];
  /** Save, reopen, and paint again give the same markup; undefined where not checked. */
  readonly consistent: readonly (boolean | undefined)[];
}

/**
 * The fonts of a set as `exportPdf` takes them: the same faces, in the same order, that the
 * harness page gives the editor (`openDeck`). A collection gives the face alone, as the harness
 * server does.
 */
function exportFonts(fonts: readonly LocalFont[]) {
  return fonts.flatMap((f) => {
    const file = new Uint8Array(readFileSync(f.file));
    const bytes = f.faceIndex === undefined ? file : collectionFace(file, f.faceIndex);
    return bytes
      ? [{ family: f.family, bold: f.bold, italic: f.italic, bytes, origin: 'local' }]
      : [];
  });
}

/**
 * Export the visible slides with `exportPdf`, with the settings of the evaluation protocol
 * (`packages/pptx-to-pdf/scripts/evaluation/settings.ts`). Each slide's diagnostics read
 * `<severity>: <code>`. Throws when a visible slide has no page.
 */
async function exportDeck(
  deck: Uint8Array,
  fonts: readonly LocalFont[],
  slides: number,
  visible: readonly number[]
): Promise<OurPdf> {
  const { exportPdf } = await import('@pptx-editor.dev/pptx-to-pdf');
  const { EXPORT_SETTINGS, SESSION_SETTINGS } =
    await import('../../packages/pptx-to-pdf/scripts/evaluation/settings.ts');
  const result = await exportPdf(deck, {
    ...EXPORT_SETTINGS,
    ...SESSION_SETTINGS,
    fonts: exportFonts(fonts),
  });
  const pages = result.slideNumbers.map((n) => n - 1);
  if (pages.join() !== visible.join())
    throw new Error(
      `the export has pages for slides ${result.slideNumbers.join(', ') || 'none'}, the deck shows slides ${visible.map((i) => i + 1).join(', ')}`
    );
  const diagnostics = Array.from({ length: slides }, (_, i) =>
    result.diagnostics.filter((d) => d.slideIndex === i).map((d) => `${d.severity}: ${d.code}`)
  );
  return {
    bytes: visible.length ? result.bytes : undefined,
    diagnostics,
    consistent: diagnostics.map(() => undefined),
  };
}

/** A deck that takes longer than this is stopped and reported as an error. */
export const DECK_TIMEOUT_MS = 10 * 60_000;

/**
 * Print and score one deck. With `imagesDir`, both PDFs and each slide's two pages and diff are
 * written there; without it, nothing stays on disk. Throws when the page fails.
 */
export async function scoreDeck(
  page: Page,
  input: DeckInput,
  imagesDir?: string,
  /** The page that draws both PDFs (`HarnessControl.rasterPage`); `page` when absent. */
  rasterPage?: () => Promise<Page>
): Promise<DeckResult> {
  const started = Date.now();
  let facts;
  let deckBytes: Uint8Array;
  try {
    deckBytes = new Uint8Array(readFileSync(input.file));
    facts = deckFacts(deckBytes);
  } catch (e) {
    return { name: input.name, error: (e as Error).message, slides: [] };
  }
  const fixtures = input.fontSet === 'fixtures';
  const fonts = fixtures
    ? localFonts(undefined, 'fixtures')
    : localFonts(deckFontFamilies(deckBytes));
  const pages = slidePages(facts);
  const visible = pages.flatMap((p, i) => (p === undefined ? [] : [i]));
  let made: OurPdf;
  if (input.lane === 'export') made = await exportDeck(deckBytes, fonts, facts.slides, visible);
  else {
    const printed = await printWithEditor(page, input, fonts, facts.slides, visible);
    if (typeof printed === 'string') return { name: input.name, error: printed, slides: [] };
    made = printed;
  }
  const oursBytes = made.bytes;

  let gap = input.gap;
  if (!gap && !input.pdf) gap = { status: 'no-reference', reasons: ['no reference PDF'] };
  let reference: Pdf | undefined;
  let ours: Pdf | undefined;
  const result: DeckResult = { name: input.name, slides: [] };
  try {
    const size = rasterSizePx(facts, THRESHOLDS.renderWidthPx);
    const scale = THRESHOLDS.renderWidthPx / facts.widthPt;
    if (oursBytes) {
      ours = await openPdf(oursBytes, visible.length);
      if (ours.pages.length !== visible.length)
        throw new Error(`our PDF has ${ours.pages.length} pages for ${visible.length} slides`);
      // The same bound as the reference's pages (`referenceProblems`).
      const wrong = ours.pages.findIndex(
        ([w, h]) =>
          Math.abs(w - facts.widthPt) > PAGE_SIZE_TOLERANCE_PT + 1e-6 ||
          Math.abs(h - facts.heightPt) > PAGE_SIZE_TOLERANCE_PT + 1e-6
      );
      if (wrong >= 0)
        throw new Error(
          `our PDF page ${wrong + 1} is ${ours.pages[wrong][0]} × ${ours.pages[wrong][1]} pt, the slide is ${facts.widthPt} × ${facts.heightPt} pt`
        );
      // Every font of our PDF must be a face that the page loaded: never a system font.
      // `exportPdf` reads no system font, and it names each font it embeds `PptxFont`, so the
      // export lane checks only that every font is embedded.
      const loaded = postScriptFamilies(harnessFontFiles(fonts.map((f) => f.file)));
      const drawn = await ours.fonts();
      result.foreignFonts =
        input.lane === 'export'
          ? foreignOurFonts(
              drawn.filter((f) => !f.embedded),
              loaded
            )
          : foreignOurFonts(drawn, loaded);
    }
    if (!gap && input.pdf) {
      try {
        // More pages than slides is stale, and no page of such a PDF is read.
        reference = await openPdfFile(input.pdf, facts.slides);
        const issues = referenceProblems(facts, reference);
        const refFonts = await reference.fonts();
        for (const f of refFonts)
          if (!f.embedded) issues.push(`the PDF does not embed the font ${f.name}`);
        if (issues.length) gap = { status: 'stale', reasons: issues };
        // The editor lane checks the reference fonts; the export lane would repeat it.
        if (fixtures && input.lane !== 'export')
          result.foreignReferenceFonts = foreignReferenceFonts(
            refFonts,
            deckFontFamilies(deckBytes, false),
            postScriptFamilies(harnessFontFiles(fonts.map((f) => f.file)))
          );
      } catch (e) {
        // A fixed reason, never the tool's message: it can name the PDF's path.
        const reason =
          e instanceof PdfPageLimit
            ? `the PDF has ${e.pages} pages, the deck has ${visible.length} visible slides`
            : 'pdf.js cannot read the PDF';
        gap = { status: 'stale', reasons: [reason] };
      }
    }
    if (imagesDir) {
      mkdirSync(imagesDir, { recursive: true });
      if (oursBytes) writeFileSync((result.oursPdf = join(imagesDir, 'ours.pdf')), oursBytes);
      if (input.pdf && reference)
        copyFileSync(input.pdf, (result.referencePdf = join(imagesDir, 'reference.pdf')));
    }
    // Both PDFs are drawn by the same pdf.js in the harness page.
    const drawer = rasterPage ? await rasterPage() : page;
    const [oursPages, referencePages] =
      !gap && ours && reference && oursBytes && input.pdf
        ? [
            await rasterPdf(drawer, oursBytes, visible.length, size, scale),
            await rasterPdf(
              drawer,
              new Uint8Array(readFileSync(input.pdf)),
              visible.length,
              size,
              scale
            ),
          ]
        : [[], []];
    for (let i = 0; i < facts.slides; i += 1) {
      const pageIndex = pages[i];
      const diagnostics = made.diagnostics[i];
      const s: SlideResult = {
        index: i + 1,
        hidden: pageIndex === undefined,
        page: pageIndex === undefined ? undefined : pageIndex + 1,
        status: gap?.status ?? 'no-reference',
        cause: gap ? gap.reasons.join('; ') : '',
        ...(made.consistent[i] === undefined ? {} : { selfConsistent: made.consistent[i] }),
        diagnostics,
        fontsSubstituted: diagnostics
          .filter((d) => d.startsWith('font substituted: '))
          .map((d) => d.slice('font substituted: '.length)),
      };
      result.slides.push(s);
      if (pageIndex === undefined) {
        s.status = 'no-reference';
        s.cause = 'hidden slide: the reference PDF leaves it out';
        continue;
      }
      if (gap || !ours || !reference) continue;
      const [a, b] = [oursPages[pageIndex], referencePages[pageIndex]];
      const px = comparePixels(a, b, Boolean(imagesDir));
      const lines = compareLines(await ours.lines(pageIndex), await reference.lines(pageIndex));
      if (imagesDir) {
        writeFileSync((s.ours = join(imagesDir, `slide${i + 1}-ours.png`)), a.png);
        writeFileSync((s.reference = join(imagesDir, `slide${i + 1}-reference.png`)), b.png);
        if (px.diff) writeFileSync((s.diff = join(imagesDir, `slide${i + 1}-diff.png`)), px.diff);
      }
      s.pixelMismatch = px.mismatch;
      if (px.diffBox) {
        const pt = 1 / scale;
        s.diffBoxPt = px.diffBox.map((v) => Math.round(v * pt * 10) / 10) as [
          number,
          number,
          number,
          number,
        ];
      }
      // No reference line, no line score: `undefined`, never a perfect 1 or 0 pt.
      s.lineMatch = lines.total ? lines.matched / lines.total : undefined;
      s.linesMatched = lines.matched;
      s.linesTotal = lines.total;
      s.baselineDelta = lines.matched ? lines.worstBaseline : undefined;
      s.lineShift = lines.matched ? lines.worstShift : undefined;
      s.missingLines = lines.missing;
      s.extraLines = lines.extra;
      const pass = slidePasses(px, lines);
      s.status = pass ? 'pass' : 'fail';
      s.cause = pass ? '' : probableCause(s, diagnostics);
    }
  } finally {
    await ours?.close();
    await reference?.close();
  }
  result.reference = gap ? (gap.status === 'stale' ? 'stale' : 'none') : 'pdf';
  result.referenceIssues = gap?.reasons ?? [];
  result.milliseconds = Date.now() - started;
  return result;
}

/**
 * Open the deck in the editor, paint every slide, print the visible ones, then save, reopen, and
 * paint again. A string is the reason the deck cannot be scored.
 */
async function printWithEditor(
  page: Page,
  input: DeckInput,
  fonts: readonly LocalFont[],
  slides: number,
  visible: readonly number[]
): Promise<OurPdf | string> {
  const opened = await openDeck(page, input.file, [...fonts]);
  if (opened.error) return `the deck does not open: ${opened.error}`;
  if (opened.slides !== slides)
    return `the editor opens ${opened.slides} slides, the package has ${slides}`;
  const renders: HarnessRender[] = [];
  for (let i = 0; i < slides; i += 1) renders.push(await renderSlide(page, i));
  const bytes = visible.length ? await printDeck(page, [...visible]) : undefined;
  const reopened = await reopenSaved(page);
  const consistent: boolean[] = [];
  for (let i = 0; i < slides; i += 1) {
    const same =
      !reopened.error &&
      reopened.slides === slides &&
      (await renderSlide(page, i)).markup === renders[i].markup;
    // The markup is the same or not: a share of 0 or 1.
    consistent.push((same ? 0 : 1) <= THRESHOLDS.maxSelfConsistencyMismatch);
  }
  return { bytes, diagnostics: renders.map((r) => r.diagnostics), consistent };
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
  control: Pick<HarnessControl, 'reset'> & Partial<Pick<HarnessControl, 'rasterPage'>>,
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
    const result = await Promise.race([
      scoreDeck(page, input, options.imagesDir, control.rasterPage),
      timeout,
    ]);
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
