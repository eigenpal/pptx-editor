/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Copied from DOCX Editor `packages/docx-to-pdf/scripts/evaluation/layout-summary.ts`, commit
// 398963c7f67207cebf3c38dcc41d5f0bd1339b5f. PPTX changes: one page per visible slide, the
// page-to-slide map, the engine's layout notes per page, and the font plan (`--fonts`).
/** Layout-only evaluation through the same session as PDF export. */
import { readFile, writeFile } from 'node:fs/promises';
import type { FontSource } from '@pptx-editor.dev/core/layout';
import { PdfPageLimitError } from '../../src/index.ts';
import { openExportSession, sessionState } from '../../src/open-session.ts';
import { recordSlideText } from './layout-text.ts';
import {
  errorCode,
  EXPORT_SETTINGS,
  loadFontPlan,
  openFailure,
  parseArguments,
  protocolDeadline,
  PROTOCOL,
  SESSION_SETTINGS,
  type FontPlanIdentity,
} from './settings.ts';

const EMU_PER_PT = 12_700;
const MAX_PAGES = 10_000;
const MAX_NOTES = 200;
const USAGE = 'input.pptx output.json [--fonts plan.json]';

/** A file-derived note, bounded and without control characters. */
const clip = (text: string): string => text.replace(/[\u0000-\u001f\u007f]/g, ' ').slice(0, 200);

export async function summarizePages(
  source: Uint8Array,
  plan?: { fonts: FontSource[]; identity: FontPlanIdentity },
  timeoutMs: number = EXPORT_SETTINGS.timeoutMs
) {
  const started = performance.now();
  // One deadline for the open and every slide, as the export has (`timeoutMs` is for tests).
  const deadline = protocolDeadline(timeoutMs);
  const opened = await openExportSession(source, {
    ...SESSION_SETTINGS,
    ...(plan ? { fonts: plan.fonts } : {}),
  });
  if (!opened.ok) throw openFailure(opened);
  try {
    const openMs = performance.now() - started;
    const document = sessionState(opened.session)!.document;
    const info = document.info;
    // The same slides as the PDF: the reference PDF leaves hidden slides out.
    const shown = info.slides
      .map((slide, index) => ({ slide, index }))
      .filter(({ slide }) => EXPORT_SETTINGS.includeHiddenSlides || !slide.hidden);
    if (shown.length > MAX_PAGES) throw new PdfPageLimitError(MAX_PAGES, shown.length);
    const budget = { lines: 0, characters: 0 };
    const pages = [];
    const notLaidOut: number[] = [];
    for (const { index } of shown) {
      deadline.check();
      const layout = document.layout(index);
      if (!layout) {
        notLaidOut.push(index + 1);
        continue;
      }
      pages.push({
        page: pages.length + 1,
        slideNumber: index + 1,
        width: info.width / EMU_PER_PT,
        height: info.height / EMU_PER_PT,
        layoutNotes: layout.diagnostics.slice(0, MAX_NOTES).map(clip),
        lines: recordSlideText(layout, budget),
      });
    }
    // The last slide's layout counts against the deadline too.
    deadline.check();
    return {
      protocol: PROTOCOL,
      scope: 'layout-slides',
      pageCount: pages.length,
      slideNumbers: pages.map((p) => p.slideNumber),
      slideCount: info.slides.length,
      hiddenSlides: info.slides.flatMap((s, i) => (s.hidden ? [i + 1] : [])),
      notLaidOut,
      fontResolution: opened.session.fontResolution,
      fontPlan: plan?.identity ?? null,
      timings: { openMs, layoutMs: performance.now() - started - openMs },
      text: {
        version: 'layout-text-v1',
        coordinateSpace: 'slide-points-top-left',
        geometryStatus: 'before-paint',
        untrusted: true,
        pages: pages.map(({ page, slideNumber, lines }) => ({ page, slideNumber, lines })),
      },
      pages: pages.map(({ lines: _lines, ...page }) => page),
      textStatus: 'recorded-logical-layout',
      visualStatus: 'not-measured',
    };
  } finally {
    opened.session.dispose();
  }
}

if (import.meta.main) {
  try {
    const args = parseArguments(process.argv.slice(2), USAGE);
    const plan = args.fonts ? await loadFontPlan(args.fonts) : undefined;
    const result = await summarizePages(new Uint8Array(await readFile(args.input)), plan);
    await writeFile(args.output, JSON.stringify(result));
    console.log(
      JSON.stringify({ protocol: PROTOCOL, status: 'laid-out', pages: result.pageCount })
    );
  } catch (error) {
    const code = errorCode(error);
    console.log(
      JSON.stringify({
        protocol: PROTOCOL,
        status: 'error',
        error: error instanceof Error ? error.message : String(error),
        code,
      })
    );
    process.exitCode = 1;
  }
}
