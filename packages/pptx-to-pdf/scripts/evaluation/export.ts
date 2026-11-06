/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// After DOCX Editor `packages/docx-to-pdf/scripts/evaluation/export.ts`, commit
// 398963c7f67207cebf3c38dcc41d5f0bd1339b5f. PPTX changes: the font plan (`--fonts`), the
// page-to-slide map, the deck's hidden slides, a typed error `code`, one deadline for the open
// and the export, and `exportForEvaluation` for tests. It reads the slide list through the
// package's internal session state, because the public session has no slide list.
/** Eval protocol, version 1. No document-model API additions. */
import { readFile, writeFile } from 'node:fs/promises';
import type { FontSource } from '@pptx-editor.dev/core/layout';
import { exportPdfFrom } from '../../src/index.ts';
import { openExportSession, sessionState } from '../../src/open-session.ts';
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

const USAGE = 'input.pptx output.pdf [--fonts plan.json]';

/**
 * Export one deck with the protocol settings. Returns the stdout report. One deadline covers
 * the open and the export (`timeoutMs` is for tests).
 */
export async function exportForEvaluation(
  source: Uint8Array,
  plan?: { fonts: FontSource[]; identity: FontPlanIdentity },
  timeoutMs: number = EXPORT_SETTINGS.timeoutMs
) {
  const started = performance.now();
  const deadline = protocolDeadline(timeoutMs);
  const opened = await openExportSession(source, {
    ...SESSION_SETTINGS,
    ...(plan ? { fonts: plan.fonts } : {}),
  });
  if (!opened.ok) throw openFailure(opened);
  // exportPdfFrom does not open the deck, so its openMs is 0; the protocol reports the open.
  const openMs = Math.round(performance.now() - started);
  try {
    deadline.check();
    const slides = sessionState(opened.session)!.document.info.slides;
    const result = await exportPdfFrom(opened.session, {
      ...EXPORT_SETTINGS,
      timeoutMs: deadline.remaining(),
    });
    return {
      bytes: result.bytes,
      report: {
        protocol: PROTOCOL,
        status: 'exported' as const,
        pages: result.pageCount,
        slideNumbers: result.slideNumbers,
        slideCount: slides.length,
        hiddenSlides: slides.flatMap((s, i) => (s.hidden ? [i + 1] : [])),
        diagnostics: result.diagnostics,
        fontResolution: result.fontResolution,
        fontPlan: plan?.identity ?? null,
        timings: { ...result.timings, openMs },
      },
    };
  } finally {
    opened.session.dispose();
  }
}

if (import.meta.main) {
  try {
    const args = parseArguments(process.argv.slice(2), USAGE);
    const plan = args.fonts ? await loadFontPlan(args.fonts) : undefined;
    const { bytes, report } = await exportForEvaluation(
      new Uint8Array(await readFile(args.input)),
      plan
    );
    await writeFile(args.output, bytes);
    console.log(JSON.stringify(report));
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
