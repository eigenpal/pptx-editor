// The fidelity HTML report: per slide, the reference page, our render, and the diff overlay, with
// pass/fail, the probable cause, and the approximations and diagnostics that apply. Written to
// fidelity-report/ (gitignored: corpus renders contain private content).

import { writeFileSync } from 'node:fs';
import { join, relative } from 'node:path';
import { APPROXIMATIONS } from '../../packages/core/src/layout/approximations.ts';
import { BASELINE_ESTIMATE_OFFSET_PT, THRESHOLDS } from './thresholds.ts';
import { isListMarker, sizeLimitText } from './compare.ts';
import { gateExitCode } from './known-gaps.ts';

const APPROXIMATION_PREFIX = 'approximation: ';

/** The diagnostics after a prefix, without it (`effect not rendered: blur` → `blur`). */
const tagged = (diagnostics: readonly string[], prefix: string): string[] =>
  diagnostics.filter((d) => d.startsWith(prefix)).map((d) => d.slice(prefix.length));

/** The titles (`APPROXIMATIONS`) of the approximations a slide's diagnostics name. */
export function approximationTitles(diagnostics: readonly string[]): string[] {
  return tagged(diagnostics, APPROXIMATION_PREFIX).map(
    (id) => APPROXIMATIONS.find((a) => a.id === id)?.title ?? id
  );
}

/** Diagnostics that are not approximations: substitutions, unrendered content. */
export function otherDiagnostics(diagnostics: readonly string[]): string[] {
  return diagnostics.filter((d) => !d.startsWith(APPROXIMATION_PREFIX));
}

/** The approximations layout cannot detect on a slide: they may apply to any slide. */
export function undetectedApproximations(): string[] {
  return APPROXIMATIONS.filter((a) => !a.perSlide).map((a) => a.title);
}

/**
 * Whether every laid-out line that the PDF does not hold is in the PDF after a short bullet or
 * number (`•`, `1.`, `a.`, `iv.`): then only the marker differs, not the line break.
 */
export function onlyListMarkers(r: SlideResult): boolean {
  // `compareLines` reads the full texts; the 60-character texts are a fallback for old results.
  if (r.lineMarkersOnly !== undefined) return r.lineMarkersOnly;
  const ours = r.unmatchedTexts ?? [];
  const pdf = (r.pdfUnmatchedTexts ?? []).map((t) => t.replace(/\s+/g, ''));
  return (
    ours.length > 0 &&
    ours.every((line) => {
      const bare = line.replace(/\s+/g, '');
      return pdf.some(
        (p) => p.length > bare.length && p.endsWith(bare) && isListMarker(p.slice(0, -bare.length))
      );
    })
  );
}

/**
 * Why the laid-out lines that the PDF does not hold differ. Each cause is 80 characters or less,
 * because `eigenpal/pptx-eval` keeps the first 80.
 */
function lineCause(r: SlideResult): string {
  if (onlyListMarkers(r))
    return 'bullet or number: only the PDF line holds it (a limit of the line check)';
  const pdf = r.pdfUnmatchedTexts ?? [];
  if (!pdf.length) return 'the PDF has no text for the laid-out line';
  if (r.lineSplitOnly)
    return 'the PDF splits the same text into other words (a limit of the line check)';
  // The same characters in other lines: the paragraph breaks at other places.
  const chars = (texts: readonly string[]): string =>
    [...texts.join('').replace(/\s+/g, '')].sort().join('');
  if (r.lineSameText ?? chars(r.unmatchedTexts ?? []) === chars(pdf)) return 'line breaks differ';
  return 'line breaks or text differ from the PDF';
}

/** A slide's causes without its size cause, which the verdict already names. */
function otherCauses(s: SlideResult): string {
  const causes = s.cause ? s.cause.split('; ') : [];
  return (s.sizeMismatch ? causes.filter((c) => !c.startsWith('slide size differs')) : causes).join(
    '; '
  );
}

/** The causes of a slide's failed line scores, from its measurements alone (no diagnostics). */
export function measuredLineCauses(r: SlideResult): string[] {
  const causes: string[] = [];
  if (r.lineMatch !== undefined && r.lineMatch < THRESHOLDS.minLineBreakMatch)
    causes.push(lineCause(r));
  if (r.baselineDelta !== undefined && r.baselineDelta > THRESHOLDS.maxBaselineDeltaPt)
    causes.push(
      // Glyph origins (no estimated baseline) are exact, so only the model is left.
      r.estimatedBaselines === 0
        ? 'line height or baseline model'
        : // The bound is a value rounded to 2 decimals, as the gate records baselines.
          Math.round(r.baselineDelta * 100) / 100 <= BASELINE_ESTIMATE_OFFSET_PT
          ? 'line height or baseline model, or the PDF baseline estimate'
          : `a baseline is off by more than the largest estimate offset seen (${BASELINE_ESTIMATE_OFFSET_PT} pt)`
    );
  return causes;
}

/** The probable cause of a failing slide, from its measurements and its layout diagnostics. */
export function probableCause(r: SlideResult, diagnostics: readonly string[]): string {
  const causes: string[] = [];
  if (diagnostics.some((d) => d.startsWith('font substituted'))) causes.push('font substitution');
  if (diagnostics.some((d) => d.startsWith('graphic frame')))
    causes.push('unrendered graphic frame (chart, diagrams, OLE)');
  if (diagnostics.some((d) => d.startsWith('unknown preset')))
    causes.push('unknown preset geometry');
  if (
    r.referenceJpeg &&
    r.pixelMismatch !== undefined &&
    r.pixelMismatch > THRESHOLDS.maxPixelMismatch
  )
    causes.push('JPEG images in the reference PDF');
  if (r.fontsSubstituted?.reference.length)
    causes.push(
      `font substitution in the reference PDF (${r.fontsSubstituted.reference.join(', ')})`
    );
  const effects = tagged(diagnostics, 'effect not rendered: ');
  if (effects.length) causes.push(`effect not rendered (${effects.join(', ')})`);
  const lineCauses = measuredLineCauses(r);
  causes.push(...lineCauses);
  // Every text body uses the line model. It moves baselines only, so it is a cause only when a
  // line check fails.
  const approximations = approximationTitles(
    lineCauses.length
      ? diagnostics
      : diagnostics.filter((d) => d !== `${APPROXIMATION_PREFIX}line-model`)
  );
  if (approximations.length) causes.push(`approximations: ${approximations.join(', ')}`);
  if (
    causes.length === 0 &&
    r.pixelMismatch !== undefined &&
    r.pixelMismatch > THRESHOLDS.maxPixelMismatch
  )
    causes.push('pixels differ, and no layout diagnostic names a cause: see the diff area');
  return causes.join('; ');
}

export interface SlideResult {
  index: number;
  /** A hidden slide has no PDF page. */
  hidden?: boolean;
  /** Its page in the reference PDF (1-based). */
  page?: number;
  render?: string;
  referenceImage?: string;
  /** The reference page holds JPEG images, so its pixels include JPEG loss (`jpegPages`). */
  referenceJpeg?: boolean;
  diff?: string;
  /** `stale`: the reference no longer matches the deck. Like `no-reference`, a gap. */
  status: 'pass' | 'fail' | 'no-reference' | 'stale';
  /** A failing fixture slide listed in `FIXTURE_KNOWN_GAPS`. It is still a failure. */
  knownGap?: boolean;
  /** The scores a failing slide misses, and its known gap (`slideVerdict` in known-gaps.ts). */
  verdict?: string;
  /**
   * The render or the PDF raster differs from its expected size (`renderSizePx`,
   * `expectedRasterSizePx` in references.ts) by more than `maxSizeDeltaPx`. Then the pixels are
   * not compared, and `pixelMismatch` is undefined.
   */
  sizeMismatch?: boolean;
  /** The render's and the PDF raster's size, `[width, height]` in pixels. */
  renderSizePx?: readonly [number, number];
  referenceSizePx?: readonly [number, number];
  /** The size each must have (`renderSizePx`, `expectedRasterSizePx` in references.ts). */
  expectedRenderSizePx?: readonly [number, number];
  expectedReferenceSizePx?: readonly [number, number];
  /** On a size mismatch: each size and the size it must have (`sizeMismatchDetail`). */
  sizeDetail?: string;
  pixelMismatch?: number;
  /** `pdftotext` split each unmatched line into other words on its row (`LineScore`). */
  lineSplitOnly?: boolean;
  /** Each unmatched line differs from a PDF line only by its bullet or number (`LineScore`). */
  lineMarkersOnly?: boolean;
  /** The unmatched lines of both sides hold the same characters (`LineScore`). */
  lineSameText?: boolean;
  /** Matched share of the laid-out lines; undefined when the slide has no checked line. */
  lineMatch?: number;
  linesMatched?: number;
  linesTotal?: number;
  /** PDF text lines that no laid-out line matched. Reported, not gated (`CONTRIBUTING.md`, "Fidelity harness"). */
  pdfLinesUnmatched?: number;
  /** Worst baseline distance of a matched line; undefined when no line matched. */
  baselineDelta?: number;
  /** Matched lines whose PDF baseline was estimated from the line box (no glyph origin). */
  estimatedBaselines?: number;
  /** Lines that start outside the slide and that the PDF does not hold: not checked. */
  skippedLines?: number;
  unmatchedLines?: number;
  /** Laid-out lines that the PDF does not hold, and PDF lines that no laid-out line matched. */
  unmatchedTexts?: readonly string[];
  pdfUnmatchedTexts?: readonly string[];
  /** The box around every differing pixel, `[x0, y0, x1, y1]` in points from the top-left. */
  diffBoxPt?: readonly [number, number, number, number];
  /** The fonts of the reference PDF page. */
  fontsReference?: readonly string[];
  /**
   * Requested families that the reference PDF draws with another font (absent from its PDF page),
   * and families that the editor substituted. Reported apart from layout gaps.
   */
  fontsSubstituted?: { readonly reference: readonly string[]; readonly editor: readonly string[] };
  selfConsistent?: boolean;
  diagnostics: readonly string[];
  cause: string;
}

export interface DeckResult {
  name: string;
  error?: string;
  reference?: 'pdf' | 'none' | 'stale';
  /** Why the deck has no usable reference. */
  referenceIssues?: readonly string[];
  milliseconds?: number;
  /** The reference rasters came from the raster cache. */
  referenceCached?: boolean;
  slides: SlideResult[];
}

/** What the report says about the run besides the slides. */
export interface ReportContext {
  /** The Poppler version that rasterized the PDFs, or null when Poppler is not installed. */
  readonly poppler: string | null;
  readonly chromium: string | null;
  /** The fixture gate, when this run is gated. */
  readonly gate?: {
    /** The machine: platform, and whether the reference fonts are installed. */
    readonly where: string;
    /** The known-gap list that applies, or undefined when this environment has none. */
    readonly list?: string;
    readonly problems: readonly string[];
    readonly notes: readonly string[];
    /** The known-gap entries that this run's scores record (`--print-gaps`). */
    readonly entries: readonly string[];
  };
}

const esc = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const pct = (v: number | undefined): string =>
  v === undefined ? '—' : `${(v * 100).toFixed(2)} %`;

const list = (label: string, items: readonly string[]): string =>
  items.length ? `<p class="diag"><b>${label}:</b> ${items.map(esc).join(' · ')}</p>` : '';

function gateSection(gate: NonNullable<ReportContext['gate']>): string {
  const verdict = gate.problems.length
    ? `<p class="fail"><b>Gate: fails</b> with ${gate.problems.length} problem(s).</p>`
    : gate.list
      ? '<p class="ok"><b>Gate: passes.</b></p>'
      : `<p class="fail"><b>Gate: slide scores not gated.</b> There is no known-gap list for ${esc(gate.where)}, so only references, open errors, and save consistency were gated. Run the gate in CI or on macOS with the reference fonts.</p>`;
  const items = (label: string, lines: readonly string[]): string =>
    lines.length
      ? `<p><b>${label}</b></p><ul>${lines.map((l) => `<li>${esc(l)}</li>`).join('')}</ul>`
      : '';
  return (
    `<section class="gate"><h2>Gate</h2>${verdict}` +
    `<p>Known gaps: ${gate.list ? `<code>${esc(gate.list)}</code>` : 'no list for this environment, so slide failures are reported, not gated'} (<code>scripts/fidelity/thresholds.ts</code>).</p>` +
    items('Problems', gate.problems) +
    items('Notes', gate.notes) +
    `<details><summary>Known-gap entries that this run records</summary><pre>${esc(gate.entries.join('\n'))}</pre></details></section>`
  );
}

const lineScore = (s: SlideResult): string =>
  s.linesTotal
    ? `line breaks ${s.linesMatched} of ${s.linesTotal}${s.pdfLinesUnmatched ? `, ${s.pdfLinesUnmatched} PDF line(s) not laid out` : ''}`
    : `line breaks not checked${s.pdfLinesUnmatched ? ` (${s.pdfLinesUnmatched} PDF line(s) not laid out)` : ''}`;

export function writeReport(
  out: string,
  set: string,
  decks: DeckResult[],
  context: ReportContext
): void {
  const all = decks.flatMap((d) => d.slides);
  const count = (s: SlideResult['status']): number => all.filter((x) => x.status === s).length;
  const consistent = all.filter((s) => s.selfConsistent).length;
  const rows = decks
    .map((d) => {
      if (d.error)
        return `<section><h2>${esc(d.name)}</h2><p class="fail">Open failed: ${esc(d.error)}</p></section>`;
      const slides = d.slides
        .map((s) => {
          const img = (p: string | undefined, label: string): string =>
            p
              ? `<figure><img loading="lazy" src="${esc(relative(out, p))}" alt="${label}"><figcaption>${label}</figcaption></figure>`
              : `<figure class="empty"><div>no ${label.toLowerCase()}</div><figcaption>${label}</figcaption></figure>`;
          return (
            `<article class="${s.status}"><header><b>Slide ${s.index}</b> <span class="badge ${s.status}">${s.knownGap ? 'fail (known gap)' : s.status}</span>` +
            ` <span>pixels ${s.sizeMismatch ? 'not compared (size differs)' : pct(s.pixelMismatch)}</span> <span>${lineScore(s)}${s.skippedLines ? ` (${s.skippedLines} off the slide, not checked)` : ''}</span>` +
            ` <span>baseline Δ ${s.baselineDelta === undefined || !s.linesMatched ? 'not checked' : `${s.baselineDelta.toFixed(2)} pt${s.estimatedBaselines ? ` (${s.estimatedBaselines} estimated)` : ''}`}</span>` +
            ` <span class="${s.selfConsistent ? 'ok' : 'fail'}">save/reopen ${s.selfConsistent ? 'identical' : 'DIFFERS'}</span></header>` +
            (s.verdict ? `<p class="cause">${esc(s.verdict)}</p>` : '') +
            (s.referenceJpeg
              ? '<p class="diag">The reference page holds JPEG images, so its pixels include JPEG loss.</p>'
              : '') +
            (s.diffBoxPt
              ? `<p class="diag">Pixels differ inside x ${s.diffBoxPt[0]}–${s.diffBoxPt[2]} pt, y ${s.diffBoxPt[1]}–${s.diffBoxPt[3]} pt (from the top-left).</p>`
              : '') +
            list(
              'Laid-out lines that the PDF does not hold',
              (s.unmatchedTexts ?? []).slice(0, 8)
            ) +
            list(
              'PDF lines that no laid-out line matched',
              (s.pdfUnmatchedTexts ?? []).slice(0, 8)
            ) +
            list('Font substitution (reported apart from layout gaps)', [
              ...(s.fontsSubstituted?.reference ?? []).map((f) => `reference: ${f}`),
              ...(s.fontsSubstituted?.editor ?? []).map((f) => `editor: ${f}`),
            ]) +
            // A size failure is measured, and the verdict names it; its other causes stay.
            (otherCauses(s) ? `<p class="cause">Probable cause: ${esc(otherCauses(s))}</p>` : '') +
            list('Approximations', approximationTitles(s.diagnostics)) +
            list('Diagnostics', otherDiagnostics(s.diagnostics)) +
            `<div class="imgs">${img(s.referenceImage, 'Reference')}${img(s.render, 'PPTX Editor')}${s.diff || s.pixelMismatch === undefined ? img(s.diff, 'Diff') : '<figure class="empty"><div>no differing pixels</div><figcaption>Diff</figcaption></figure>'}</div></article>`
          );
        })
        .join('');
      const reference =
        d.reference === 'pdf'
          ? 'reference PDF'
          : `${d.reference === 'stale' ? 'stale reference' : 'no reference'}: ${(d.referenceIssues ?? []).join('; ')}`;
      return `<section><h2>${esc(d.name)} <small>${esc(reference)}</small></h2>${slides}</section>`;
    })
    .join('');
  const tools = `Poppler ${context.poppler ?? 'not installed'}, Chromium ${context.chromium ?? 'unknown'}`;
  const html = `<!doctype html><html lang="en"><head><meta charset="utf-8"><title>Fidelity report: ${esc(set)}</title>
<style>
body{font:14px/1.5 -apple-system,system-ui,sans-serif;margin:24px;color:#1b1a19;background:#f4f3f0}
h1{font-size:20px}h2{font-size:16px;margin-top:32px}small{color:#75716b;font-weight:400}
.summary span{display:inline-block;margin-right:16px}
article,.gate{background:#fff;border-radius:12px;padding:12px 16px;margin:12px 0;box-shadow:0 0 0 1px #e7e4df}
.gate h2{margin-top:0}pre{white-space:pre-wrap;font-size:12px}
article header span{margin-left:12px;color:#44413d}
.badge{border-radius:999px;padding:1px 8px;font-size:12px;font-weight:600}
.badge.pass{background:#e6f4ea;color:#1e7a34}.badge.fail{background:#fbe9e7;color:#a8432a}.badge.no-reference,.badge.stale{background:#eceae6;color:#44413d}
.fail{color:#a8432a}.ok{color:#1e7a34}.cause{color:#a8432a;margin:6px 0}.diag{color:#75716b;font-size:12px;margin:4px 0}
.imgs{display:grid;grid-template-columns:repeat(3,1fr);gap:8px}figure{margin:0}img{width:100%;border:1px solid #e7e4df}
figcaption{font-size:12px;color:#75716b}.empty div{aspect-ratio:16/9;display:grid;place-items:center;border:1px dashed #d9d5ce;color:#a29e97}
</style></head><body><h1>Fidelity report: ${esc(set)}</h1>
<p class="summary"><span>Slides: ${all.length}</span><span>With a font substitution: ${all.filter((x) => x.fontsSubstituted && (x.fontsSubstituted.reference.length || x.fontsSubstituted.editor.length)).length}</span><span>Pass: ${count('pass')}</span><span>Fail: ${count('fail')} (known gaps: ${all.filter((x) => x.knownGap).length})</span><span>No reference: ${count('no-reference')}</span><span>Stale reference: ${count('stale')}</span><span>Save/reopen identical: ${consistent}/${all.length}</span>${context.poppler ? '' : '<span class="fail">Poppler not installed: reference PDFs not compared</span>'}</p>
<p>Thresholds: pixels ≤ ${pct(THRESHOLDS.maxPixelMismatch)}, line breaks ${pct(THRESHOLDS.minLineBreakMatch)} identical, baseline ≤ ${THRESHOLDS.maxBaselineDeltaPt} pt, slide size ${sizeLimitText()} (<code>scripts/fidelity/thresholds.ts</code>). ${esc(tools)}.</p>
<p class="diag">Approximations that layout cannot detect on a slide, so any slide may use them (<code>docs/site/content/pptx-fidelity.mdx</code>): ${undetectedApproximations().map(esc).join(' · ')}.</p>
<p class="diag">Limits of the checks (<code>CONTRIBUTING.md</code>, "Fidelity harness"): the PDF baseline is an estimate; geometry is checked by pixels only; text in tables, rotated text, and vertical text is checked by pixels only; the line check asks only whether each laid-out line is in the PDF, and a laid-out line has no bullet or number while the PDF line has it; the reference PDFs store pictures and rasterized effects as JPEG, so those pixels include JPEG loss; the slide past a page side that the reference PDF rounded down is not compared; and the harness has the reference fonts only, so a font that the reference PDF took from macOS is a substitution here.</p>
${context.gate ? gateSection(context.gate) : ''}
${rows}</body></html>`;
  writeFileSync(join(out, 'index.html'), html);
  writeFileSync(
    join(out, 'summary.json'),
    JSON.stringify(
      decks.map((d) => ({
        name: d.name,
        error: d.error,
        reference: d.reference,
        referenceIssues: d.referenceIssues,
        slides: d.slides.map((s) => ({
          index: s.index,
          status: s.status,
          knownGap: s.knownGap,
          verdict: s.verdict,
          sizeMismatch: s.sizeMismatch,
          renderSizePx: s.renderSizePx,
          referenceSizePx: s.referenceSizePx,
          expectedRenderSizePx: s.expectedRenderSizePx,
          expectedReferenceSizePx: s.expectedReferenceSizePx,
          pixelMismatch: s.pixelMismatch,
          diffBoxPt: s.diffBoxPt,
          referenceJpeg: s.referenceJpeg,
          fontsReference: s.fontsReference,
          fontsSubstituted: s.fontsSubstituted,
          unmatchedTexts: s.unmatchedTexts,
          pdfUnmatchedTexts: s.pdfUnmatchedTexts,
          lineMatch: s.lineMatch,
          linesMatched: s.linesMatched,
          linesTotal: s.linesTotal,
          pdfLinesUnmatched: s.pdfLinesUnmatched,
          baselineDelta: s.baselineDelta,
          estimatedBaselines: s.estimatedBaselines,
          skippedLines: s.skippedLines,
          selfConsistent: s.selfConsistent,
          cause: s.cause,
          approximations: approximationTitles(s.diagnostics),
          diagnostics: otherDiagnostics(s.diagnostics),
        })),
      })),
      null,
      2
    )
  );
  writeFileSync(
    join(out, 'gate.json'),
    `${JSON.stringify({ tools: { poppler: context.poppler, chromium: context.chromium }, gated: Boolean(context.gate?.list), exitCode: context.gate ? gateExitCode(context.gate.problems.length, Boolean(context.gate.list)) : null, gate: context.gate ?? null }, null, 2)}\n`
  );
}
