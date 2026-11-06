// The fidelity HTML report: per slide, our PDF page, the reference page, and the diff side by
// side, with the scores, the verdict, and the probable cause. Written to fidelity-report/
// (gitignored: corpus pages hold private content).

import { writeFileSync } from 'node:fs';
import { join, relative } from 'node:path';
import { APPROXIMATIONS } from '../../packages/core/src/layout/approximations.ts';
import { THRESHOLDS } from './thresholds.ts';

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

/** The causes of a slide's failed line scores, from its measurements alone. */
export function lineCauses(r: SlideResult): string[] {
  const causes: string[] = [];
  if (r.lineMatch !== undefined && r.lineMatch < THRESHOLDS.minLineBreakMatch) {
    const chars = (texts: readonly string[] | undefined): string =>
      [...(texts ?? []).join('').replace(/\s+/g, '')].sort().join('');
    causes.push(
      !r.extraLines?.length
        ? 'our page lacks text of the reference page'
        : chars(r.missingLines) === chars(r.extraLines)
          ? 'line breaks differ'
          : 'line breaks or text differ'
    );
  }
  if (r.baselineDelta !== undefined && r.baselineDelta > THRESHOLDS.maxBaselineDeltaPt)
    causes.push('line height or baseline model');
  if ((r.extraLines?.length ?? 0) > THRESHOLDS.maxExtraLines && !causes.length)
    causes.push('our page has text that the reference page lacks');
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
  const effects = tagged(diagnostics, 'effect not rendered: ');
  if (effects.length) causes.push(`effect not rendered (${effects.join(', ')})`);
  const lines = lineCauses(r);
  causes.push(...lines);
  // Every text body uses the line model. It moves baselines only, so it is a cause only when a
  // line check fails.
  const approximations = approximationTitles(
    lines.length
      ? diagnostics
      : diagnostics.filter((d) => d !== `${APPROXIMATION_PREFIX}line-model`)
  );
  if (approximations.length) causes.push(`approximations: ${approximations.join(', ')}`);
  if (causes.length === 0 && (r.pixelMismatch ?? 0) > THRESHOLDS.maxPixelMismatch)
    causes.push('pixels differ, and no layout diagnostic names a cause: see the diff');
  return causes.join('; ');
}

export interface SlideResult {
  index: number;
  /** A hidden slide has no page in either PDF. */
  hidden?: boolean;
  /** Its page in both PDFs (1-based). */
  page?: number;
  /** PNG files of our page, the reference page, and the diff, when the run keeps images. */
  ours?: string;
  reference?: string;
  diff?: string;
  /** `stale`: the reference no longer matches the deck. Like `no-reference`, a gap. */
  status: 'pass' | 'fail' | 'no-reference' | 'stale';
  /** A failing fixture slide listed in `FIXTURE_KNOWN_GAPS`. It is still a failure. */
  knownGap?: boolean;
  /** The scores a failing slide misses, and its known gap (`slideVerdict` in known-gaps.ts). */
  verdict?: string;
  pixelMismatch?: number;
  /** The share of reference lines that our page holds; undefined without reference lines. */
  lineMatch?: number;
  linesMatched?: number;
  linesTotal?: number;
  /** Worst baseline distance of a matched line; undefined when no line matched. */
  baselineDelta?: number;
  /** Reference lines that our page does not hold, and our lines that the reference lacks. */
  missingLines?: readonly string[];
  extraLines?: readonly string[];
  /** The box around every differing pixel, `[x0, y0, x1, y1]` in points from the top-left. */
  diffBoxPt?: readonly [number, number, number, number];
  /** Families that the editor substituted. */
  fontsSubstituted?: readonly string[];
  /** Save and reopen paint the same markup. */
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
  /** Our PDF and the reference PDF, when the run keeps images. */
  oursPdf?: string;
  referencePdf?: string;
  /** Fonts of our PDF that the harness page did not load, or that it does not embed. */
  foreignFonts?: readonly string[];
  /** Fonts of a fixture's reference PDF that are not open fonts that its deck names. */
  foreignReferenceFonts?: readonly string[];
  slides: SlideResult[];
}

/** What the report says about the run besides the slides. */
export interface ReportContext {
  /** The tools: pdf.js, Chromium, and the fixture fonts. */
  readonly tools: string;
  /** The fixture gate, when this run is gated. */
  readonly gate?: { readonly problems: readonly string[]; readonly next: string };
}

const esc = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const pct = (v: number | undefined): string =>
  v === undefined ? '—' : `${(v * 100).toFixed(2)} %`;

const list = (label: string, items: readonly string[]): string =>
  items.length ? `<p class="diag"><b>${label}:</b> ${items.map(esc).join(' · ')}</p>` : '';

function slideHtml(out: string, s: SlideResult): string {
  const img = (p: string | undefined, label: string, empty: string): string =>
    p
      ? `<figure><a href="${esc(relative(out, p))}"><img loading="lazy" src="${esc(relative(out, p))}" alt="${label}"></a><figcaption>${label}</figcaption></figure>`
      : `<figure class="empty"><div>${empty}</div><figcaption>${label}</figcaption></figure>`;
  const lines = s.linesTotal ? `lines ${s.linesMatched} of ${s.linesTotal}` : 'no lines to check';
  const baseline =
    s.baselineDelta === undefined
      ? ''
      : ` <span>baseline Δ ${s.baselineDelta.toFixed(2)} pt</span>`;
  const badge = s.knownGap ? 'fail (known gap)' : s.status;
  return (
    `<article class="${s.status}"><header><b>Slide ${s.index}</b> <span class="badge ${s.status}">${badge}</span>` +
    ` <span>pixels ${pct(s.pixelMismatch)}</span> <span>${lines}</span>${baseline}` +
    (s.selfConsistent === false ? ' <span class="fail">save/reopen DIFFERS</span>' : '') +
    `</header>` +
    (s.verdict ? `<p class="cause">${esc(s.verdict)}</p>` : '') +
    (s.cause && s.status !== 'pass' ? `<p class="cause">Probable cause: ${esc(s.cause)}</p>` : '') +
    list('Reference lines that our page lacks', (s.missingLines ?? []).slice(0, 8)) +
    list('Our lines that the reference lacks', (s.extraLines ?? []).slice(0, 8)) +
    list('Fonts substituted', s.fontsSubstituted ?? []) +
    list('Approximations', approximationTitles(s.diagnostics)) +
    list('Diagnostics', otherDiagnostics(s.diagnostics)) +
    `<div class="imgs">${img(s.ours, 'Our PDF', 'no page')}${img(s.reference, 'Reference PDF', 'no page')}${img(s.diff, 'Diff', s.pixelMismatch === 0 ? 'no differing pixels' : 'no diff')}</div></article>`
  );
}

export function writeReport(
  out: string,
  set: string,
  decks: readonly DeckResult[],
  context: ReportContext
): void {
  const all = decks.flatMap((d) => d.slides);
  const count = (s: SlideResult['status']): number =>
    all.filter((x) => x.status === s && !x.hidden).length;
  const hidden = all.filter((x) => x.hidden).length;
  const rows = decks
    .map((d) => {
      if (d.error)
        return `<section><h2>${esc(d.name)}</h2><p class="fail">Not scored: ${esc(d.error)}</p></section>`;
      const pdfs = [
        d.oursPdf ? `<a href="${esc(relative(out, d.oursPdf))}">our PDF</a>` : '',
        d.referencePdf ? `<a href="${esc(relative(out, d.referencePdf))}">reference PDF</a>` : '',
      ].filter(Boolean);
      const reference =
        d.reference === 'pdf'
          ? pdfs.join(' · ')
          : `${d.reference === 'stale' ? 'stale reference' : 'no reference'}: ${esc((d.referenceIssues ?? []).join('; '))}`;
      const fonts =
        list('Fonts of our PDF that the harness did not load', d.foreignFonts ?? []) +
        list('Reference fonts that are not open fonts of the deck', d.foreignReferenceFonts ?? []);
      return `<section><h2>${esc(d.name)} <small>${reference}</small></h2>${fonts}${d.slides.map((s) => slideHtml(out, s)).join('')}</section>`;
    })
    .join('');
  const gate = context.gate
    ? `<section class="gate"><h2>Gate: ${context.gate.problems.length ? 'fails' : 'passes'}</h2>${
        context.gate.problems.length
          ? `<ul>${context.gate.problems.map((p) => `<li>${esc(p)}</li>`).join('')}</ul>`
          : ''
      }<p>${esc(context.gate.next)}</p></section>`
    : '';
  const html = `<!doctype html><html lang="en"><head><meta charset="utf-8"><title>Fidelity report: ${esc(set)}</title>
<style>
body{font:14px/1.5 -apple-system,system-ui,sans-serif;margin:24px;color:#1b1a19;background:#f4f3f0}
h1{font-size:20px}h2{font-size:16px;margin-top:32px}small{color:#75716b;font-weight:400}
.summary span{display:inline-block;margin-right:16px}
article,.gate{background:#fff;border-radius:12px;padding:12px 16px;margin:12px 0;box-shadow:0 0 0 1px #e7e4df}
.gate h2{margin-top:0}
article header span{margin-left:12px;color:#44413d}
.badge{border-radius:999px;padding:1px 8px;font-size:12px;font-weight:600}
.badge.pass{background:#e6f4ea;color:#1e7a34}.badge.fail{background:#fbe9e7;color:#a8432a}.badge.no-reference,.badge.stale{background:#eceae6;color:#44413d}
.fail{color:#a8432a}.cause{color:#a8432a;margin:6px 0}.diag{color:#75716b;font-size:12px;margin:4px 0}
.imgs{display:grid;grid-template-columns:repeat(3,1fr);gap:8px}figure{margin:0}img{width:100%;border:1px solid #e7e4df}
figcaption{font-size:12px;color:#75716b}.empty div{aspect-ratio:16/9;display:grid;place-items:center;border:1px dashed #d9d5ce;color:#a29e97}
@media (max-width:800px){.imgs{grid-template-columns:1fr}}
</style></head><body><h1>Fidelity report: ${esc(set)}</h1>
<p class="summary"><span>Slides: ${all.length}</span><span>Pass: ${count('pass')}</span><span>Fail: ${count('fail')} (known gaps: ${all.filter((x) => x.knownGap).length})</span><span>No reference: ${count('no-reference')}</span><span>Stale reference: ${count('stale')}</span><span>Hidden: ${hidden}</span></p>
<p class="diag">Our PDF is printed with the editor's own paint. pdf.js draws and reads both PDFs. Thresholds: pixels ≤ ${pct(THRESHOLDS.maxPixelMismatch)}, every reference line, no extra lines, baseline ≤ ${THRESHOLDS.maxBaselineDeltaPt} pt (<code>scripts/fidelity/thresholds.ts</code>). ${esc(context.tools)}. Rotated, vertical, and right-to-left text is checked by pixels only.</p>
${gate}
${rows}</body></html>`;
  writeFileSync(join(out, 'index.html'), html);
  writeFileSync(
    join(out, 'summary.json'),
    `${JSON.stringify(
      decks.map((d) => ({
        name: d.name,
        error: d.error,
        reference: d.reference,
        referenceIssues: d.referenceIssues,
        foreignFonts: d.foreignFonts,
        foreignReferenceFonts: d.foreignReferenceFonts,
        slides: d.slides.map(({ ours: _o, reference: _r, diff: _d, ...s }) => ({
          ...s,
          approximations: approximationTitles(s.diagnostics),
          diagnostics: otherDiagnostics(s.diagnostics),
        })),
      })),
      null,
      2
    )}\n`
  );
}
