// Line model calibration: lays out every slide with the engine and the reference
// fonts, pairs each laid-out line with the PDF text run that starts it, and counts how many
// baselines are within 0.25, 0.5, 1, and 2 pt of the reference PDF's glyph origin. Layout only: it
// does not paint, so it is a measurement of `layoutSlide`, not a second renderer.
//
// Reports COUNTS ONLY. It never prints or writes deck text or file names.
//
//   bun run fidelity:calibrate                     # the fixtures and their reference PDFs
//   PPTX_EVAL_CORPUS=<dir> bun run fidelity:calibrate --smoke [--generation <name>]
//                                                  # the smoke set of eigenpal/pptx-eval (local)

import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { openPresentation } from '../../packages/core/src/editor/index.ts';
import { WHOLE_POINT_BASELINES, type LayoutItem } from '../../packages/core/src/layout/index.ts';
import { deckFontFamilies, localFonts, referenceFontsInstalled } from './local-fonts.ts';
import { deckFacts, slidePages } from './reference-identity.ts';
import { glyphOrigins, normalize, type GlyphOrigin } from './references.ts';

const bare = (s: string): string => s.replace(/\s+/g, '');

const EMU_PER_PT = 12_700;

interface Deck {
  readonly pptx: string;
  readonly pdf: string;
}

/** One paired line: our baseline minus the reference baseline, and the groups it counts in. */
interface Pair {
  readonly delta: number;
  readonly groups: readonly string[];
}

/** A PDF text run, and whether a laid-out line already took it. */
type Origin = GlyphOrigin & { used: boolean };

/**
 * The pairing window: a PDF run pairs with a laid-out line that starts with the same text and
 * lies within this distance. A line outside it counts as unpaired, not as a miss.
 */
const WINDOW = { dyPt: 15, dxPt: 40 } as const;

function decks(args: readonly string[]): Deck[] {
  if (!args.includes('--smoke')) {
    if (args.includes('--generation')) usage('--generation applies to --smoke only.');
    const refs = JSON.parse(readFileSync('fixtures/references.json', 'utf8')) as {
      references: Record<string, unknown>;
    };
    return Object.keys(refs.references).map((n) => ({
      pptx: `fixtures/${n}.pptx`,
      pdf: `fixtures/${n}.pdf`,
    }));
  }
  const root = process.env.PPTX_EVAL_CORPUS;
  if (!root) usage('Set PPTX_EVAL_CORPUS to the eigenpal/pptx-eval corpus folder.');
  if (!existsSync(join(root, 'smoke.json')))
    usage('PPTX_EVAL_CORPUS has no smoke.json: name the eigenpal/pptx-eval corpus folder.');
  const at = args.indexOf('--generation');
  const generation = at >= 0 ? args[at + 1] : process.env.PPTX_EVAL_GENERATION;
  if (!generation || generation.startsWith('--'))
    usage('--generation needs a reference generation, or set PPTX_EVAL_GENERATION.');
  if (!existsSync(join(root, 'references', generation)))
    usage(`The corpus has no references for the generation ${generation}.`);
  const smoke = JSON.parse(readFileSync(join(root, 'smoke.json'), 'utf8')) as {
    sources: string[];
  };
  return smoke.sources
    .map((sha) => ({
      pptx: join(root, 'objects', 'sha256', `${sha}.pptx`),
      pdf: join(root, 'references', generation, `${sha}.pdf`),
    }))
    .filter((d) => existsSync(d.pptx) && existsSync(d.pdf));
}

function usage(message: string): never {
  console.error(message);
  process.exit(2);
}

/**
 * Pair the unrotated shape lines of one slide with PDF runs, nearest by baseline and left edge.
 * Returns the number of laid-out lines that were candidates.
 */
function pairLines(items: readonly LayoutItem[], runs: Origin[], out: Pair[]): number {
  let candidates = 0;
  for (const it of items) {
    if (it.kind === 'group') candidates += pairLines(it.children, runs, out);
    if (it.kind !== 'shape' || !it.text || !it.textMatrix || !it.element.text) continue;
    const m = it.textMatrix;
    if (Math.abs(m[1]) > 1e-6 || Math.abs(m[2]) > 1e-6 || m[3] < 0 || it.text.rotation % 360)
      continue;
    const body = it.element.text;
    let paragraph = -1;
    let first = true;
    for (const line of it.text.lines) {
      const text = normalize(
        line.runs
          .filter((r) => r.kind === 'text')
          .map((r) => r.text)
          .join('')
      );
      const lineOfParagraph = line.paragraph === paragraph ? 'later line' : 'first line';
      paragraph = line.paragraph;
      if (!text) continue;
      const lead = line.runs.find((r) => r.kind === 'text');
      if (!lead) continue;
      candidates += 1;
      const x = (m[0] * (it.text.box.x + lead.x[0]) + m[4]) / EMU_PER_PT;
      const y = (m[3] * (it.text.box.y + line.baseline) + m[5]) / EMU_PER_PT;
      // Compare without whitespace: `pdftotext` and pdf.js add spaces inside letter-spaced text.
      // A short run (a raised digit, a single capital) pairs when it starts the line's text.
      const lineKey = bare(text);
      let best: Origin | undefined;
      let score = Infinity;
      for (const r of runs) {
        const runKey = bare(r.text);
        if (r.used || !runKey) continue;
        const head = Math.min(8, lineKey.length, runKey.length);
        if (runKey.slice(0, head) !== lineKey.slice(0, head)) continue;
        const dy = Math.abs(r.y - y);
        const dx = Math.abs(r.x - x);
        if (dy > WINDOW.dyPt || dx > WINDOW.dxPt) continue;
        if (dy + dx * 0.05 < score) {
          score = dy + dx * 0.05;
          best = r;
        }
      }
      const spacing = body.paragraphs[line.paragraph].lineSpacing;
      if (best) {
        best.used = true;
        const autofit = body.body.autofit.kind;
        out.push({
          delta: y - best.y,
          groups: [
            'pts' in spacing ? 'exact spacing' : `${Math.round(spacing.pct * 100)} % spacing`,
            `autofit ${autofit}`,
            `anchor ${body.body.anchor}`,
            first ? 'first line of the body' : lineOfParagraph,
          ],
        });
      }
      first = false;
    }
  }
  return candidates;
}

async function main(): Promise<void> {
  const args = process.argv.slice(2);
  const pairs: Pair[] = [];
  let candidates = 0;
  let scored = 0;
  let unreadable = 0;
  const list = decks(args);
  for (const deck of list) {
    const bytes = new Uint8Array(readFileSync(deck.pptx));
    const fonts = localFonts(deckFontFamilies(bytes)).map((f) => ({
      family: f.family,
      bold: f.bold,
      italic: f.italic,
      bytes: new Uint8Array(readFileSync(f.file)),
      origin: 'local',
    }));
    const opened = await openPresentation(bytes, { fonts });
    if (!opened.ok) {
      unreadable += 1;
      continue;
    }
    const pdf = await glyphOrigins(deck.pdf).catch(() => undefined);
    if (!pdf) {
      unreadable += 1;
      opened.session.dispose();
      continue;
    }
    const pages = slidePages(deckFacts(bytes));
    for (let i = 0; i < opened.session.info.slides.length; i += 1) {
      const page = pages[i];
      const layout = opened.session.layout(i);
      if (page === undefined || !layout || !pdf[page]) continue;
      const runs = pdf[page].map((o) => ({ ...o, used: false }));
      candidates += pairLines(layout.items, runs, pairs);
    }
    opened.session.dispose();
    scored += 1;
  }
  const within = (t: number): number => pairs.filter((p) => Math.abs(p.delta) <= t).length;
  const fontsNote = referenceFontsInstalled() ? 'installed' : 'NOT installed';
  const rounding = WHOLE_POINT_BASELINES ? 'on' : 'off';
  console.log(
    `Decks ${scored} of ${list.length}; reference fonts ${fontsNote}; ` +
      `whole-point baselines ${rounding} (WHOLE_POINT_BASELINES).`
  );
  if (unreadable)
    console.warn(
      `${unreadable} decks skipped: the deck did not open or pdf.js could not read its PDF.`
    );
  console.log(
    `Laid-out lines ${candidates}; paired ${pairs.length}; unpaired ${candidates - pairs.length} ` +
      `(no PDF run with the same start within ${WINDOW.dyPt} pt down and ${WINDOW.dxPt} pt across).`
  );
  console.log(
    `Paired within 0.25 pt: ${within(0.25)}; 0.5 pt: ${within(0.5)}; 1 pt: ${within(1)}; 2 pt: ${within(2)}.`
  );
  const groups = new Map<string, number[]>();
  for (const p of pairs)
    for (const g of p.groups) {
      const values = groups.get(g) ?? [];
      values.push(Math.abs(p.delta));
      groups.set(g, values);
    }
  for (const [g, d] of [...groups].sort((a, b) => b[1].length - a[1].length).slice(0, 20)) {
    const ok = d.filter((v) => v <= 0.5).length;
    console.log(`  ${g}: ${ok} of ${d.length} within 0.5 pt`);
  }
}

await main();
