// Reference identity: the deck and PDF bytes of each reference, recorded in `references.json`
// next to the decks (`fixtures/references.json` is committed). A reference is current only while the deck
// and the PDF are the bytes that were recorded. Any other state is a gap, never a pass:
//
// - missing: no PDF next to the deck;
// - stale: the deck or the PDF changed since they were recorded, or the PDF has no recorded identity.
//
// Content hashes, not file times, decide staleness: git does not keep file times.

import { createHash } from 'node:crypto';
import { existsSync, readFileSync, renameSync, writeFileSync } from 'node:fs';
import { basename, dirname, join } from 'node:path';
import { EMU_PER_POINT, presentationInfo, readPptx } from '../../packages/core/src/store/index.ts';
import { PAGE_SIZE_TOLERANCE_PT } from './thresholds.ts';

export const MANIFEST = 'references.json';

export interface ReferenceIdentity {
  /** SHA-256 of the deck bytes. */
  readonly deck: string;
  /** SHA-256 of the PDF bytes. */
  readonly pdf: string;
  readonly pages: number;
  /** Page 1 size in points. */
  readonly pageSizePt: readonly [number, number];
}

export interface ReferenceManifest {
  readonly version: 1;
  /** By deck base name (`text-runs` for `text-runs.pptx`). */
  readonly references: Readonly<Record<string, ReferenceIdentity>>;
}

export type ReferenceState =
  | { readonly status: 'current'; readonly pdf: string; readonly identity: ReferenceIdentity }
  | { readonly status: 'missing' }
  | { readonly status: 'stale'; readonly pdf: string; readonly reasons: readonly string[] };

export const sha256 = (bytes: Uint8Array): string =>
  createHash('sha256').update(bytes).digest('hex');

export const sha256File = (path: string): string => sha256(readFileSync(path));

export const deckKey = (deck: string): string => basename(deck).replace(/\.pptx$/i, '');

/** The PDF path next to a deck, whether or not it exists. */
export const pdfPathFor = (deck: string): string => join(dirname(deck), `${deckKey(deck)}.pdf`);

export function readManifest(dir: string): ReferenceManifest {
  const file = join(dir, MANIFEST);
  if (!existsSync(file)) return { version: 1, references: {} };
  const parsed = JSON.parse(readFileSync(file, 'utf8')) as ReferenceManifest;
  if (parsed.version !== 1) throw new Error(`${file}: unknown version ${String(parsed.version)}`);
  return parsed;
}

/**
 * Write the manifest with sorted keys, so a new reference changes only its own entry. It is written
 * to a temporary file and renamed, so a crash never leaves half a file.
 */
export function writeManifest(dir: string, manifest: ReferenceManifest): void {
  const references: Record<string, ReferenceIdentity> = Object.create(null);
  for (const key of Object.keys(manifest.references).sort())
    references[key] = manifest.references[key];
  const file = join(dir, MANIFEST);
  const tmp = `${file}.${process.pid}.tmp`;
  writeFileSync(tmp, `${JSON.stringify({ version: 1, references }, null, 2)}\n`);
  renameSync(tmp, file);
}

/** Whether a deck's PDF exists and matches the identity recorded for it. */
export function referenceState(deck: string, manifest: ReferenceManifest): ReferenceState {
  const pdf = pdfPathFor(deck);
  if (!existsSync(pdf)) return { status: 'missing' };
  const identity = Object.hasOwn(manifest.references, deckKey(deck))
    ? manifest.references[deckKey(deck)]
    : undefined;
  if (!identity) return { status: 'stale', pdf, reasons: ['no identity recorded for the PDF'] };
  const reasons: string[] = [];
  if (sha256File(deck) !== identity.deck)
    reasons.push('the deck changed after its reference was recorded');
  if (sha256File(pdf) !== identity.pdf) reasons.push('the PDF changed after it was recorded');
  return reasons.length ? { status: 'stale', pdf, reasons } : { status: 'current', pdf, identity };
}

// ---------------------------------------------------------------- deck and PDF facts

export interface DeckFacts {
  readonly slides: number;
  /** 1-based indexes of hidden slides (`p:sld/@show="0"`). Reference PDFs leave them out. */
  readonly hidden: readonly number[];
  readonly widthPt: number;
  readonly heightPt: number;
  /** `p:sldSz` in EMU. */
  readonly widthEmu: number;
  readonly heightEmu: number;
}

/** Slide count, hidden slides, and slide size, read with the engine's bounded package reader. */
export function deckFacts(bytes: Uint8Array): DeckFacts {
  const result = readPptx(bytes);
  if (!result.ok) throw new Error(`the deck does not open: ${result.error.reason}`);
  const info = presentationInfo(result.pkg);
  return {
    slides: info.slides.length,
    hidden: info.slides.flatMap((s, i) => (s.hidden ? [i + 1] : [])),
    widthPt: info.width / EMU_PER_POINT,
    heightPt: info.height / EMU_PER_POINT,
    widthEmu: info.width,
    heightEmu: info.height,
  };
}

/**
 * For each slide (0-based), its PDF page (0-based), or undefined for a hidden slide. A
 * reference PDF holds the visible slides only, in order (hidden slides have no page).
 */
export function slidePages(facts: DeckFacts): (number | undefined)[] {
  const hidden = new Set(facts.hidden);
  let page = 0;
  return Array.from({ length: facts.slides }, (_, i) => (hidden.has(i + 1) ? undefined : page++));
}

export interface PdfFacts {
  /** Size of each page in points. */
  readonly pages: readonly (readonly [number, number])[];
  /** The PDF's `Creator` and `Producer` (empty when absent). */
  readonly producer?: string;
}

/** Tools whose PDFs are never references (AGENTS.md, "Fidelity rules"). */
const NOT_REFERENCE = /keynote|libreoffice|openoffice|google|pages/i;

/**
 * The problems that make a PDF unusable as the reference of a deck: the page count must equal
 * the visible slide count, every page must be the slide size within `PAGE_SIZE_TOLERANCE_PT`,
 * and no tool in `NOT_REFERENCE` may be its creator or producer.
 */
export function referenceProblems(deck: DeckFacts, pdf: PdfFacts): string[] {
  const visible = deck.slides - deck.hidden.length;
  const out: string[] = [];
  if (pdf.producer && NOT_REFERENCE.test(pdf.producer))
    out.push(`the PDF was made by ${pdf.producer}, which never makes a reference PDF`);
  if (pdf.pages.length !== visible)
    out.push(`the PDF has ${pdf.pages.length} pages, the deck has ${visible} visible slides`);
  const wrong = pdf.pages.findIndex(
    ([w, h]) =>
      // A tiny epsilon: an EMU size in points is not always exact in floating point.
      Math.abs(w - deck.widthPt) > PAGE_SIZE_TOLERANCE_PT + 1e-6 ||
      Math.abs(h - deck.heightPt) > PAGE_SIZE_TOLERANCE_PT + 1e-6
  );
  if (wrong >= 0) {
    const [w, h] = pdf.pages[wrong];
    out.push(
      `page ${wrong + 1} is ${w} × ${h} pt, the slide is ${deck.widthPt} × ${deck.heightPt} pt`
    );
  }
  return out;
}
