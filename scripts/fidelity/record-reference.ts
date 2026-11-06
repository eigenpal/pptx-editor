// Record the reference of each named fixture in `fixtures/references.json`: the SHA-256 of the
// deck and of its PDF, the page count, and the size of page 1. Run it after you export a
// reference PDF. The fidelity check then treats the PDF as current until either file changes.
// It refuses a deck that changed with its old PDF (export the PDF again), and a PDF that fails
// the reference checks: page count and size, embedded fonts, and the deck's open fonts.
//
//   bun run fidelity:record-reference tables [more fixture names...]

import { existsSync, readFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import {
  deckFontFamilies,
  foreignReferenceFonts,
  harnessFontFiles,
  localFonts,
  postScriptFamilies,
} from './local-fonts.ts';
import { openPdf } from './pdf.ts';
import {
  deckFacts,
  readManifest,
  referenceProblems,
  sha256,
  writeManifest,
  type ReferenceIdentity,
} from './reference-identity.ts';

/**
 * The identity of `<dir>/<name>.pptx` and `<dir>/<name>.pdf`, checked against `recorded` (the
 * identity in `references.json`, if any). Throws with every problem when the pair may not become
 * a reference.
 */
export async function referenceIdentity(
  dir: string,
  name: string,
  recorded?: ReferenceIdentity
): Promise<ReferenceIdentity> {
  const deck = join(dir, `${name}.pptx`);
  const pdf = join(dir, `${name}.pdf`);
  for (const file of [deck, pdf]) if (!existsSync(file)) throw new Error(`${file} does not exist`);
  const deckBytes = new Uint8Array(readFileSync(deck));
  const bytes = new Uint8Array(readFileSync(pdf));
  const identity = { deck: sha256(deckBytes), pdf: sha256(bytes) };
  const problems: string[] = [];
  if (recorded && recorded.deck !== identity.deck && recorded.pdf === identity.pdf)
    problems.push('the deck changed, and the PDF did not: export the PDF again');
  const facts = deckFacts(deckBytes);
  const doc = await openPdf(bytes, facts.slides);
  try {
    problems.push(...referenceProblems(facts, doc));
    const fonts = await doc.fonts();
    for (const f of fonts)
      if (!f.embedded) problems.push(`the PDF does not embed the font ${f.name}`);
    const open = postScriptFamilies(
      harnessFontFiles(localFonts(undefined, 'fixtures').map((f) => f.file))
    );
    const foreign = foreignReferenceFonts(fonts, deckFontFamilies(deckBytes, false), open);
    if (foreign.length)
      problems.push(
        `the PDF embeds fonts that are not open fonts of the deck: ${foreign.join(', ')}`
      );
    if (problems.length) throw new Error(`${name}: ${problems.join('; ')}`);
    const [w, h] = doc.pages[0];
    return { ...identity, pages: doc.pages.length, pageSizePt: [w, h] };
  } finally {
    await doc.close();
  }
}

/** Record the references of `names` in `<dir>/references.json`; other entries stay. */
export async function recordReferences(dir: string, names: readonly string[]): Promise<string[]> {
  const manifest = readManifest(dir);
  const references = { ...manifest.references };
  const lines: string[] = [];
  for (const name of names) {
    const was = Object.hasOwn(references, name) ? references[name] : undefined;
    const identity = await referenceIdentity(dir, name, was);
    references[name] = identity;
    lines.push(
      `${name}: ${identity.pages} pages of ${identity.pageSizePt[0]} × ${identity.pageSizePt[1]} pt`
    );
  }
  writeManifest(dir, { version: 1, references });
  return lines;
}

if (import.meta.main) {
  const names = process.argv.slice(2);
  if (!names.length || names.some((n) => n.startsWith('-'))) {
    console.error('Usage: bun run fidelity:record-reference <fixture> [<fixture>...]');
    process.exit(2);
  }
  try {
    for (const line of await recordReferences(resolve('fixtures'), names)) console.log(line);
  } catch (e) {
    console.error(`fidelity: ${(e as Error).message}`);
    process.exit(1);
  }
}
