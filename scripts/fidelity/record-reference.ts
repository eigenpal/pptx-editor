// Record the reference of each named fixture in `fixtures/references.json`: the SHA-256 of the
// deck and of its PDF, the page count, and the size of page 1. Run it after you export a
// reference PDF. The fidelity check then treats the PDF as current until either file changes.
//
//   bun run fidelity:record-reference tables [more fixture names...]

import { existsSync, readFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { openPdf } from './pdf.ts';
import {
  readManifest,
  sha256,
  writeManifest,
  type ReferenceIdentity,
} from './reference-identity.ts';

/** The identity of `<dir>/<name>.pptx` and `<dir>/<name>.pdf`. Throws when either is missing. */
export async function referenceIdentity(dir: string, name: string): Promise<ReferenceIdentity> {
  const deck = join(dir, `${name}.pptx`);
  const pdf = join(dir, `${name}.pdf`);
  for (const file of [deck, pdf]) if (!existsSync(file)) throw new Error(`${file} does not exist`);
  const bytes = new Uint8Array(readFileSync(pdf));
  const doc = await openPdf(bytes);
  try {
    if (!doc.pages.length) throw new Error(`${pdf} has no pages`);
    const [w, h] = doc.pages[0];
    return {
      deck: sha256(new Uint8Array(readFileSync(deck))),
      pdf: sha256(bytes),
      pages: doc.pages.length,
      pageSizePt: [w, h],
    };
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
    const identity = await referenceIdentity(dir, name);
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
