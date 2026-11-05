// List every deck and fixture whose reference PDF is missing or stale. A stale reference is a
// gap, never a pass: the deck or the PDF changed after its identity was recorded, or the PDF has
// no identity in `references.json` (`reference-identity.ts`).
//
//   bun run fidelity:missing             # report; exit 0
//   bun run fidelity:missing -- --check  # exit 1 when a fixture reference is missing or stale
//   bun run fidelity:missing -- --names  # local file names instead of deckN (never commit them)
//   bun run fidelity:missing -- --help
//
// Corpus decks are listed as deck1…deckN.

import { existsSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { parseArgs } from 'node:util';
import { readManifest, referenceState, type ReferenceState } from './reference-identity.ts';

export interface DeckReference {
  readonly file: string;
  readonly label: string;
  readonly state: ReferenceState;
}

/** Every deck in `dir` with the state of its reference, sorted by file name. */
export function deckReferences(dir: string): DeckReference[] {
  if (!existsSync(dir)) return [];
  const manifest = readManifest(dir);
  return readdirSync(dir)
    .filter((f) => f.toLowerCase().endsWith('.pptx'))
    .sort()
    .map((f, i) => {
      const file = join(dir, f);
      return {
        file,
        label: dir === 'demo-presentations' ? `deck${i + 1}` : f,
        state: referenceState(file, manifest),
      };
    });
}

/** The decks whose reference is missing or stale; with `all`, every deck. */
export function missingOrStale(dir: string, options: { all?: boolean } = {}): DeckReference[] {
  return deckReferences(dir).filter((d) => options.all || d.state.status !== 'current');
}

export const MISSING_USAGE = `Usage: bun run fidelity:missing [-- --check] [--names]
  List the decks whose reference PDF is missing or stale, in fixtures/ and demo-presentations/.
  --check   exit 1 when a fixture reference is missing or stale
  --names   print local file names instead of deckN (never commit them)

Exit codes: 0 done, 1 a fixture reference is missing or stale (with --check), 2 a usage error.`;

/** The report lines and the exit code. */
export function missingReport(
  dirs: readonly string[],
  options: {
    readonly check: boolean;
    readonly names: boolean;
    /** The folder whose gaps `check` counts (default `fixtures`). */
    readonly gated?: string;
  }
): { readonly lines: string[]; readonly exitCode: number } {
  const lines: string[] = [];
  let gaps = 0;
  for (const dir of dirs) {
    if (!existsSync(dir)) continue;
    const all = deckReferences(dir);
    const missing = all.filter((d) => d.state.status === 'missing');
    const stale = all.filter((d) => d.state.status === 'stale');
    if (dir === (options.gated ?? 'fixtures')) gaps += missing.length + stale.length;
    const name = (d: DeckReference): string => (options.names ? d.file : d.label);
    lines.push(`${dir}: ${missing.length} without a reference PDF, ${stale.length} stale`);
    for (const d of missing) lines.push(`  missing  ${name(d)}`);
    for (const d of stale)
      if (d.state.status === 'stale')
        lines.push(`  stale    ${name(d)}: ${d.state.reasons.join('; ')}`);
    if (missing.length + stale.length)
      lines.push(
        `  Next: add the reference PDF of each deck next to it, and its entry in ${dir}/references.json`
      );
  }
  return { lines, exitCode: options.check && gaps ? 1 : 0 };
}

if (import.meta.main) {
  let values: { check?: boolean; names?: boolean; help?: boolean };
  try {
    ({ values } = parseArgs({
      strict: true,
      allowPositionals: false,
      options: {
        check: { type: 'boolean' },
        names: { type: 'boolean' },
        help: { type: 'boolean', short: 'h' },
      },
    }));
  } catch (e) {
    console.error(`fidelity:missing: ${(e as Error).message}\n\n${MISSING_USAGE}`);
    process.exit(2);
  }
  if (values.help) {
    console.log(MISSING_USAGE);
    process.exit(0);
  }
  const report = missingReport(['fixtures', 'demo-presentations'], {
    check: values.check ?? false,
    names: values.names ?? false,
  });
  for (const line of report.lines) console.log(line);
  process.exit(report.exitCode);
}
