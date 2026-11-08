// Run DOCX Editor's tests for the shared modules against `@eigenpal/ooxml-core`
// (parity-baseline task 3.3), and report how far each module drifts from its DOCX Editor source.
//
// The DOCX Editor tests import their modules by relative path (`../package/zip.ts`). This script
// copies them into `.cache/docx-module-tests/`, in the same folder layout, next to one-line shims
// that re-export the `ooxml-core` module of the same name. Then it runs them with `bun test`, so
// every import of a shared module, and every `Bun.build` of one, reaches `ooxml-core`. The drift
// report formats both sides with this repository's Prettier first, because the two repositories
// wrap a few lines differently.
//
// Usage (local only; it needs a DOCX Editor checkout):
//   bun run test:docx-shared [--docx <path>] [--ref <git ref>] [--strict]
// `--docx` defaults to $DOCX_EDITOR_DIR, then ~/Developer/docx-editor, ~/docx-editor, and the
// docx-editor folder next to the main PPTX Editor checkout. `--ref` reads the tests, the sources, and the e2e fixtures from that commit
// with git and leaves the checkout alone; without it, the working tree is read. `--strict` also
// fails the run when a module drifts or a listed test is missing at the ref. The run exits with code 2, before any test, on an unknown
// option, an option with no value, a checkout or ref that does not exist, or no test file found.

import { execFileSync, spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, rmSync, symlinkSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { format, resolveConfig } from 'prettier';
import { findDocxCheckout } from './docx-checkout.ts';

const ROOT = resolve(import.meta.dir, '../..');
const CORE = join(ROOT, 'packages/ooxml-core/src');
const MIRROR = join(ROOT, '.cache/docx-module-tests');
const STORE = 'packages/core/src/store';
const LAYOUT = 'packages/core/src/layout';

/** Each shared module: its `ooxml-core` file and its DOCX Editor source. */
const MODULES: readonly (readonly [string, string])[] = [
  ['limits.ts', `${STORE}/runtime/limits.ts`],
  ['opc-names.ts', `${STORE}/package/opc-names.ts`],
  ['zip.ts', `${STORE}/package/zip.ts`],
  ['zip-write.ts', `${STORE}/package/zip-write.ts`],
  ['sinks.ts', `${STORE}/package/sinks.ts`],
  ['sha256.ts', `${STORE}/package/sha256.ts`],
  ['xml-text.ts', `${STORE}/package/xml-text.ts`],
  ['xml-reader.ts', `${STORE}/package/xml-reader.ts`],
  ['xml-fast-read.ts', `${STORE}/package/xml-fast-read.ts`],
  ['grapheme.ts', `${LAYOUT}/grapheme.ts`],
  ['word-segment.ts', `${LAYOUT}/word-segment.ts`],
];

/** DOCX Editor's tests of these modules. A test missing at the ref is skipped, or fails `--strict`. */
const TESTS: readonly string[] = [
  `${STORE}/__tests__/opc-names.test.ts`,
  `${STORE}/__tests__/sinks.test.ts`,
  `${STORE}/__tests__/xml-text-validation.test.ts`,
  `${STORE}/__tests__/sha256.test.ts`,
  `${STORE}/__tests__/sha256-memory.test.ts`,
  `${STORE}/__tests__/xml-reader.test.ts`,
  `${STORE}/__tests__/xml-fast-read.test.ts`,
  `${STORE}/__tests__/zip-proto.test.ts`,
  `${STORE}/package/__tests__/zip-determinism.test.ts`,
  `${STORE}/package/__tests__/zip-write.test.ts`,
  `${LAYOUT}/__tests__/grapheme.test.ts`,
  `${LAYOUT}/__tests__/word-segment.test.ts`,
];

/** The DOCX Editor lines `ooxml-core` leaves out on purpose: WordprocessingML names. */
const DOCX_ONLY = `const MAIN_DOCUMENT_PART_RESULT = normalizePartName('word/document.xml');

/** Canonical OPC part name for the main document story (\`/word/document.xml\`). */
export const WML_MAIN_DOCUMENT_PART = MAIN_DOCUMENT_PART_RESULT.ok
  ? MAIN_DOCUMENT_PART_RESULT.partName
  : '/word/document.xml';

`;

const HEADER =
  /^\/\/ ((?:Copied|Adapted) from docx-editor\/|It also holds fixes that DOCX Editor|Keep it equal to that file)/;

/**
 * Deliberate differences, as [DOCX Editor text, ooxml-core text]. `@pptx-editor.dev/core/store`
 * publishes `escapeXmlAttribute`, so its release tag stays public here.
 */
const KNOWN: readonly (readonly [string, string])[] = [
  [
    '/** XML-escape a value bound for an ATTRIBUTE, preserving CR/LF/TAB across a re-parse. @internal */',
    '/** XML-escape a value bound for an ATTRIBUTE, preserving CR/LF/TAB across a re-parse. */',
  ],
];

const OPTIONS_WITH_VALUE = new Set(['--docx', '--ref']);
const FLAGS = new Set(['--strict']);

// Refuse anything this script does not know, so a typo never runs at the wrong ref.
for (let at = 2; at < process.argv.length; at += 1) {
  const arg = process.argv[at]!;
  if (OPTIONS_WITH_VALUE.has(arg)) at += 1;
  else if (!FLAGS.has(arg)) {
    console.error(`Unknown argument: ${arg}. Use --docx <path>, --ref <ref>, and --strict.`);
    process.exit(2);
  }
}

function docxCheckout(): string {
  const found = findDocxCheckout(option('--docx'));
  if (found.kind === 'found') return found.dir;
  console.error(
    found.kind === 'invalid'
      ? found.message
      : 'No DOCX Editor checkout found. Pass --docx <path> or set DOCX_EDITOR_DIR.'
  );
  process.exit(2);
}

function option(name: string): string | undefined {
  const at = process.argv.indexOf(name);
  if (at < 0) return undefined;
  const value = process.argv[at + 1];
  if (value === undefined || value.startsWith('--')) {
    console.error(`${name} needs a value.`);
    process.exit(2);
  }
  return value;
}

const docx = docxCheckout();
const ref = option('--ref');
const strict = process.argv.includes('--strict');
if (ref !== undefined) {
  const known = spawnSync('git', [
    '-C',
    docx,
    'rev-parse',
    '--verify',
    '--quiet',
    `${ref}^{commit}`,
  ]);
  if (known.status !== 0) {
    console.error(`No commit ${ref} in ${docx}.`);
    process.exit(2);
  }
}

/** A file of the DOCX Editor checkout, at `--ref` when given, or null when it does not exist. */
function docxFile(path: string): string | null {
  if (!ref) return existsSync(join(docx, path)) ? readFileSync(join(docx, path), 'utf8') : null;
  try {
    return execFileSync('git', ['-C', docx, 'show', `${ref}:${path}`], {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'ignore'],
      maxBuffer: 64 * 1024 * 1024,
    });
  } catch {
    return null;
  }
}

/** A module without the source header that names its DOCX Editor file and commit. */
function withoutHeader(text: string): string {
  const lines = text.split('\n');
  let start = 0;
  while (start < lines.length && HEADER.test(lines[start]!)) start += 1;
  return lines.slice(start).join('\n');
}

function write(path: string, text: string): void {
  mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, text);
}

/** `text` as this repository's Prettier prints it, so line wrapping is never drift. */
async function formatted(text: string, file: string): Promise<string> {
  return format(text, { ...(await resolveConfig(file)), filepath: file });
}

// The drift report: each module without its header, against its DOCX Editor source.
const where = ref ? `${docx} at ${ref}` : docx;
console.log(`DOCX Editor: ${where}\n\nDrift (ooxml-core body against the DOCX Editor source):`);
let drifted = 0;
for (const [name, source] of MODULES) {
  const theirs = docxFile(source);
  const ours = withoutHeader(readFileSync(join(CORE, name), 'utf8')).replace(
    "from './limits.ts';",
    "from '../runtime/limits.ts';"
  );
  if (theirs === null) {
    console.log(`  ${name.padEnd(18)} no source at this ref`);
    drifted += 1;
    continue;
  }
  const file = join(CORE, name);
  const known = KNOWN.reduce((text, [docx, here]) => text.replace(docx, here), theirs);
  const expected = await formatted(known.replace(DOCX_ONLY, ''), file);
  const actual = await formatted(ours, file);
  if (expected === actual) console.log(`  ${name.padEnd(18)} same`);
  else {
    const a = new Set(expected.split('\n'));
    const b = new Set(actual.split('\n'));
    const only = (x: Set<string>, y: Set<string>): number => [...x].filter((l) => !y.has(l)).length;
    console.log(
      `  ${name.padEnd(18)} differs: ${only(b, a)} lines only here, ${only(a, b)} only in DOCX`
    );
    console.log(
      ref === undefined
        ? `    see: diff ${join(docx, source)} ${join(CORE, name)}`
        : `    see: git -C ${docx} show ${ref}:${source} | diff - ${join(CORE, name)}`
    );
    drifted += 1;
  }
}

// The mirror: DOCX Editor's tests next to shims that re-export ooxml-core.
rmSync(MIRROR, { recursive: true, force: true });
for (const [name, source] of MODULES)
  write(join(MIRROR, source), `export * from ${JSON.stringify(join(CORE, name))};\n`);
const tests: string[] = [];
const missing: string[] = [];
for (const path of TESTS) {
  const text = docxFile(path);
  if (text === null) missing.push(path);
  else {
    write(join(MIRROR, path), text);
    tests.push(join(MIRROR, path));
  }
}
// `xml-fast-read.test.ts` reads DOCX Editor's e2e fixtures relative to itself: from the ref when
// one is given, else from the working tree.
mkdirSync(join(MIRROR, 'e2e'), { recursive: true });
if (ref !== undefined) {
  const archive = spawnSync('git', ['-C', docx, 'archive', ref, 'e2e/fixtures'], {
    maxBuffer: 512 * 1024 * 1024,
  });
  if (archive.status === 0)
    spawnSync('tar', ['-x', '-C', MIRROR], {
      input: archive.stdout,
      stdio: ['pipe', 'ignore', 'inherit'],
    });
} else if (existsSync(join(docx, 'e2e/fixtures'))) {
  symlinkSync(join(docx, 'e2e/fixtures'), join(MIRROR, 'e2e/fixtures'));
}

if (missing.length) console.log(`\nNot at this ref (skipped):\n  ${missing.join('\n  ')}`);
// With no file arguments, `bun test` would run every test in this repository instead.
if (tests.length === 0) {
  console.error('\nNo DOCX Editor test file found; nothing ran.');
  process.exit(2);
}
console.log(`\nRunning ${tests.length} DOCX Editor test files against @eigenpal/ooxml-core\n`);
const run = spawnSync('bun', ['test', '--timeout', '60000', ...tests], {
  cwd: ROOT,
  stdio: 'inherit',
});
console.log(`\n${drifted} of ${MODULES.length} modules differ from the DOCX Editor source.`);
process.exit(
  run.status !== 0 ? (run.status ?? 1) : strict && (drifted > 0 || missing.length > 0) ? 1 : 0
);
