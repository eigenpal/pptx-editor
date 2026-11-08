// Drift guard for the files that this repository takes from DOCX Editor. The record is the
// manifest `docx-copies.json`: entries with a file, its kind (`copied`, `adapted`, or
// `reference`), its DOCX Editor sources, and the DOCX Editor commit (`docx-copies.ts`). The
// checks catch honest drift and honest labeling mistakes, not deliberate evasion. The check
// verifies the manifest:
//
// 1. Each entry has a valid shape, and each manifest file is a tracked file on disk.
// 2. One comment paragraph of each file names the kind, DOCX Editor, each source, and the
//    commit of each entry, each as a whole word (`headerProblems`). No paragraph that mentions
//    DOCX Editor names a foreign file that no entry of the file lists. A file without a
//    `copied` entry puts no copy word near a mention of DOCX Editor, unless it is a Markdown
//    guide whose entry gives a `reason`. Elsewhere a `reason` is a note for reviewers.
// 3. Each tracked file whose bytes are UTF-8 once the NUL bytes go, of any type, with a copy
//    signal has an entry, or is in `MENTIONS_ALLOWED` with a reason (`copySignals`). A signal is a
//    `docx-editor/` path, or a paragraph that mentions DOCX Editor and names a commit, a path
//    that is not in this repository, or a path right after the mention, or has a copy verb in
//    it or a copy word near the mention. A commit or a foreign path in the next two comment
//    paragraphs counts too. A plain name mention needs no entry.
// 4. With a DOCX Editor checkout: each commit exists, and each source exists at its commit.
// 5. With a checkout: each `copied` file equals its source at the commit, after the
//    normalization of `normalizedBody` in `docx-normalize.ts`. Comments other than directives
//    go, layout is reprinted, local imports become module paths from the repository root, and
//    the package scopes `@docx-editor.dev/` and `@pptx-editor.dev/` compare equal. Code that
//    does not parse fails. `RENAMED_MODULES` maps each DOCX Editor module that a copy imports
//    from another path here. Any other difference is drift.
// 6. With a checkout: an `adapted` or `reference` file must not equal any source. An `adapted`
//    file must keep `ADAPTED_FLOOR` of the lines of its sources together and of each source on
//    its own. A `reference` file must share at most `REFERENCE_CEILING` of its lines with each
//    source (`kindProblems`). A folder source counts as each file in it that the check can
//    compare (`folderProblems`). An `adapted` source that the check cannot compare fails.
//
// Usage:
//   bun run check:docx-copies [--docx <path>] [--ref <commit>] [--write] [--diff] [--report]
//     [--only <path>]
// `--docx` defaults to $DOCX_EDITOR_DIR, then ~/Developer/docx-editor, ~/docx-editor, and the
// docx-editor folder next to the main checkout. `--ref` compares each copy with that commit
// instead of the one its entry names, and `--write` then puts that commit in the entry and the
// header paragraph of each copy that is the same. `--diff` prints each difference. `--report`
// prints the overlap of each `adapted` and `reference` file with each source. `--only <path>`
// checks the entries and the signals of one tracked file.
// Exit codes: 0 when every check passes, 1 on a failure, 2 on a usage error, and 3 when the
// checks 1 to 3 pass but no DOCX Editor checkout is found. A skip is never a pass.

import { spawnSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { extname, join, relative, resolve } from 'node:path';
import { format, resolveConfig } from 'prettier';
import { findDocxCheckout } from './docx-checkout.ts';
import {
  ADAPTED_FLOOR,
  REFERENCE_CEILING,
  type CopyEntry,
  MENTION,
  MENTIONS_ALLOWED,
  allowedReason,
  comparable,
  comparablePair,
  copySignals,
  folderProblems,
  headerOf,
  headerProblems,
  textOf,
  kindProblems,
  lineOverlap,
  localPaths,
  parseManifest,
  sourceProblems,
  unlistedMentions,
} from './docx-copies.ts';
import { firstDifference, normalizedBody } from './docx-normalize.ts';

const ROOT = resolve(import.meta.dir, '../..');
const MANIFEST = 'scripts/shared/docx-copies.json';

/**
 * Local modules that hold the same exports at another path here, for each copy, as
 * [DOCX Editor module, this repository's module]. Both are module paths from the repository
 * root, without the extension or a trailing `/index` (`localModule` in `docx-normalize.ts`).
 * Each entry maps one module to one module.
 */
const RENAMED_MODULES: ReadonlyMap<string, readonly (readonly [string, string])[]> = new Map([
  // DOCX Editor keeps the font contracts in `contracts/editor.ts`. Here they are in
  // `contracts/editor-fonts.ts`, and the `editor` lane reads them through `contracts/index.ts`.
  [
    'packages/core/src/editor/__tests__/font-composition.test.ts',
    [['packages/core/src/contracts/editor', 'packages/core/src/contracts']],
  ],
  [
    'packages/core/src/editor/custom-fonts.ts',
    [['packages/core/src/contracts/editor', 'packages/core/src/contracts']],
  ],
  [
    'packages/core/src/editor/load-fonts.ts',
    [['packages/core/src/contracts/editor', 'packages/core/src/contracts']],
  ],
  [
    'packages/core/src/layout/font-composition.ts',
    [['packages/core/src/contracts/editor', 'packages/core/src/contracts/editor-fonts']],
  ],
  [
    'packages/core/src/layout/font-deferred.ts',
    [['packages/core/src/contracts/editor', 'packages/core/src/contracts/editor-fonts']],
  ],
  // The text modules live in `packages/ooxml-core/src/` here.
  [
    'packages/ooxml-core/src/__tests__/grapheme.test.ts',
    [['packages/core/src/layout/grapheme', 'packages/ooxml-core/src/grapheme']],
  ],
  [
    'packages/ooxml-core/src/__tests__/word-segment.test.ts',
    [
      ['packages/core/src/layout/grapheme', 'packages/ooxml-core/src/grapheme'],
      ['packages/core/src/layout/word-segment', 'packages/ooxml-core/src/word-segment'],
    ],
  ],
  [
    'packages/ooxml-core/src/word-segment.ts',
    [['packages/core/src/layout/grapheme', 'packages/ooxml-core/src/grapheme']],
  ],
  // The SHA-256 module is shared through the `@eigenpal/ooxml-core` package here.
  [
    'packages/core/src/layout/font-resource.ts',
    [['packages/core/src/store/package/sha256', 'packages/ooxml-core/src']],
  ],
]);

const OPTIONS_WITH_VALUE = new Set(['--docx', '--ref', '--only']);
const FLAGS = new Set(['--write', '--diff', '--report', '--help']);

for (let at = 2; at < process.argv.length; at += 1) {
  const arg = process.argv[at]!;
  if (OPTIONS_WITH_VALUE.has(arg)) at += 1;
  else if (!FLAGS.has(arg)) {
    console.error(
      `Unknown argument: ${arg}. Use --docx, --ref, --write, --diff, --report, --only, --help.`
    );
    process.exit(2);
  }
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

if (process.argv.includes('--help')) {
  const text = readFileSync(join(ROOT, 'scripts/shared/check-docx-copies.ts'), 'utf8');
  console.log(text.split('\n\nimport ')[0]!.replace(/^\/\/ ?/gm, ''));
  process.exit(0);
}

const ref = option('--ref');
const onlyOption = option('--only');
const only = onlyOption === undefined ? undefined : relative(ROOT, resolve(onlyOption));
const write = process.argv.includes('--write');
const showDiff = process.argv.includes('--diff');
const report = process.argv.includes('--report');
if (write && ref === undefined) {
  console.error('--write needs --ref <commit>.');
  process.exit(2);
}

function git(args: readonly string[], cwd: string, input?: string): string | null {
  const run = spawnSync('git', ['-C', cwd, ...args], {
    encoding: 'utf8',
    input,
    maxBuffer: 64 * 1024 * 1024,
  });
  return run.status === 0 ? run.stdout : null;
}

let failures = 0;
const fail = (path: string, message: string): void => {
  failures += 1;
  console.log(`  FAIL ${path}: ${message}`);
};

// Checks 1 to 3 need no DOCX Editor checkout.
const tracked = git(['ls-files', '-z'], ROOT)?.split('\0').filter(Boolean);
if (tracked === undefined) {
  console.error(`FAIL: ${ROOT} is not a git checkout, so no file was checked.`);
  process.exit(1);
}
const trackedSet = new Set(tracked);
if (only !== undefined && !trackedSet.has(only)) {
  console.error(`FAIL: --only ${onlyOption} is not a tracked file of this repository.`);
  process.exit(1);
}
const manifestText = readFileSync(join(ROOT, MANIFEST), 'utf8');
const parsed = parseManifest(manifestText);
for (const problem of parsed.problems) fail(MANIFEST, problem);
if (parsed.entries.length === 0)
  fail(MANIFEST, 'the manifest has no entry, so nothing was checked');
const entries = parsed.entries.filter((entry) => only === undefined || entry.file === only);

const isLocal = localPaths(tracked);
const recorded = parsed.entries.flatMap((entry) => entry.source);
const texts = new Map<string, string>();
const broken = new Set<CopyEntry>();
for (const entry of entries) {
  const file = join(ROOT, entry.file);
  if (!trackedSet.has(entry.file) || !existsSync(file)) {
    broken.add(entry);
    fail(entry.file, 'the manifest names a file that is not tracked or not on disk');
    continue;
  }
  if (!texts.has(entry.file)) texts.set(entry.file, readFileSync(file, 'utf8'));
  const siblings = parsed.entries.filter((other) => other.file === entry.file);
  const options = {
    isLocal,
    sources: siblings.flatMap((other) => other.source),
    copied: siblings.some((other) => other.kind === 'copied'),
    noted: siblings.some((other) => other.reason !== undefined),
    recorded,
  };
  for (const problem of headerProblems(entry, texts.get(entry.file)!, entry.file, options))
    fail(entry.file, `${problem} (${entry.kind} entry)`);
}

const signals = new Map<string, readonly string[]>();
let mentioning = 0;
for (const path of only === undefined ? tracked : [only]) {
  const file = join(ROOT, path);
  if (!existsSync(file)) continue;
  // Every file whose bytes are UTF-8 apart from NUL bytes, whatever git calls it.
  const body = textOf(readFileSync(file));
  if (body === null || !MENTION.test(body)) continue;
  mentioning += 1;
  signals.set(path, copySignals(body, path, isLocal));
}
for (const path of unlistedMentions(signals, parsed.entries))
  fail(path, `${signals.get(path)![0]}, but the file has no entry in ${MANIFEST}`);

const found = findDocxCheckout(option('--docx'));
if (found.kind === 'invalid') {
  console.error(found.message);
  process.exit(2);
}
if (found.kind === 'absent') {
  console.log(
    `\n${failures} failures in the manifest, header, and signal checks.\n` +
      'SKIPPED: no DOCX Editor checkout found, so no commit, source, or copy was checked. Pass\n' +
      '--docx <path> or set DOCX_EDITOR_DIR. This is a local check: a skip is not a pass.'
  );
  process.exit(failures > 0 ? 1 : 3);
}
const docx = found.dir;

function commitOf(name: string): string | null {
  return git(['rev-parse', '--verify', '--quiet', `${name}^{commit}`], docx)?.trim() ?? null;
}

const refCommit = ref === undefined ? undefined : commitOf(ref);
if (refCommit === null) {
  console.error(`No commit ${ref} in ${docx}.`);
  process.exit(2);
}

/** Both normalized bodies as a unified diff, for `--diff`. */
function unifiedDiff(expected: string, actual: string, path: string): string {
  const dir = mkdtempSync(join(tmpdir(), 'docx-copies-'));
  try {
    writeFileSync(join(dir, 'docx'), expected);
    writeFileSync(join(dir, 'here'), actual);
    const labels = ['--label', 'DOCX Editor', '--label', path];
    const run = spawnSync('diff', ['-u', ...labels, join(dir, 'docx'), join(dir, 'here')], {
      encoding: 'utf8',
    });
    return run.stdout.replace(/^/gm, '      ');
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

console.log(`DOCX Editor: ${docx}${ref ? `, copies compared at ${ref}` : ''}\n`);

// Check 4: commits and source paths, in one `cat-file` batch.
const commits = new Map<string, string | null>();
const resolved = (name: string): string | null => {
  if (!commits.has(name)) commits.set(name, commitOf(name));
  return commits.get(name)!;
};
/** The object type of each `<commit>:<path>` that exists: `blob` (a file) or `tree` (a folder). */
const types = new Map<string, string>();
const lookup = {
  commit: resolved,
  exists: (objects: readonly string[]) => {
    const batch = git(['cat-file', '--batch-check'], docx, objects.map((o) => `${o}\n`).join(''));
    const answers = (batch ?? '').split('\n');
    return objects.map((object, at) => {
      const type = answers[at]?.split(' ')[1];
      if (batch === null || answers[at]!.endsWith(' missing') || type === undefined) return false;
      types.set(object, type);
      return true;
    });
  },
};
const live = entries.filter((entry) => !broken.has(entry));
for (const { entry, message } of sourceProblems(live, lookup)) {
  broken.add(entry);
  fail(entry.file, message);
}

/** The normalized body of `text`, after the module renames of `path` when `docx` is set. */
async function body(text: string, path: string, renamesOf?: string): Promise<string> {
  const normalized = await normalizedBody(text, extname(path), path);
  return (RENAMED_MODULES.get(renamesOf ?? '') ?? []).reduce(
    (out, [docxModule, here]) => out.replaceAll(`'<module>/${docxModule}'`, `'<module>/${here}'`),
    normalized
  );
}

// Check 5: each copy against its source.
let same = 0;
const moved: CopyEntry[] = [];
for (const entry of entries) {
  if (entry.kind !== 'copied' || broken.has(entry)) continue;
  const path = entry.file;
  const source = entry.source[0]!;
  const label = ref ?? entry.commit!;
  const theirs = git(['show', `${refCommit ?? resolved(entry.commit!)}:${source}`], docx);
  if (theirs === null) {
    fail(path, `${source} is not in DOCX Editor at ${label}`);
    continue;
  }
  let expected: string;
  let actual: string;
  try {
    expected = await body(theirs, source, path);
    actual = await body(texts.get(path)!, path);
  } catch (error) {
    fail(path, `cannot be compared: ${(error as Error).message}`);
    continue;
  }
  if (expected !== actual) {
    fail(path, `differs from ${source} at ${label}, ${firstDifference(expected, actual)}`);
    if (showDiff) console.log(unifiedDiff(expected, actual, path));
    continue;
  }
  same += 1;
  if (write && refCommit !== undefined && resolved(entry.commit!) !== refCommit) {
    moved.push(entry);
    console.log(`  ok   ${path} (copied, the same at ${ref}; moved there)`);
  } else console.log(`  ok   ${path} (copied, the same)`);
}

// Check 6: the kind of each adapted and reference entry, from its overlap with each source.
let measured = 0;
let lowest = Number.POSITIVE_INFINITY;
let highestReference = 0;
for (const entry of entries) {
  if (entry.kind === 'copied' || broken.has(entry)) continue;
  const commit = resolved(entry.commit!)!;
  const files = entry.source.map(
    (source) => comparablePair(entry.file, source) && types.get(`${commit}:${source}`) === 'blob'
  );
  // The files of each folder source that the check can compare, as [folder, file].
  const inFolders = !comparable(entry.file)
    ? []
    : entry.source
        .filter((source) => types.get(`${commit}:${source}`) === 'tree')
        .flatMap((folder) =>
          (git(['ls-tree', '-r', '--name-only', commit, '--', `${folder}/`], docx) ?? '')
            .split('\n')
            .filter((path) => path !== '' && comparablePair(entry.file, path))
            .map((path) => [folder, path] as const)
        );
  if (!comparable(entry.file) || (!files.some(Boolean) && inFolders.length === 0)) {
    if (entry.kind === 'adapted')
      fail(entry.file, 'an adapted entry needs source files that the check can compare');
    if (report) console.log(`  report ${entry.file} (${entry.kind}): not compared`);
    continue;
  }
  let ours: string;
  const theirs: (string | null)[] = [];
  try {
    ours = await body(texts.get(entry.file)!, entry.file);
    for (const [at, source] of entry.source.entries())
      theirs.push(
        files[at]
          ? await body(git(['show', `${commit}:${source}`], docx) ?? '', source, entry.file)
          : null
      );
  } catch (error) {
    if (entry.kind === 'adapted')
      fail(entry.file, `cannot be compared: ${(error as Error).message}`);
    continue;
  }
  if (entry.kind === 'adapted' && !files.some(Boolean))
    fail(entry.file, 'an adapted entry needs source files that the check can compare');
  else
    for (const problem of kindProblems(entry, ours, theirs))
      fail(entry.file, `${problem} (${entry.kind} entry)`);
  const folderOverlaps: string[] = [];
  for (const folder of new Set(inFolders.map(([name]) => name))) {
    const pairs: [string, string][] = [];
    for (const [, path] of inFolders.filter(([name]) => name === folder)) {
      try {
        pairs.push([
          path,
          await body(git(['show', `${commit}:${path}`], docx) ?? '', path, entry.file),
        ]);
      } catch {
        // A folder file that does not parse is not compared.
      }
    }
    for (const problem of folderProblems(entry, folder, ours, pairs))
      fail(entry.file, `${problem} (${entry.kind} entry)`);
    for (const [path, other] of pairs) {
      const overlap = lineOverlap(other, ours);
      folderOverlaps.push(`${path} ${overlap.toFixed(3)}`);
      if (entry.kind === 'reference') highestReference = Math.max(highestReference, overlap);
    }
  }
  const overlaps = theirs.map((other) => (other === null ? null : lineOverlap(other, ours)));
  if (report)
    console.log(
      `  report ${entry.file} (${entry.kind}): ` +
        entry.source
          .map((source, at) => `${source} ${overlaps[at]?.toFixed(3) ?? 'not compared'}`)
          .concat(folderOverlaps)
          .join(', ')
    );
  if (entry.kind === 'adapted' && files.some(Boolean)) {
    measured += 1;
    const joined = theirs.filter((other): other is string => other !== null).join('\n');
    lowest = Math.min(lowest, lineOverlap(joined, ours));
  } else if (entry.kind === 'reference')
    for (const overlap of overlaps)
      if (overlap !== null) highestReference = Math.max(highestReference, overlap);
}

// `--write`: move each unchanged copy, in its header paragraph and in its entry, to `--ref`.
if (moved.length > 0) {
  const short = git(['rev-parse', '--short=9', refCommit!], docx)!.trim();
  for (const entry of moved) {
    const text = texts.get(entry.file)!;
    const header = headerOf(entry, text, entry.file)!;
    const lines = text.split('\n');
    for (let line = header.first; line <= header.last; line += 1)
      lines[line] = lines[line]!.replace(new RegExp(`(?<![\\w-])${entry.commit}(?![\\w-])`), short);
    writeFileSync(join(ROOT, entry.file), lines.join('\n'));
  }
  const next = JSON.parse(manifestText) as { file: string; kind: string; commit: string }[];
  for (const item of next)
    if (moved.some((entry) => entry.file === item.file && entry.kind === item.kind))
      item.commit = short;
  const options = await resolveConfig(join(ROOT, MANIFEST));
  writeFileSync(
    join(ROOT, MANIFEST),
    await format(JSON.stringify(next), { ...options, parser: 'json' })
  );
  console.log(`\nMoved ${moved.length} copies to ${short}.`);
}

const count = (kind: string) => entries.filter((entry) => entry.kind === kind).length;
const signalled = [...signals].filter(
  ([path, found]) => found.length > 0 && allowedReason(path) === undefined
).length;
console.log(
  `\n${entries.length} manifest entries (${count('copied')} copied, ${count('adapted')} adapted, ` +
    `${count('reference')} references), ${same} copies the same as their source, ` +
    `${measured} adapted files measured (lowest line overlap ` +
    `${measured > 0 ? lowest.toFixed(3) : 'none'}, floor ${ADAPTED_FLOOR}), highest ` +
    `reference overlap ${highestReference.toFixed(3)} (ceiling ${REFERENCE_CEILING}), ` +
    `${mentioning} files that mention DOCX Editor, ${signalled} with a copy signal outside the allowed files, ` +
    `${MENTIONS_ALLOWED.length} allowed patterns, ${failures} failures.`
);
if (failures > 0 && !showDiff) console.log('Run with --diff to print each difference.');
process.exit(failures > 0 ? 1 : 0);
