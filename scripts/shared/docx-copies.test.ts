import { describe, expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import {
  ADAPTED_FLOOR,
  type CopyEntry,
  MENTION,
  MENTIONS_ALLOWED,
  REFERENCE_CEILING,
  allowedReason,
  claimsCopy,
  comparablePair,
  copySignals,
  folderProblems,
  headerProblems,
  kindProblems,
  lineOverlap,
  localPaths,
  parseManifest,
  sourceProblems,
  rawBody,
  textOf,
  unlistedMentions,
} from './docx-copies.ts';
import { localModule, normalizedBody } from './docx-normalize.ts';

/** `normalizedBody` for a file in the layout lane, unless a test names another path. */
const norm = (text: string, ext: string, path = `packages/core/src/layout/x${ext}`) =>
  normalizedBody(text, ext, path);

const ROOT = resolve(import.meta.dir, '../..');

const COPY: CopyEntry = {
  file: 'packages/core/src/layout/font-composition.ts',
  kind: 'copied',
  source: ['packages/core/src/layout/font-composition.ts'],
  commit: 'adf0e178',
};
const COPY_HEADER =
  '// Copied from DOCX Editor `packages/core/src/layout/font-composition.ts` at commit adf0e178.\n' +
  'export const a = 1;\n';

/** A checkout that holds `files` at the commit `adf0e178`, and no other commit. */
const lookup = (files: readonly string[]) => ({
  commit: (name: string) => (name === 'adf0e178' ? 'adf0e178'.repeat(5) : null),
  exists: (objects: readonly string[]) =>
    objects.map((object) => files.includes(object.slice(object.indexOf(':') + 1))),
});

/** The signals of `text` in a TypeScript file, with no local paths. */
const signalsOf = (text: string, path = 'packages/core/src/x.ts') => copySignals(text, path);

describe('parseManifest', () => {
  test('reads the repository manifest without a problem', () => {
    const text = readFileSync(join(import.meta.dir, 'docx-copies.json'), 'utf8');
    const { entries, problems } = parseManifest(text);
    expect(problems).toEqual([]);
    expect(entries.filter((entry) => entry.kind === 'copied').length).toBeGreaterThan(0);
  });

  test('reports each entry with a wrong shape', () => {
    const { problems } = parseManifest(
      JSON.stringify([
        { file: 'a.ts', kind: 'copied', source: ['a.ts'], commit: null },
        { file: 'b.ts', kind: 'copied', source: ['b.ts', 'c.ts'], commit: 'adf0e178' },
        { file: 'c.ts', kind: 'reference', source: [], commit: 'adf0e178' },
        { file: 'd.ts', kind: 'ported', source: ['d.ts'], commit: 'adf0e178' },
        { file: 'e.ts', kind: 'reference', source: ['e.ts'], commit: 'not-a-hash' },
        { file: 'f.ts', kind: 'adapted', source: ['f.ts'], commit: 'adf0e178', line: 3 },
        { file: 'g.ts', kind: 'adapted', source: ['g.ts'], commit: 'adf0e178' },
        { file: 'g.ts', kind: 'adapted', source: ['g.ts'], commit: 'adf0e178' },
        { file: 'h.ts', kind: 'reference', source: ['h.ts'], commit: null },
      ])
    );
    expect(problems).toEqual([
      'a.ts: an entry must name the commit of its sources',
      'b.ts: a copied entry names one source',
      'c.ts: an entry must name a source',
      'd.ts: `kind` must be copied, adapted, reference',
      'e.ts: `commit` must be a commit hash or null',
      'f.ts: unknown fields line',
      'g.ts: a second adapted entry with the same sources at adf0e178',
      'h.ts: an entry must name the commit of its sources',
    ]);
  });
});

describe('sourceProblems', () => {
  test('passes when each commit and source exists', () => {
    expect(sourceProblems([COPY], lookup(COPY.source))).toEqual([]);
  });

  test('fails a manifest commit that does not exist', () => {
    const entry = { ...COPY, commit: 'deadbee12' };
    expect(sourceProblems([entry], lookup(COPY.source))).toEqual([
      { entry, message: 'no commit deadbee12 in DOCX Editor' },
    ]);
  });

  test('fails a manifest path that does not exist', () => {
    const reference: CopyEntry = { ...COPY, kind: 'reference', source: ['a/b'] };
    expect(sourceProblems([COPY, reference], lookup([])).map((p) => p.message)).toEqual([
      'packages/core/src/layout/font-composition.ts is not in DOCX Editor at adf0e178',
      'a/b is not in DOCX Editor at adf0e178',
    ]);
  });
});

describe('headerProblems', () => {
  test('passes one comment paragraph that names the kind, the source, and the commit', () => {
    expect(headerProblems(COPY, COPY_HEADER)).toEqual([]);
    expect(headerProblems(COPY, COPY_HEADER.toLowerCase())).toEqual([]);
    const wrapped =
      '/*\n * Copied from DOCX Editor\n * `packages/core/src/layout/font-composition.ts`\n' +
      ' * at commit adf0e178.\n */\n';
    expect(headerProblems(COPY, wrapped)).toEqual([]);
    const yaml =
      '# Copied from docx-editor/packages/core/src/layout/font-composition.ts at adf0e178\n';
    expect(headerProblems(COPY, yaml, 'ci.yml')).toEqual([]);
  });

  test('fails a header that disagrees with the manifest', () => {
    const header = COPY_HEADER.replace('adf0e178', 'deadbee12');
    expect(headerProblems(COPY, header)).toEqual([
      'no comment paragraph names all of the entry; the closest, at line 1, lacks commit adf0e178',
    ]);
    expect(headerProblems({ ...COPY, kind: 'adapted' }, COPY_HEADER)).toEqual([
      'no comment paragraph names all of the entry; the closest, at line 1, lacks "adapted"',
      'line 1 puts a copy word near DOCX Editor, but the entry is adapted',
    ]);
    expect(headerProblems(COPY, 'export const a = 1;\n')).toEqual([
      'the file has no comment paragraph',
    ]);
  });

  test('takes the kind, the source, and the commit from one paragraph, each as a whole word', () => {
    const cases = [
      // The real path in another paragraph.
      '// Copied from DOCX Editor `packages/core/src/layout/nope.ts` at adf0e178.\n' +
        'export const a = 1;\n// See packages/core/src/layout/font-composition.ts.\n',
      // The commit in another paragraph.
      '// Copied from DOCX Editor `packages/core/src/layout/font-composition.ts` at deadbee12.\n' +
        '\n// cache salt: adf0e178\n',
      // A longer path, a longer commit, and a longer kind word.
      '// Copied from DOCX Editor `packages/core/src/layout/font-composition.ts.bak` at adf0e178.\n',
      '// Copied from DOCX Editor `packages/core/src/layout/font-composition.ts` at adf0e1789.\n',
      '// Uncopied from DOCX Editor `packages/core/src/layout/font-composition.ts` at adf0e178.\n',
      // The paragraph in code, not in a comment.
      "const a = 'Copied from DOCX Editor packages/core/src/layout/font-composition.ts adf0e178';\n",
    ];
    for (const text of cases) expect(headerProblems(COPY, text)).not.toEqual([]);
  });

  test('fails a copy claim in a file without a copied entry, in any case and across lines', () => {
    const adapted: CopyEntry = { ...COPY, kind: 'adapted' };
    const header = COPY_HEADER.replace('Copied', 'Adapted');
    expect(headerProblems(adapted, header)).toEqual([]);
    for (const claim of [
      '// This module is copied\n// from DOCX Editor.\n',
      '/* a copy of docx-editor/a.ts */\n',
      '# COPIED VERBATIM FROM DOCX EDITOR\n',
    ])
      expect(headerProblems(adapted, `${header}\n${claim}`, 'a.sh').at(-1)).toContain(
        'puts a copy word near DOCX Editor'
      );
    // A Markdown guide that describes the copies in other files gives a reason. In code, a
    // reason is a note only.
    const guide = { ...adapted, reason: 'The guide describes the copies of this package.' };
    const claim = '// a copy of DOCX Editor\n';
    expect(headerProblems(guide, `${header}\n${claim}`)).not.toEqual([]);
    const md =
      '<!-- Adapted from DOCX Editor `packages/core/src/layout/font-composition.ts` at ' +
      'adf0e178. -->\n\nThe modules are copies of DOCX Editor files.\n';
    expect(headerProblems(guide, md, 'README.md')).toEqual([]);
    expect(headerProblems(adapted, md, 'README.md')).not.toEqual([]);
  });

  test('reads a copy word near a mention, in the same sentence and outside code spans', () => {
    for (const text of [
      'The modules are copies of DOCX Editor files',
      'DOCX Editor code, copied with one change',
      'a DocxEditor copy',
    ])
      expect(claimsCopy(text)).toBe(true);
    for (const text of [
      'The second copy of the editor state. DOCX Editor keeps one.',
      "DOCX Editor's `Copy` button",
      'DOCX Editor draws the shape in the same order as the slide, so this long sentence is a copy',
    ])
      expect(claimsCopy(text)).toBe(false);
  });

  test('fails a paragraph that names a foreign file that no entry lists', () => {
    const isLocal = localPaths([COPY.file]);
    const header =
      '// Copied from DOCX Editor `packages/core/src/layout/font-composition.ts` at commit ' +
      'adf0e178, with\n// `packages/core/src/layout/font-names.ts`.\n';
    expect(headerProblems(COPY, header, COPY.file, { isLocal })).toEqual([
      'line 1 names packages/core/src/layout/font-names.ts, which no entry of the file lists',
    ]);
    // Any paragraph that mentions DOCX Editor, not only the header.
    const second = `${COPY_HEADER}\n// DOCX Editor's \`packages/core/src/layout/other.ts\` too.\n`;
    expect(headerProblems(COPY, second, COPY.file, { isLocal })).toEqual([
      'line 4 names packages/core/src/layout/other.ts, which no entry of the file lists',
    ]);
    // A guide may name a source that the manifest records for another file.
    const guide = second.replaceAll('// ', '');
    const recorded = ['packages/core/src/layout/other.ts'];
    expect(headerProblems(COPY, guide, 'a.md', { isLocal, recorded })).not.toContainEqual(
      expect.stringContaining('other.ts')
    );
    const sources = [...COPY.source, 'packages/core/src/layout/font-names.ts'];
    expect(headerProblems(COPY, header, COPY.file, { isLocal, sources })).toEqual([]);
    // A folder that holds a source is covered by it.
    const folder = header.replace('font-names.ts', '');
    expect(headerProblems(COPY, folder, COPY.file, { isLocal })).toEqual([]);
  });

  // Each review found a comment form that an earlier check missed. The manifest records each
  // file, so a change to the comment fails against the manifest.
  test('fails each review repro against the manifest', () => {
    const reference = (source: string[], commit: string): CopyEntry => ({
      file: 'x.ts',
      kind: 'reference',
      source,
      commit,
    });
    const cases: [CopyEntry, string][] = [
      // A commit that changes in a reference.
      [reference(['a/b.ts'], '26095c63'), "// After DOCX Editor's `a/b.ts` at deadbee12.\n"],
      // A path that changes in a reference.
      [
        reference(['packages/core/src/layout/caret-word.ts'], '5993bd223'),
        '// The rule of DOCX Editor (`packages/core/src/layout/nope.ts` at 5993bd223).\n',
      ],
      // A note added to prose about another file.
      [
        reference(['a/b.ts'], 'adf0e178'),
        "// After DOCX Editor's `a/b.ts` at adf0e178. Copied from DOCX Editor `a.ts`.\n",
      ],
      // "A copy of", a far verb, and an "after" header with another commit.
      [COPY, COPY_HEADER.replace('Copied from', 'A copy of').replace('adf0e178', 'deadbee12')],
      [COPY, COPY_HEADER.replace('adf0e178.', 'deadbee12.\n// Copied, with no change.')],
      // A path without an extension, a path after the first, and a commit after the first.
      [
        { ...COPY, kind: 'adapted', source: ['packages/i18n', 'packages/a/b.ts'] },
        "// Adapted from DOCX Editor's `packages/i18n` at adf0e178.\n",
      ],
      [{ ...COPY, kind: 'adapted', commit: 'abc12345' }, COPY_HEADER.replace('Copied', 'Adapted')],
    ];
    for (const [entry, text] of cases) expect(headerProblems(entry, text)).not.toEqual([]);
  });
});

describe('kindProblems', () => {
  const adapted: CopyEntry = { ...COPY, kind: 'adapted' };
  const body = 'export function scale(a) {\n  return a * 2;\n}\nexport const unit = 4;\n';

  const other = 'export function format(n) {\n  return String(n);\n}\n';
  const two: CopyEntry = { ...adapted, source: ['a.ts', 'b.ts'] };

  test('fails an adapted or reference entry that equals any of its sources', () => {
    expect(kindProblems(adapted, body, [body])).toEqual([
      `the code equals ${COPY.source[0]}: mark the entry as copied`,
    ]);
    expect(kindProblems({ ...adapted, kind: 'reference' }, body, [body])).toHaveLength(1);
    expect(kindProblems(COPY, body, [body])).toEqual([]);
    expect(kindProblems(two, body, [other, body])).toEqual([
      'the code equals b.ts: mark the entry as copied',
    ]);
    expect(kindProblems({ ...two, kind: 'reference' }, body, [null, body])).toHaveLength(1);
  });

  test('fails each source of an adapted entry under the floor on its own, with or without a reason', () => {
    const changed = body.replace('a * 2', 'a * 3');
    expect(kindProblems(two, body, [changed, other])).toEqual([
      'the code shares 0.000 of its lines with b.ts, under 0.1: fix the source, or move it to a ' +
        'reference entry',
    ]);
    const why = { ...two, reason: 'The second source holds the type only.' };
    expect(kindProblems(why, body, [changed, other])).toHaveLength(1);
    expect(kindProblems(two, body, [changed, null])).toEqual([
      'the check cannot compare b.ts with the code: name a file in the same language, or move ' +
        'it to a reference entry',
    ]);
    expect(kindProblems({ ...two, kind: 'reference' }, body, [other, other])).toEqual([]);
  });

  test('fails a folder file that equals the code, and a reference over the ceiling with one', () => {
    const changed = body.replace('a * 2', 'a * 3');
    const reference: CopyEntry = { ...two, kind: 'reference', source: ['src/menu'] };
    expect(folderProblems(reference, 'src/menu', body, [['src/menu/a.ts', body]])).toEqual([
      'the code equals src/menu/a.ts in src/menu: mark the entry as copied',
    ]);
    expect(folderProblems(reference, 'src/menu', changed, [['src/menu/a.ts', body]])).toEqual([
      'the code shares 0.667 of its lines with src/menu/a.ts in src/menu, over 0.5: mark the ' +
        'entry as adapted',
    ]);
    expect(folderProblems(reference, 'src/menu', body, [['src/menu/b.ts', other]])).toEqual([]);
    expect(
      folderProblems({ ...reference, kind: 'adapted' }, 'src/menu', changed, [
        ['src/menu/a.ts', body],
      ])
    ).toEqual([]);
  });

  test('fails a reference over the ceiling with any one source', () => {
    const reference: CopyEntry = { ...two, kind: 'reference' };
    const changed = body.replace('a * 2', 'a * 3');
    expect(lineOverlap(body, changed)).toBeGreaterThan(REFERENCE_CEILING);
    expect(kindProblems(reference, changed, [other, body])).toEqual([
      'the code shares 0.667 of its lines with b.ts, over 0.5: mark the entry as adapted',
    ]);
    expect(kindProblems(two, changed, [body, body])).toEqual([]);
  });

  test('fails an adapted entry under the line overlap floor of its sources together', () => {
    expect(lineOverlap(other, body)).toBeLessThan(ADAPTED_FLOOR);
    expect(kindProblems(adapted, body, [other])[0]).toContain('under 0.1');
    const changed = body.replace('a * 2', 'a * 3');
    expect(lineOverlap(body, changed)).toBeGreaterThan(ADAPTED_FLOOR);
    expect(kindProblems(adapted, changed, [body])).toEqual([]);
    expect(kindProblems({ ...adapted, kind: 'reference' }, body, [other])).toEqual([]);
  });

  test('reads a file as text when its bytes are UTF-8 once the NUL bytes go', () => {
    expect(textOf(new TextEncoder().encode('a\0b'))).toBe('ab');
    // Bytes that are not UTF-8 read as Latin-1, so a mention in them is still found.
    const latin1 = new Uint8Array([0xe9, ...new TextEncoder().encode(' DOCX Editor')]);
    expect(textOf(latin1)).toBe('\u00e9 DOCX Editor');
  });

  test('scans a staged file with a NUL byte, which git calls binary', () => {
    const dir = mkdtempSync(join(tmpdir(), 'docx-copies-'));
    try {
      const run = (...args: string[]) =>
        spawnSync('git', ['-C', dir, ...args], { encoding: 'utf8' }).stdout;
      run('init', '-q');
      const claim = '// \0\n// Copied from DOCX Editor `packages/core/src/a.ts` at 5993bd223.\n';
      writeFileSync(join(dir, 'a.ts'), claim);
      run('add', 'a.ts');
      expect(run('ls-files', '--eol', 'a.ts')).toStartWith('i/-text');
      expect(copySignals(textOf(readFileSync(join(dir, 'a.ts'))), 'a.ts')).not.toEqual([]);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  test('compares a guide or a data file by its trimmed raw lines', () => {
    expect(comparablePair('reference/a.md', 'reference/a.md')).toBe(true);
    expect(comparablePair('a.json', 'b.json')).toBe(true);
    expect(comparablePair('a.md', 'a.ts')).toBe(false);
    const ours = rawBody('# Colors\n\n  The theme has twelve colors.\n');
    expect(lineOverlap(rawBody('# Colors\nThe theme has twelve colors.\n'), ours)).toBe(1);
    const reference: CopyEntry = { ...COPY, kind: 'reference', file: 'a.md', source: ['a.md'] };
    expect(
      kindProblems(reference, ours, [rawBody('# Colors\nThe theme has twelve colors.\nMore.\n')])
    ).toEqual([
      'the code shares 1.000 of its lines with a.md, over 0.5: mark the entry as adapted',
    ]);
  });

  test('counts lines with repeats, and not lines without a letter or a digit', () => {
    expect(lineOverlap('a\n}\n}\n', 'a\n}\n')).toBe(1);
    expect(lineOverlap('a\na\nb\n', 'a\nc\nd\n')).toBe(1 / 3);
    expect(lineOverlap('}\n', '}\n')).toBe(0);
  });
});

describe('the copy signals', () => {
  test('finds a mention in any wording, across a comment line break, and in a path', () => {
    for (const text of [
      "// Copied verbatim, after DOCX Editor's `a.ts`.",
      '/* code */ const a = 1; /* A copy of DOCX Editor */',
      '// as DOCX\n// Editor does',
      ' * as DOCX\n * Editor does',
      '# after docx editor',
      '// docx-editor keeps the same list',
      '// the docx_editor fork',
      "const url = 'https://github.com/eigenpal/docx-editor/blob/main/README.md';",
      '<!-- the DOCX Editor demo -->',
      'const a = DocxEditor.create();',
      '// https://github.com/eigenpal/docx-js-editor/blob/main/a.ts',
      'packages/editor-api/compat/docxeditor/declarations.ts',
    ])
      expect(MENTION.test(text)).toBe(true);
  });

  test('does not read a package name or another word as a mention', () => {
    for (const text of [
      "import { a } from '@docx-editor.dev/fonts';",
      'a docx file in the editor',
      "import { a } from '@docxeditor.dev/core';",
      'see https://docx-editor.dev/docs',
    ])
      expect(MENTION.test(text)).toBe(false);
  });

  test('needs an entry for a structural signal only', () => {
    const isLocal = localPaths(['packages/react/src/ChromeTooltip.tsx']);
    const signal = (text: string) => copySignals(text, 'packages/vue/src/a.ts', isLocal);
    for (const text of [
      "const url = 'https://github.com/eigenpal/docx-editor/blob/main/README.md';",
      '// After DOCX Editor (5993bd223), with the same rule.',
      '// One list (docx-editor keeps the same list in scripts/lib/packages.mjs).',
      "// As DOCX Editor's `surface-pointer.ts` does.",
      '// The menu, copied\n// from DOCX Editor.',
      '/* code */ const a = 1; /* adapted from DOCX Editor */',
      '// The harness, forked from DOCX Editor.',
      '// After DOCX Editor. The rule is the same code, line for line.',
      '// The menu as DOCX Editor has it, a copy with two rows less.',
      '// As DOCX Editor does.\n\n// See commit 5993bd223.',
      '// A port of DOCX Editor.',
      '// It duplicates DOCX Editor.',
      '// Carried over from DOCX Editor.',
      '// Identical to DOCX Editor.',
    ])
      expect(signal(text)).not.toEqual([]);
    for (const text of [
      "// As DOCX Editor's toolbar does.",
      "// The Vue twin of the React adapter's `ChromeTooltip.tsx`, after DOCX Editor's `Tooltip`.",
      "// The editor follows DOCX Editor's names, as in Office.js and React/Vue.",
      '// The color, #1a2b3c4d, as DOCX Editor draws it.',
      "// The second copy of the slide list keeps the order. DOCX Editor's list has none.",
      '// As DOCX Editor does.\n\n// One.\n\n// Two.\n\n// See commit 5993bd223.',
      // Words that name a design, not a copy.
      '// It mirrors the value into the root, as DOCX Editor does.',
      '// It reuses the slide list, as DOCX Editor does.',
    ])
      expect(signal(text)).toEqual([]);
  });

  test('fails a new file with a signal but no entry, of any type', () => {
    const signals = new Map<string, readonly string[]>([
      [COPY.file, ['line 1: a commit hash']],
      ['packages/vue/src/New.vue', ['line 1: a docx-editor/ path']],
      ['scripts/new.cjs', ['line 3: a copy verb near the mention']],
      ['.husky/post-commit', ['line 2: a commit hash']],
      ['packages/react/src/new.jsx', ['line 1: a path right after the mention']],
      ['packages/react/src/plain.tsx', []],
      ['docs/site/content/guide.mdx', ['line 4: a commit hash']],
      ['packages/core/README.md', ['line 1: a docx-editor/ path']],
    ]);
    expect(unlistedMentions(signals, [COPY])).toEqual([
      'packages/vue/src/New.vue',
      'scripts/new.cjs',
      '.husky/post-commit',
      'packages/react/src/new.jsx',
      // A package guide is outside the root and `docs/` guides.
      'packages/core/README.md',
    ]);
    expect(signalsOf('// see DOCX Editor (abc1234f)')).toEqual(['line 1: a commit hash']);
  });

  test('gives a reason for each allowed file', () => {
    expect(MENTIONS_ALLOWED.length).toBeLessThan(14);
    for (const [pattern, reason] of MENTIONS_ALLOWED) {
      expect(reason.length).toBeGreaterThan(10);
      expect(allowedReason(pattern.replaceAll('**', 'a').replaceAll('*', 'a'))).toBe(reason);
    }
    expect(allowedReason('scripts/shared/new-copy.ts')).toBeUndefined();
    // `*.md` is the root folder only, and `docs/**/*.md` any folder under `docs/`.
    expect(allowedReason('CONTRIBUTING.md')).toBeDefined();
    expect(allowedReason('docs/architecture/engine-lanes.md')).toBeDefined();
    expect(allowedReason('packages/core/README.md')).toBeUndefined();
    expect(allowedReason('reference/README.md')).toBeUndefined();
  });
});

// The checks that need no DOCX Editor checkout, over the tracked files, so CI runs them. The
// full check, with a checkout, is `bun run check:docx-copies`.
describe('the repository', () => {
  const run = spawnSync('git', ['-C', ROOT, 'ls-files', '-z'], { encoding: 'utf8' });
  const tracked = run.stdout.split('\0').filter(Boolean);
  const { entries } = parseManifest(
    readFileSync(join(import.meta.dir, 'docx-copies.json'), 'utf8')
  );

  const isLocal = localPaths(tracked);

  test('names each manifest entry in one comment paragraph of its file', () => {
    expect(run.status).toBe(0);
    const problems = entries.flatMap((entry) => {
      const file = join(ROOT, entry.file);
      if (!tracked.includes(entry.file) || !existsSync(file)) return [`${entry.file}: missing`];
      const siblings = entries.filter((other) => other.file === entry.file);
      const options = {
        isLocal,
        sources: siblings.flatMap((other) => other.source),
        copied: siblings.some((other) => other.kind === 'copied'),
        noted: siblings.some((other) => other.reason !== undefined),
        recorded: entries.flatMap((other) => other.source),
      };
      return headerProblems(entry, readFileSync(file, 'utf8'), entry.file, options).map(
        (p) => `${entry.file}: ${p}`
      );
    });
    expect(problems).toEqual([]);
  });

  test('has an entry for each tracked file with a copy signal', () => {
    const signals = new Map<string, readonly string[]>();
    for (const path of tracked) {
      const file = join(ROOT, path);
      if (!existsSync(file)) continue;
      const body = textOf(readFileSync(file));
      if (MENTION.test(body)) signals.set(path, copySignals(body, path, isLocal));
    }
    expect(signals.size).toBeGreaterThan(0);
    expect(unlistedMentions(signals, entries)).toEqual([]);
  });
});

describe('normalizedBody', () => {
  test('removes comments, the header, a license, and a shebang, and reprints the layout', async () => {
    const copy =
      '#!/usr/bin/env node\n/*\nCopyright (c) 2026 Eigenpal, Inc.\n*/\n' +
      '// Copied from DOCX Editor `a/b.mjs` at commit 1234abcd.\n' +
      'export function f(a, b) { return a +\n b; } // the sum\n';
    const source =
      '#!/usr/bin/env node\n// The sum.\nexport function f(a, b) {\n  return a + b;\n}\n';
    expect(await norm(copy, '.mjs')).toBe(await norm(source, '.mjs'));
  });

  test('resolves local imports to module paths, and compares the package scopes', async () => {
    const copy =
      "import type { A } from '../contracts/editor.ts';\n" +
      "import { b } from '../../../ooxml-core/src/index.ts';\n" +
      "export const name = '@pptx-editor.dev/fonts';\n";
    const source =
      "import type { A } from '@docx-editor.dev/core/contracts/editor';\n" +
      "import { b } from '@eigenpal/ooxml-core';\n" +
      "export const name = '@docx-editor.dev/fonts';\n";
    expect(await norm(copy, '.ts')).toBe(await norm(source, '.ts'));
    expect(localModule('./font-resource.ts', 'packages/core/src/layout/a.ts')).toBe(
      'packages/core/src/layout/font-resource'
    );
    expect(localModule('node:path', 'packages/core/src/layout/a.ts')).toBeNull();
  });

  test('keeps a change of import target', async () => {
    expect(await norm("import { a } from '../contracts/index.ts';\n", '.ts')).not.toBe(
      await norm("import { a } from '../layout/index.ts';\n", '.ts')
    );
    expect(await norm("import { a } from './font-resource.ts';\n", '.ts')).not.toBe(
      await norm("import { a } from '../evil/font-resource.ts';\n", '.ts')
    );
    expect(await norm("import { a } from './font-resource.ts';\n", '.ts')).not.toBe(
      await norm("import { a } from '@eigenpal/evil/font-resource.ts';\n", '.ts')
    );
  });

  test('fails on code that does not parse', async () => {
    expect(norm('export const a = 1;\n}\n', '.ts')).rejects.toThrow('does not parse');
    expect(norm('f();\n)))\ng();\n', '.ts')).rejects.toThrow('does not parse');
  });

  test('keeps JSX text that looks like a comment, and bundler annotations', async () => {
    expect(await norm('export const a = <p>// one</p>;\n', '.tsx')).not.toBe(
      await norm('export const a = <p>// two</p>;\n', '.tsx')
    );
    expect(await norm('export const a = /*#__PURE__*/ f();\n', '.ts')).not.toBe(
      await norm('export const a = f();\n', '.ts')
    );
    expect(await norm('// see the @jsx docs\nexport const a = 1;\n', '.ts')).toBe(
      await norm('export const a = 1;\n', '.ts')
    );
  });

  test('ignores the override modifier and constant bracket access', async () => {
    const copy =
      "class E extends Error {\n  override readonly name = 'E';\n}\nx['href'] = undefined;\n";
    const source = "class E extends Error {\n  readonly name = 'E';\n}\nx.href = undefined;\n";
    expect(await norm(copy, '.ts')).toBe(await norm(source, '.ts'));
  });

  test('keeps directive comments', async () => {
    const plain = 'export const a: number = 1;\n';
    expect(await norm(`// @ts-expect-error\n${plain}`, '.ts')).not.toBe(await norm(plain, '.ts'));
    expect(await norm(`/* eslint-disable */\n${plain}`, '.ts')).not.toBe(await norm(plain, '.ts'));
  });

  test('rewrites code, not the text of strings', async () => {
    expect(await norm(`const a = "x['y']";\n`, '.ts')).not.toBe(
      await norm("const a = 'x.y';\n", '.ts')
    );
    expect(await norm("const a = 'override readonly b';\n", '.ts')).not.toBe(
      await norm("const a = 'readonly b';\n", '.ts')
    );
  });

  test('keeps the file name of a local import', async () => {
    expect(await norm("export { a } from './a.ts';\n", '.ts')).not.toBe(
      await norm("export { a } from './b.ts';\n", '.ts')
    );
    expect(await norm("const m = await import('./a.mjs');\n", '.mjs')).not.toBe(
      await norm("const m = await import('./b.mjs');\n", '.mjs')
    );
  });

  test('keeps hash lines inside a multi-line shell string, and reads a here-string', async () => {
    expect(await norm('echo "a\n# one\n"\n', '.sh')).not.toBe(
      await norm('echo "a\n# two\n"\n', '.sh')
    );
    expect(await norm('x <<< foo\n# c1\ny\n', '.sh')).toBe(await norm('x <<< foo\ny\n', '.sh'));
  });

  test('keeps heredoc lines and the text after a hash on a code line', async () => {
    const script = "cat <<EOF\n# Signed\nEOF\necho '{} # x'\n";
    expect(await norm(script, '.sh')).not.toBe(await norm(script.replace('# Signed\n', ''), '.sh'));
    expect(await norm(script, '.sh')).not.toBe(await norm(script.replace('# x', '# y'), '.sh'));
    expect(await norm(`# A comment.\n\n${script}`, '.sh')).toBe(await norm(script, '.sh'));
  });

  test('keeps a code difference and a third-party import difference', async () => {
    expect(await norm('export const a = 1;\n', '.ts')).not.toBe(
      await norm('export const a = 2;\n', '.ts')
    );
    expect(await norm("import { ref } from 'vue';\n", '.ts')).not.toBe(
      await norm("import { ref } from './vue.ts';\n", '.ts')
    );
  });

  test('compares hash comment files without comments and blank lines', async () => {
    const copy = '# Copied from DOCX Editor `a/b.py` at commit 1234abcd.\n\nimport os\n';
    const source = '# A comment.\nimport os\n';
    expect(await norm(copy, '.py')).toBe(await norm(source, '.py'));
  });
});
