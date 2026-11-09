// The release plan never goes major by accident (DOCX Editor's scripts/release-plan.test.ts). DOCX
// Editor commit adf0e178.
//
// The packages are one fixed group, and the adapters have `core` as a peer. By
// default changesets bumps a peer dependent to major whenever the peer changes, and the
// fixed group then takes every package to major. `onlyUpdatePeerDependentsWhenOutOfRange`
// and the tilde peer ranges (`~0.1.0`, as DOCX Editor) keep 0.x minors as minors.
import { afterEach, expect, test } from 'bun:test';
import getReleasePlanModule from '@changesets/get-release-plan';
import { execFileSync } from 'node:child_process';
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const getReleasePlan =
  (getReleasePlanModule as unknown as { default?: typeof getReleasePlanModule }).default ??
  getReleasePlanModule;
const root = join(import.meta.dir, '..');
const directories: string[] = [];
afterEach(() => {
  for (const directory of directories.splice(0))
    rmSync(directory, { recursive: true, force: true });
});

const read = (path: string): Record<string, unknown> =>
  JSON.parse(readFileSync(join(root, path), 'utf8')) as Record<string, unknown>;

/** A copy of the workspace manifests and changeset config, at `version` when given. */
function fixture(version?: string): { directory: string; publicPackages: string[] } {
  const directory = mkdtempSync(join(tmpdir(), 'pptx-release-plan-'));
  directories.push(directory);
  const write = (path: string, value: unknown): void => {
    const destination = join(directory, path);
    mkdirSync(dirname(destination), { recursive: true });
    writeFileSync(destination, typeof value === 'string' ? value : JSON.stringify(value));
  };
  write('.changeset/config.json', read('.changeset/config.json'));
  // The config names the repository changelog generator by relative path.
  const generator = 'scripts/changesets/changelog.mjs';
  mkdirSync(join(directory, dirname(generator)), { recursive: true });
  copyFileSync(join(root, generator), join(directory, generator));
  // Workspace discovery uses the lockfile to identify Bun projects.
  write('bun.lock', {});
  write('package.json', read('package.json'));
  const publicPackages: string[] = [];
  for (const parent of ['packages', 'examples']) {
    for (const entry of readdirSync(join(root, parent), { withFileTypes: true })) {
      const path = join(parent, entry.name, 'package.json');
      // `examples/shared/` and the assembled preview (`examples/parity/`) are not packages.
      if (!entry.isDirectory() || !existsSync(join(root, path))) continue;
      const manifest = read(path);
      if (!manifest.private) {
        publicPackages.push(manifest.name as string);
        if (version) manifest.version = version;
      }
      write(path, manifest);
    }
  }
  return { directory, publicPackages: publicPackages.sort() };
}

async function plan(directory: string) {
  const result = await getReleasePlan(directory);
  return result.releases.filter((release) => release.type !== 'none');
}

test('the fixed group matches the public packages', () => {
  const { publicPackages } = fixture();
  const config = read('.changeset/config.json') as { fixed: string[][] };
  expect(config.fixed.flat().sort()).toEqual(publicPackages);
});

function copyPendingChangesets(directory: string): void {
  for (const file of readdirSync(join(root, '.changeset')))
    if (file.endsWith('.md') && file !== 'README.md')
      writeFileSync(
        join(directory, '.changeset', file),
        readFileSync(join(root, '.changeset', file))
      );
}

test('the pending changesets release no major version', async () => {
  const { directory } = fixture();
  copyPendingChangesets(directory);
  const releases = await plan(directory);
  expect(releases.some((release) => release.type === 'major')).toBe(false);
  expect(new Set(releases.map((release) => release.newVersion)).size).toBeLessThanOrEqual(1);
});

// A changelog entry says what changed for the reader in one short paragraph. Details go in
// the documentation. The fixed group writes each entry into every package that it names.
const SUMMARY_MAX_CHARACTERS = 400;

test('each pending changeset is one short paragraph', async () => {
  const { directory } = fixture();
  copyPendingChangesets(directory);
  const { changesets } = await getReleasePlan(directory);
  const long = changesets
    .filter(
      (changeset) =>
        changeset.summary.length > SUMMARY_MAX_CHARACTERS || changeset.summary.includes('\n')
    )
    .map((changeset) => `.changeset/${changeset.id}.md (${changeset.summary.length} characters)`);
  expect(long).toEqual([]);
});

// changesets/action omits every changelog from the release PR body when the body passes
// 60,000 characters. Estimate the body with room for commit links and headings.
test('the release PR body keeps its changelogs', async () => {
  const { directory } = fixture();
  copyPendingChangesets(directory);
  const { changesets, releases } = await getReleasePlan(directory);
  const summaries = new Map(changesets.map((changeset) => [changeset.id, changeset.summary]));
  let characters = 0;
  for (const release of releases) {
    if (release.type === 'none') continue;
    characters += 200;
    for (const id of release.changesets) characters += (summaries.get(id)?.length ?? 0) + 120;
  }
  expect(characters).toBeLessThan(60_000);
});

test('a minor change after 0.1.0 releases 0.2.0, not 1.0.0', async () => {
  const { directory, publicPackages } = fixture('0.1.0');
  writeFileSync(
    join(directory, '.changeset', 'core-feature.md'),
    '---\n"@pptx-editor.dev/core": minor\n---\nAdd a feature.\n'
  );
  const releases = await plan(directory);
  expect(releases.map((release) => release.name).sort()).toEqual(publicPackages);
  for (const release of releases) {
    expect(release.type).toBe('minor');
    expect(release.newVersion).toBe('0.2.0');
  }
});

test('a dependent package gets one dependency line, not one per changeset', () => {
  // At 0.1.0 the real `~0.2.0` peer ranges are out of range, so dependents get a
  // dependency line. An in-range patch writes no dependency line at all.
  const { directory } = fixture('0.1.0');
  for (const name of ['first', 'second'])
    writeFileSync(
      join(directory, '.changeset', `core-${name}.md`),
      `---\n"@pptx-editor.dev/core": patch\n---\nFix the ${name} thing.\n`
    );
  execFileSync(
    process.execPath,
    [fileURLToPath(import.meta.resolve('@changesets/cli/bin.js')), 'version'],
    { cwd: directory, timeout: 30_000, stdio: 'pipe' }
  );
  const changelog = readFileSync(join(directory, 'packages/react/CHANGELOG.md'), 'utf8');
  const lines = changelog.split('\n').filter((line) => line.startsWith('- Updated dependencies'));
  expect(lines).toHaveLength(1);
  expect(lines[0]).toContain('@pptx-editor.dev/core@0.1.1');
  expect(changelog).not.toContain('Updated dependencies [');
});
