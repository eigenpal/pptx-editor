// The release plan never goes major by accident (DOCX Editor's scripts/release-plan.test.ts).
//
// The packages are one fixed group, and the adapters have `core` as a peer. By
// default changesets bumps a peer dependent to major whenever the peer changes, and the
// fixed group then takes every package to major. `onlyUpdatePeerDependentsWhenOutOfRange`
// and the tilde peer ranges (`~0.1.0`, as DOCX Editor) keep 0.x minors as minors.
import { afterEach, expect, test } from 'bun:test';
import getReleasePlanModule from '@changesets/get-release-plan';
import {
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

test('the pending changesets release no major version', async () => {
  const { directory } = fixture();
  for (const file of readdirSync(join(root, '.changeset')))
    if (file.endsWith('.md') && file !== 'README.md')
      writeFileSync(
        join(directory, '.changeset', file),
        readFileSync(join(root, '.changeset', file))
      );
  const releases = await plan(directory);
  expect(releases.some((release) => release.type === 'major')).toBe(false);
  expect(new Set(releases.map((release) => release.newVersion)).size).toBeLessThanOrEqual(1);
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
