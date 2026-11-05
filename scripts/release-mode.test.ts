// After DOCX Editor's scripts/release-mode.test.ts (commit b9a04a30f): the canary cases are
// the same; the stable cases cover this repository's pending-or-publish rule.
import { afterEach, expect, test } from 'bun:test';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import {
  canaryMode,
  canarySuffix,
  distTags,
  pendingReleases,
  publicPackages,
  releaseMode,
  stableReleaseClear,
} from './release-mode.mjs';

const directories: string[] = [];
afterEach(() => {
  for (const directory of directories.splice(0))
    rmSync(directory, { recursive: true, force: true });
});

function fixture(changesets: string[] = [], fixed = ['@scope/core', '@scope/react']): string {
  const root = mkdtempSync(join(tmpdir(), 'pptx-release-mode-'));
  directories.push(root);
  const write = (path: string, value: unknown): void => {
    mkdirSync(dirname(join(root, path)), { recursive: true });
    writeFileSync(join(root, path), typeof value === 'string' ? value : JSON.stringify(value));
  };
  write('packages/core/package.json', { name: '@scope/core', version: '1.2.0' });
  write('packages/react/package.json', { name: '@scope/react', version: '1.2.0' });
  write('packages/internal/package.json', {
    name: '@scope/internal',
    version: '1.2.0',
    private: true,
  });
  write('.changeset/README.md', 'Not a changeset.');
  write('.changeset/config.json', { fixed: [fixed] });
  for (const name of changesets) write(`.changeset/${name}`, '---\n"@scope/core": patch\n---\n');
  return root;
}

test('lists the public members of the fixed group', () => {
  expect(publicPackages(fixture([], ['@scope/core', '@scope/react', '@scope/internal']))).toEqual([
    { name: '@scope/core', version: '1.2.0' },
    { name: '@scope/react', version: '1.2.0' },
  ]);
});

test('pending changesets update the release PR; none take the publish path', () => {
  expect(releaseMode(fixture(['brave-fox.md']))).toEqual({ mode: 'pending' });
  expect(releaseMode(fixture())).toEqual({ mode: 'publish' });
});

test('a group member without a manifest fails', () => {
  const root = fixture([], ['@scope/core', '@scope/react', '@scope/new']);
  expect(() => publicPackages(root)).toThrow('No packages/*/package.json is named @scope/new');
});

test('manifests resolve by package name, not by directory name', () => {
  const root = fixture([], ['@scope/core', '@scope/pdf-tools']);
  mkdirSync(join(root, 'packages/pdf'), { recursive: true });
  writeFileSync(
    join(root, 'packages/pdf/package.json'),
    JSON.stringify({ name: '@scope/pdf-tools', version: '1.2.0' })
  );
  expect(publicPackages(root)).toEqual([
    { name: '@scope/core', version: '1.2.0' },
    { name: '@scope/pdf-tools', version: '1.2.0' },
  ]);
});

test('the repository resolves the published package set', () => {
  const names = publicPackages(join(import.meta.dir, '..')).map(({ name }) => name);
  expect(names).toContain('@pptx-editor.dev/core');
  expect(names).toContain('@pptx-editor.dev/pptx-to-pdf');
  expect(names.every((name) => name.startsWith('@pptx-editor.dev/'))).toBe(true);
});

const commit = 'b20ba9fcd2e0a1b5c3d4e5f60718293a4b5c6d7e';
const atCommit = `1.2.1-canary.20261103032300${canarySuffix(commit)}`;
const tagsOf =
  (canary: Record<string, string>, failing: string[] = []) =>
  async (name: string): Promise<Record<string, string> | null> => {
    if (failing.includes(name)) return null;
    return { latest: '1.2.0', ...(canary[name] ? { canary: canary[name] } : {}) };
  };

test('a canary needs pending changesets, so a fresh stable release is idle', async () => {
  const tags = (): Promise<null> => {
    throw new Error('registry must not be called');
  };
  expect(await canaryMode(fixture(), commit, tags)).toEqual({
    mode: 'idle',
    reason: 'no pending changesets release a package',
  });
});

test('an empty changeset releases nothing, so the canary stays idle', async () => {
  const root = fixture();
  writeFileSync(join(root, '.changeset/empty-owl.md'), '---\n---\n\nDocs only.\n');
  expect(pendingReleases(root)).toEqual([]);
  expect((await canaryMode(root, commit, tagsOf({}))).mode).toBe('idle');
  writeFileSync(join(root, '.changeset/quoted.md'), "---\n'@scope/react': minor\n---\n");
  expect(pendingReleases(root)).toEqual(['quoted.md']);
});

test('a changeset that bumps only a private package releases nothing public', async () => {
  const root = fixture();
  writeFileSync(join(root, '.changeset/private-owl.md'), "---\n'@scope/internal': patch\n---\n");
  expect(pendingReleases(root)).toEqual([]);
  expect((await canaryMode(root, commit, tagsOf({}))).mode).toBe('idle');
  writeFileSync(
    join(root, '.changeset/mixed.md'),
    "---\n'@scope/internal': patch\n'@scope/react': patch\n---\n"
  );
  expect(pendingReleases(root)).toEqual(['mixed.md']);
});

test('a package without a latest version blocks the canary', async () => {
  // npm would point `latest` at the first version of a new package.
  const tags = async (name: string): Promise<Record<string, string>> =>
    name === '@scope/react' ? {} : { latest: '1.2.0' };
  expect(await canaryMode(fixture(['brave-fox.md']), commit, tags)).toEqual({
    mode: 'idle',
    reason: '@scope/react has no latest version yet',
  });
});

test('a name-reserving placeholder on latest lets the first canary publish', async () => {
  const tags = async (): Promise<Record<string, string>> => ({
    latest: '0.0.1-placeholder',
    placeholder: '0.0.1-placeholder',
  });
  expect((await canaryMode(fixture(['brave-fox.md']), commit, tags)).mode).toBe('canary');
});

test('a new green commit publishes a canary', async () => {
  const tags = tagsOf({
    '@scope/core': '1.2.1-canary.20261102032300.sha-0000000',
    '@scope/react': '1.2.1-canary.20261102032300.sha-0000000',
  });
  expect((await canaryMode(fixture(['brave-fox.md']), commit, tags)).mode).toBe('canary');
});

test('the first canary publishes when no package has the tag yet', async () => {
  expect((await canaryMode(fixture(['brave-fox.md']), commit, tagsOf({}))).mode).toBe('canary');
});

test('a canary already on every package is idle', async () => {
  const tags = tagsOf({ '@scope/core': atCommit, '@scope/react': atCommit });
  expect(await canaryMode(fixture(['brave-fox.md']), commit, tags)).toEqual({
    mode: 'idle',
    reason: `canary already holds ${commit.slice(0, 7)}`,
  });
});

test('a partial publication publishes the whole set again', async () => {
  const tags = tagsOf({ '@scope/core': atCommit });
  expect((await canaryMode(fixture(['brave-fox.md']), commit, tags)).mode).toBe('canary');
});

test('a dist-tag lookup failure skips the night rather than guessing', async () => {
  const tags = tagsOf({}, ['@scope/react']);
  expect(await canaryMode(fixture(['brave-fox.md']), commit, tags)).toEqual({
    mode: 'idle',
    reason: 'dist-tag lookup failed for @scope/react',
  });
});

interface Run {
  id: number;
  event: string;
  status: string;
}
function github(changesets: string[], runs: Run[], failing = '') {
  const calls: string[] = [];
  const request = async (url: string, init: { headers: Record<string, string> }) => {
    calls.push(url);
    expect(init.headers.authorization).toBe('Bearer token');
    if (failing && url.includes(failing)) return new Response('{}', { status: 502 });
    if (url.endsWith('/contents/.changeset?ref=main'))
      return Response.json(['README.md', 'config.json', ...changesets].map((name) => ({ name })));
    if (url.endsWith('/actions/workflows/release.yml/runs?per_page=50'))
      return Response.json({ workflow_runs: runs });
    throw new Error(`unexpected ${url}`);
  };
  const clear = () =>
    stableReleaseClear({
      repository: 'owner/repo',
      runId: '7',
      token: 'token',
      fetch: request as unknown as typeof fetch,
    });
  return { clear, calls };
}

test('the stable-release gate is clear only with pending changesets and no other active run', async () => {
  const done = [
    { id: 7, event: 'schedule', status: 'in_progress' },
    { id: 6, event: 'push', status: 'completed' },
    { id: 5, event: 'schedule', status: 'queued' },
  ];
  const { clear, calls } = github(['brave-fox.md'], done);
  expect(await clear()).toEqual({ clear: true, reason: 'no stable release is active' });
  expect(calls[0]).toBe('https://api.github.com/repos/owner/repo/contents/.changeset?ref=main');
});

test('a merged release PR closes the gate', async () => {
  expect((await github([], []).clear()).clear).toBe(false);
});

test('any other active push or manual Release run closes the gate', async () => {
  for (const event of ['push', 'workflow_dispatch'])
    for (const status of ['queued', 'in_progress', 'waiting', 'pending', 'requested']) {
      const { clear } = github(['brave-fox.md'], [{ id: 8, event, status }]);
      expect(await clear()).toEqual({ clear: false, reason: 'Release run 8 is active' });
    }
});

test('a GitHub API failure closes the gate', async () => {
  for (const failing of ['contents', 'runs']) {
    const result = await github(['brave-fox.md'], [], failing).clear();
    expect(result.clear).toBe(false);
    expect(result.reason).toContain('HTTP 502');
  }
});

test('a canary refuses a missing or short commit hash', async () => {
  for (const value of [undefined, '', 'b20ba9f', 'main']) {
    await expect(
      canaryMode(fixture(['brave-fox.md']), value as string, tagsOf({}))
    ).rejects.toThrow('The canary needs a full commit hash');
  }
});

test('dist-tag lookups answer {} for a new package (404 or 401) and null on failures', async () => {
  const notFound = async () => new Response('{}', { status: 404 });
  const unknownScoped = async () => new Response('{}', { status: 401 });
  const unavailable = async () => new Response('{}', { status: 503 });
  const offline = async (): Promise<Response> => {
    throw new Error('ECONNRESET');
  };
  const found = async (url: string) => {
    expect(url).toBe('https://registry.npmjs.org/-/package/%40scope%2Fcore/dist-tags');
    return Response.json({ latest: '1.2.0', canary: atCommit });
  };
  const as = (request: unknown) => ({ fetch: request as typeof fetch });
  expect(await distTags('@scope/core', as(unavailable))).toBeNull();
  expect(await distTags('@scope/core', as(notFound))).toEqual({});
  expect(await distTags('@scope/core', as(unknownScoped))).toEqual({});
  expect(await distTags('@scope/core', as(offline))).toBeNull();
  expect(await distTags('@scope/core', as(found))).toEqual({ latest: '1.2.0', canary: atCommit });
});

test('the snapshot template matches the canary suffix and keeps versions valid', () => {
  const config = JSON.parse(
    readFileSync(join(import.meta.dir, '..', '.changeset/config.json'), 'utf8')
  ) as { snapshot: unknown };
  expect(config.snapshot).toEqual({
    useCalculatedVersion: true,
    prereleaseTemplate: '{tag}.{datetime}.sha-{commit-short}',
  });
  // An all-digit short hash with a leading zero is not a valid numeric identifier.
  // The prefix keeps that identifier alphanumeric.
  const digits = `0.1.0-canary.20261103032300${canarySuffix('0123456'.padEnd(40, '0'))}`;
  expect(Bun.semver.satisfies(digits, digits)).toBe(true);
  // Later canaries sort higher, whatever their hashes.
  expect(
    Bun.semver.order(
      '0.1.0-canary.20261103032300.sha-fffffff',
      '0.1.0-canary.20261104032300.sha-0000000'
    )
  ).toBe(-1);
  // A normal caret range never installs a canary.
  expect(Bun.semver.satisfies('1.3.0-canary.20261103032300.sha-0123456', '^1.2.0')).toBe(false);
});

test('no published package lists a workspace range that npm would ship as it is', () => {
  // `npm publish` copies `workspace:` ranges into the tarball, and a consumer install then
  // fails. Peer ranges are pinned by changesets; runtime ranges must be real ranges.
  const root = join(import.meta.dir, '..');
  const offenders = publicPackages(root).flatMap(({ name }) => {
    const directory = name.replace('@pptx-editor.dev/', '');
    const manifest = JSON.parse(
      readFileSync(join(root, 'packages', directory, 'package.json'), 'utf8')
    ) as Record<string, Record<string, string> | undefined>;
    return ['dependencies', 'optionalDependencies', 'peerDependencies'].flatMap((field) =>
      Object.entries(manifest[field] ?? {})
        .filter(([, range]) => range.startsWith('workspace:'))
        .map(([dependency]) => `${name} ${field} ${dependency}`)
    );
  });
  expect(offenders).toEqual([]);
});
