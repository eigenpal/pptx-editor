import { expect, test } from 'bun:test';
import { cpSync, mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { resolve } from 'node:path';
import { verifyFiles, verifyReleaseTable } from './catalog.mjs';
import { hash, read, root, write } from './common.mjs';

const entry = read(resolve(root, '.collaboration/releases.json')).releases[0];
test('captured release verifies, but missing or modified evidence fails', () => {
  expect(() => verifyFiles(entry)).not.toThrow();
  expect(() => verifyFiles({ ...entry, directory: '../outside' })).toThrow();
  expect(() => verifyFiles({ ...entry, hashes: {} })).toThrow();
  expect(() =>
    verifyFiles({ ...entry, hashes: { ...entry.hashes, 'fixture.json': 'changed' } })
  ).toThrow();
});

test('an altered installation fails even when its file hash is updated', () => {
  const temp = mkdtempSync(resolve(tmpdir(), 'pptx-catalog-test-'));
  try {
    const directory = resolve(temp, entry.directory);
    cpSync(resolve(root, entry.directory), directory, { recursive: true });
    const file = resolve(directory, 'package-lock.json');
    const lock = read(file);
    lock.packages['node_modules/@pptx-editor.dev/pro'].integrity = 'different-package';
    write(file, lock);
    const changed = { ...entry, hashes: { ...entry.hashes, 'package-lock.json': hash(file) } };
    expect(() => verifyFiles(changed, temp)).toThrow();
  } finally {
    rmSync(temp, { recursive: true, force: true });
  }
});

test('published release documentation must match the immutable catalog', () => {
  const catalog = { releases: [entry] };
  const page = `{/* collaboration-releases:start */}\n| \`${entry.version}\` | \`${entry.format}\` |\n{/* collaboration-releases:end */}`;
  expect(() => verifyReleaseTable(catalog, page)).not.toThrow();
  expect(() => verifyReleaseTable(catalog, page.replace(entry.format, 'stale-format'))).toThrow();
  expect(() => verifyReleaseTable(catalog, '')).toThrow();
});
