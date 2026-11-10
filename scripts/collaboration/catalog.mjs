// Copyright (c) 2026 Eigenpal, Inc. Licensed under Apache-2.0.
// Immutable npm baselines follow DOCX Editor catalog.mjs, commit 13f08556889ef7c192023f92aa865cc9cc8e3fe4.
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { resolve } from 'node:path';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import {
  baseline,
  compare,
  hash,
  install,
  names,
  peer,
  read,
  registry,
  root,
  run,
  stable,
  write,
} from './common.mjs';
const catalogFile = resolve(root, '.collaboration/releases.json');
const files = ['package.json', 'package-lock.json', 'fixture.json'];
export function verifyFiles(entry, repository = root) {
  if (!stable(entry.version) || entry.directory !== `.collaboration/releases/${entry.version}`)
    throw new Error('Invalid baseline directory');
  verifyContents(entry, resolve(repository, entry.directory));
}
function verifyContents(entry, directory) {
  assert.deepEqual(Object.keys(entry.hashes).sort(), [...files].sort());
  for (const file of files) assert.equal(hash(resolve(directory, file)), entry.hashes[file], file);
  const lock = read(resolve(directory, 'package-lock.json'));
  for (const name of names) {
    const pkg = lock.packages[`node_modules/${name}`];
    assert.equal(pkg.version, entry.version);
    assert.equal(pkg.integrity, entry.packages[name].integrity);
    assert.equal(pkg.resolved, entry.packages[name].tarball);
  }
}
export function verifyReleaseTable(catalog, source) {
  const section = source.match(
    /{\/\* collaboration-releases:start \*\/}([\s\S]*?){\/\* collaboration-releases:end \*\/}/
  );
  assert.ok(section, 'Missing release catalog documentation markers');
  const rows = [...section[1].matchAll(/\|\s*`([^`]+)`\s*\|\s*`([^`]+)`\s*\|/g)].map((match) => [
    match[1],
    match[2],
  ]);
  assert.deepEqual(
    rows,
    catalog.releases.map((entry) => [entry.version, entry.format]),
    'The documented release table differs from the immutable catalog'
  );
}
export function verifyCatalog() {
  const catalog = read(catalogFile);
  assert.equal(catalog.baseline, baseline);
  verifyReleaseTable(
    catalog,
    readFileSync(resolve(root, 'docs/site/content/pro/collaboration-versions.mdx'), 'utf8')
  );
  const versions = JSON.parse(
    run('npm', ['view', '@pptx-editor.dev/pro', 'versions', '--json'])
  ).filter((version) => stable(version) && compare(version, baseline) >= 0);
  assert.deepEqual(
    catalog.releases.map((entry) => entry.version),
    versions.sort(compare),
    'Every stable release from the baseline must have a captured fixture'
  );
  for (const entry of catalog.releases) {
    verifyFiles(entry);
    assert.match(entry.format, /^pptx-collaboration(?:-\d+|:\d+\.\d+\.\d+\.\d+)$/);
    const source = run('git', ['show', `${entry.commit}:bun.lock`], root, false);
    assert.equal(createHash('sha256').update(source).digest('hex'), entry.sourceLockHash);

    for (const name of names) {
      const metadata = registry(name, entry.version);
      assert.equal(metadata.dist.integrity, entry.packages[name].integrity);
      assert.equal(metadata.dist.tarball, entry.packages[name].tarball);
      assert.equal(metadata.gitHead, entry.commit);
    }
  }
  return catalog;
}
export async function capture(version) {
  if (!stable(version) || compare(version, baseline) < 0) throw new Error('Invalid stable release');
  const catalog = read(catalogFile);
  const directory = `.collaboration/releases/${version}`;
  const destination = resolve(root, directory);
  const existing = catalog.releases.find((entry) => entry.version === version);
  if (existing) {
    verifyFiles(existing);
    console.log(`Baseline ${version} already captured`);
    return;
  }
  if (existsSync(destination)) throw new Error('Refusing to overwrite baseline files');
  if (catalog.releases.some((entry) => compare(entry.version, version) >= 0))
    throw new Error('Capture releases in chronological order');
  const commit = run('git', ['rev-parse', `v${version}^{commit}`]);
  const sourceLock = run('git', ['show', `v${version}:bun.lock`], root, false);
  const packages = {};
  for (const name of names) {
    const metadata = registry(name, version);
    assert.equal(metadata.version, version);
    assert.equal(metadata.gitHead, commit, `${name} source commit differs`);
    packages[name] = { tarball: metadata.dist.tarball, integrity: metadata.dist.integrity };
  }
  const directoryTemp = mkdtempSync(resolve(tmpdir(), 'pptx-baseline-'));
  let worker;
  try {
    const dependencies = Object.fromEntries(names.map((name) => [name, version]));
    for (const name of ['yjs', 'y-protocols']) {
      const pinned = sourceLock.match(new RegExp(`"${name}": \\["${name}@([^"\\s]+)`))?.[1];
      if (!pinned) throw new Error(`Missing source lock pin: ${name}`);
      dependencies[name] = pinned;
    }
    write(resolve(directoryTemp, 'package.json'), {
      name: 'pptx-collaboration-baseline',
      private: true,
      type: 'module',
      dependencies,
    });
    install(directoryTemp);
    worker = peer(directoryTemp);
    const document = readFileSync(resolve(root, '.collaboration/fixtures/content.pptx')).toString(
      'base64'
    );
    const expected = await worker.call({ op: 'open', document });
    const { format } = await worker.call({ op: 'info' });
    assert.equal(typeof format, 'string');
    const edited = await worker.call({ op: 'edit', text: 'history' });
    assert.notDeepEqual(edited, expected);
    assert.equal((await worker.call({ op: 'history' })).undo, true);
    await worker.call({ op: 'undo' });
    assert.deepEqual(await worker.call({ op: 'snapshot' }), expected);
    const state = await worker.call({ op: 'state' });
    write(resolve(directoryTemp, 'fixture.json'), { document, state, expected });
    const entry = {
      version,
      commit,
      format,
      directory,
      packages,
      sourceLockHash: createHash('sha256').update(sourceLock).digest('hex'),
      hashes: Object.fromEntries(files.map((file) => [file, hash(resolve(directoryTemp, file))])),
    };
    verifyContents(entry, directoryTemp);
    const previousCatalog = readFileSync(catalogFile);
    catalog.releases.push(entry);
    const page = resolve(root, 'docs/site/content/pro/collaboration-versions.mdx');
    const table = [
      '| Package release | Collaboration format |',
      '| --- | --- |',
      ...catalog.releases.map((release) => `| \`${release.version}\` | \`${release.format}\` |`),
    ].join('\n');
    const previousPage = readFileSync(page, 'utf8');
    const markers =
      /({\/\* collaboration-releases:start \*\/})[\s\S]*?({\/\* collaboration-releases:end \*\/})/;
    assert.match(previousPage, markers, 'Missing release catalog documentation markers');
    const nextPage = previousPage.replace(markers, `$1\n\n${table}\n\n$2`);
    // Validate every input before creating immutable evidence. Roll back a failed local write.
    mkdirSync(destination);
    try {
      for (const file of files)
        copyFileSync(resolve(directoryTemp, file), resolve(destination, file));
      verifyFiles(entry);
      write(catalogFile, catalog);
      writeFileSync(page, nextPage);
    } catch (error) {
      writeFileSync(catalogFile, previousCatalog);
      writeFileSync(page, previousPage);
      rmSync(destination, { recursive: true, force: true });
      throw error;
    }

    console.log(`Captured ${version}: ${format}`);
  } finally {
    worker?.close();
    rmSync(directoryTemp, { recursive: true, force: true });
  }
}
if (process.argv[1] === import.meta.filename) {
  const version = process.argv[process.argv.indexOf('--capture') + 1];
  if (process.argv.includes('--capture')) await capture(version);
  else {
    verifyCatalog();
    console.log('Published collaboration catalog verified');
  }
}
