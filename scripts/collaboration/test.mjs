// Copyright (c) 2026 Eigenpal, Inc. Licensed under Apache-2.0.
// Isolated release matrix follows DOCX Editor test.mjs and recovery.mjs,
// commit 13f08556889ef7c192023f92aa865cc9cc8e3fe4.
import assert from 'node:assert/strict';
import { copyFileSync, mkdirSync, mkdtempSync, rmSync } from 'node:fs';
import { resolve } from 'node:path';
import { tmpdir } from 'node:os';
import { install, names, peer, read, root, run, write } from './common.mjs';
import { verifyCatalog } from './catalog.mjs';
const catalog = verifyCatalog();
const temp = mkdtempSync(resolve(tmpdir(), 'pptx-compatibility-'));
const workers = [];
function start(directory) {
  const worker = peer(directory);
  workers.push(worker);
  return worker;
}
async function sync(peers) {
  // Deliver actual incremental updates in reverse order, twice, with delayed peers.
  let settled = false;
  for (let round = 0; round < 16; round++) {
    const batches = [];
    for (const worker of peers) batches.push(await worker.call({ op: 'updates' }));
    if (batches.every((batch) => batch.length === 0)) {
      settled = true;
      break;
    }
    for (const worker of [...peers].reverse())
      for (const batch of [...batches].reverse())
        for (const state of [...batch].reverse()) {
          await worker.call({ op: 'apply', state });
          await worker.call({ op: 'apply', state });
        }
  }
  assert.ok(settled, 'Incremental repair updates did not settle');
  const expected = await peers[0].call({ op: 'snapshot' });
  for (const worker of peers) assert.deepEqual(await worker.call({ op: 'snapshot' }), expected);
  return expected;
}
async function scenario(directories, fixture) {
  const peers = directories.map(start);
  try {
    const initial = await peers[0].call({
      op: 'open',
      document: fixture.document,
      state: fixture.state,
    });
    const state = await peers[0].call({ op: 'state' });
    for (const worker of peers.slice(1))
      await worker.call({ op: 'open', document: fixture.document, state });
    for (const [index, worker] of peers.entries())
      await worker.call({ op: 'edit', text: String(index) });
    const edited = await sync(peers);
    for (const digit of ['0', '1', '2'])
      assert.equal(edited.paragraphs.join('').split(digit).length - 1, 1);
    assert.equal(edited.paragraphs.join('').replace(/[012]/g, ''), initial.paragraphs.join(''));
    await peers[1].call({ op: 'undo' });
    const undone = await sync(peers);
    assert.equal(undone.paragraphs.join(''), edited.paragraphs.join('').replace('1', ''));
    await peers[1].call({ op: 'redo' });
    assert.deepEqual(await sync(peers), edited);
    const saved = await peers[0].call({ op: 'state' });
    assert.deepEqual(
      await peers[2].call({ op: 'open', document: fixture.document, state: saved }),
      edited
    );
    assert.deepEqual(await peers[2].call({ op: 'history' }), { undo: false, redo: false });
    await peers[2].call({ op: 'edit', text: 'reconnected' });
    const reconnected = await sync(peers);
    assert.equal(
      reconnected.paragraphs.join('').replace('reconnected', ''),
      edited.paragraphs.join('')
    );
  } finally {
    peers.forEach((worker) => worker.close());
  }
}
try {
  const candidate = resolve(temp, 'candidate');
  const packs = resolve(temp, 'packs');
  mkdirSync(candidate);
  mkdirSync(packs);
  const dependencies = {};
  for (const name of names) {
    const directory = resolve(root, 'packages', name.split('/')[1]);
    const [packed] = JSON.parse(
      run('npm', ['pack', '--json', '--ignore-scripts', '--pack-destination', packs], directory)
    );
    dependencies[name] = `file:${resolve(packs, packed.filename)}`;
  }
  write(resolve(candidate, 'package.json'), {
    name: 'pptx-candidate',
    private: true,
    type: 'module',
    dependencies,
  });
  install(candidate);
  const current = start(candidate);
  const { format } = await current.call({ op: 'info' });
  const first = read(resolve(root, catalog.releases[0].directory, 'fixture.json'));
  await scenario([candidate, candidate, candidate], { document: first.document });
  for (const entry of catalog.releases) {
    const old = resolve(temp, entry.version);
    mkdirSync(old);
    for (const file of ['package.json', 'package-lock.json'])
      copyFileSync(resolve(root, entry.directory, file), resolve(old, file));
    install(old, true);
    const released = start(old);
    const fixture = read(resolve(root, entry.directory, 'fixture.json'));
    assert.deepEqual(await released.call({ op: 'open', ...fixture }), fixture.expected);
    assert.deepEqual(await released.call({ op: 'info' }), { format: entry.format });
    const inspection = await current.call({ op: 'inspect', state: fixture.state });
    assert.deepEqual(inspection, { format: entry.format, rejected: entry.format !== format });
    if (entry.format === format) {
      await scenario([old, candidate, candidate], fixture);
      await scenario([candidate, old, old], { document: fixture.document });
    } else {
      assert.deepEqual(
        await current.call({ op: 'probe', document: fixture.document, state: fixture.state }),
        { status: 'error', code: 'schema-version-mismatch', unchanged: true }
      );
      // Recover through the package that owns the saved room, then seed a separate room.
      const document = await released.call({ op: 'save' });
      const migrated = await current.call({ op: 'open', document });
      assert.deepEqual(migrated, fixture.expected);
      assert.deepEqual(await current.call({ op: 'history' }), { undo: false, redo: false });
      await current.call({ op: 'edit', text: 'migrated' });
      assert.ok((await current.call({ op: 'snapshot' })).paragraphs.join('').includes('migrated'));
      const state = await current.call({ op: 'state' });
      assert.equal((await released.call({ op: 'inspect', state })).rejected, true);
      assert.deepEqual(await released.call({ op: 'probe', document: fixture.document, state }), {
        status: 'error',
        code: 'schema-version-mismatch',
        unchanged: true,
      });
    }
    console.log(
      `${entry.version}: saved room, admission, and ${entry.format === format ? 'mixed clients' : 'migration'} passed`
    );
  }
  console.log('Candidate three-peer convergence and release matrix passed');
} finally {
  workers.forEach((worker) => worker.close());
  rmSync(temp, { recursive: true, force: true });
}
