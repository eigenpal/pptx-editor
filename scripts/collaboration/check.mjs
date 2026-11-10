// Copyright (c) 2026 Eigenpal, Inc. Licensed under Apache-2.0.
// Compatibility decisions follow DOCX Editor policy.mjs, commit 13f08556889ef7c192023f92aa865cc9cc8e3fe4.
import assert from 'node:assert/strict';
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { resolve } from 'node:path';
import { read, root, run } from './common.mjs';
import { verifyFiles, verifyReleaseTable } from './catalog.mjs';
import { sourceFormat, validateDecision, repositoryPath } from './policy.mjs';
const arg = process.argv.indexOf('--base');
if (arg < 0 || !process.argv[arg + 1])
  throw new Error('Usage: collaboration:check --base <git-ref>');
const base = run('git', ['rev-parse', '--verify', `${process.argv[arg + 1]}^{commit}`]);
const changed = run('git', ['diff', '--name-only', base]).split('\n');
const untracked = run('git', ['ls-files', '--others', '--exclude-standard']).split('\n');
const paths = [...new Set([...changed, ...untracked])];
const oldPaths = new Set(
  run('git', ['ls-tree', '-r', '--name-only', base, '.collaboration']).split('\n')
);
for (const path of paths) {
  if (oldPaths.has(path) && /^\.collaboration\/(releases\/|changes\/)/.test(path))
    throw new Error(`Immutable collaboration evidence changed: ${path}`);
}
const catalog = read(resolve(root, '.collaboration/releases.json'));
verifyReleaseTable(
  catalog,
  readFileSync(resolve(root, 'docs/site/content/pro/collaboration-versions.mdx'), 'utf8')
);
for (const entry of catalog.releases) verifyFiles(entry);
if (oldPaths.has('.collaboration/releases.json')) {
  const previous = JSON.parse(run('git', ['show', `${base}:.collaboration/releases.json`]));
  assert.equal(catalog.baseline, previous.baseline);
  assert.deepEqual(catalog.releases.slice(0, previous.releases.length), previous.releases);
}
const implementation = paths.some(
  (path) => /^packages\/(pro|core)\/src\//.test(path) && !/(__tests__|\.test\.)/.test(path)
);
const decisions = paths.filter(
  (path) => /^\.collaboration\/changes\/[^/]+\.json$/.test(path) && !oldPaths.has(path)
);
if (implementation && !decisions.length)
  throw new Error('Core or Pro changes require a collaboration impact decision');

const formatPath = 'packages/pro/src/collaboration/yjs/format.ts';
const schemaPath = 'packages/pro/src/collaboration/yjs/schema.ts';
const current = sourceFormat(
  readFileSync(resolve(root, formatPath), 'utf8'),
  readFileSync(resolve(root, schemaPath), 'utf8')
);
const previous = sourceFormat(
  run('git', ['show', `${base}:${formatPath}`]),
  run('git', ['show', `${base}:${schemaPath}`])
);
for (const path of decisions) {
  const record = read(resolve(root, path));
  validateDecision(record, previous, current);
  for (const test of record.tests) {
    repositoryPath(test);
    assert.match(test, /(?:^|\/|\.)(?:test|spec)\.(?:[cm]?[jt]sx?)$/);
    assert.ok(existsSync(resolve(root, test)), `Missing evidence: ${test}`);
  }
  if (record.before !== record.after) {
    repositoryPath(record.migration);
    assert.match(record.migration, /^docs\/site\/content\/.+\.mdx$/);
    assert.ok(existsSync(resolve(root, record.migration)));
    assert.match(record.changeset, /^\.changeset\/[^/]+\.md$/);
    const note = readFileSync(resolve(root, record.changeset), 'utf8');
    const frontmatter = note.match(/^---\r?\n([\s\S]*?)\r?\n---/);
    assert.ok(frontmatter, 'Missing changeset frontmatter');
    assert.match(frontmatter[1], /^['"]?@pptx-editor.dev\/pro['"]?: (minor|major)\s*$/m);
    assert.match(note, /Breaking collaboration upgrade/);
    assert.ok(
      note.includes(record.migration.replace('docs/site/content/', '/docs/').replace(/\.mdx$/, '')),
      'Release note must link to the migration guide'
    );
  }
}
const all = readdirSync(resolve(root, '.collaboration/changes')).filter((file) =>
  file.endsWith('.json')
);
assert.ok(
  all.some((file) => read(resolve(root, '.collaboration/changes', file)).after === current.format),
  'The source format needs a recorded compatibility decision'
);
if (previous.format !== current.format)
  assert.ok(decisions.length, 'A format change requires a new migration decision');
console.log(`Collaboration policy passed: ${current.format}`);
