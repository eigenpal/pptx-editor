// Copyright (c) 2026 Eigenpal, Inc. Licensed under Apache-2.0.
import { afterEach, expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { REQUIRED } from './catalog-merge.mjs';

const directories = [];
afterEach(() => {
  for (const directory of directories.splice(0))
    rmSync(directory, { recursive: true, force: true });
});
const head = 'a'.repeat(40);
function invoke(mode, expected = head, pending = false) {
  const directory = mkdtempSync(join(tmpdir(), 'pptx-baseline-merge-'));
  directories.push(directory);
  const pr = {
    state: 'OPEN',
    isDraft: false,
    isCrossRepository: false,
    headRepository: { nameWithOwner: 'eigenpal/pptx-editor' },
    baseRefName: 'main',
    headRefName: 'automation/collaboration-baseline-0.5.0',
    headRefOid: head,
    author: { login: 'app/eigenpal-release-pal' },
    mergeable: 'MERGEABLE',
    files: [
      '.collaboration/releases.json',
      'docs/site/content/pro/collaboration-versions.mdx',
      ...['fixture.json', 'package.json', 'package-lock.json'].map(
        (file) => `.collaboration/releases/0.5.0/${file}`
      ),
    ].map((path) => ({ path })),
    statusCheckRollup: REQUIRED.map((name) => ({
      __typename: 'CheckRun',
      name,
      status: 'COMPLETED',
      conclusion: 'SUCCESS',
    })),
  };
  const runs = {
    workflow_runs: [
      {
        id: 1,
        path: '.github/workflows/ci.yml',
        event: 'pull_request',
        head_sha: head,
        status: pending ? 'in_progress' : 'completed',
        conclusion: pending ? null : 'success',
      },
    ],
  };
  writeFileSync(join(directory, 'responses.json'), JSON.stringify({ pr, runs }));
  writeFileSync(
    join(directory, 'gh'),
    `#!/usr/bin/env node
const { readFileSync } = require('node:fs');
const data = JSON.parse(readFileSync(process.env.MERGE_TEST_RESPONSES, 'utf8'));
process.stdout.write(JSON.stringify(process.argv[2] === 'pr' ? data.pr : data.runs));
`,
    { mode: 0o755 }
  );
  const output = join(directory, 'output');
  writeFileSync(output, '');
  const env = {
    ...process.env,
    PATH: `${directory}:${process.env.PATH}`,
    GITHUB_REPOSITORY: 'eigenpal/pptx-editor',
    PR_NUMBER: '126',
    RELEASE_VERSION: '0.5.0',
    GITHUB_OUTPUT: output,
    MERGE_TEST_RESPONSES: join(directory, 'responses.json'),
  };
  delete env.EXPECTED_HEAD;
  if (expected !== undefined) env.EXPECTED_HEAD = expected;
  const result = spawnSync(process.execPath, ['scripts/collaboration/catalog-merge.mjs', mode], {
    env,
    encoding: 'utf8',
    timeout: 10_000,
  });
  return { ...result, output: readFileSync(output, 'utf8') };
}

test('wait exports the verified head for the approval and merge steps', () => {
  const result = invoke('wait');
  expect(result.status).toBe(0);
  expect(result.output).toBe(`head=${head}\n`);
});

test('check refuses a different head and never exports it', () => {
  const result = invoke('check', 'b'.repeat(40));
  expect(result.status).toBe(1);
  expect(result.stderr).toContain('changed while waiting');
  expect(result.output).toBe('');
});

test('check fails immediately when CI becomes pending', () => {
  const result = invoke('check', head, true);
  expect(result.status).toBe(1);
  expect(result.stderr).toContain('checks are not complete');
  expect(result.output).toBe('');
});

test('check requires an explicit valid tested commit', () => {
  const result = invoke('check', '');
  expect(result.status).toBe(1);
  expect(result.stderr).toContain('Invalid catalog merge inputs');
  expect(result.output).toBe('');
});
