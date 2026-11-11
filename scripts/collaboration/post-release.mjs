// Copyright (c) 2026 Eigenpal, Inc. Licensed under Apache-2.0.
// Capture through an App-authenticated PR, following DOCX Editor post-release.mjs.
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { resolve } from 'node:path';
import { capture, verifyCatalog } from './catalog.mjs';
import { run, stable } from './common.mjs';
import { waitForPublishedPackages } from './registry-ready.mjs';
const version = process.env.RELEASE_VERSION;
if (!version || !stable(version)) throw new Error('RELEASE_VERSION must be a stable version');
await waitForPublishedPackages(version);
const branch = `automation/collaboration-baseline-${version}`;
const remote = run('git', ['ls-remote', '--heads', 'origin', branch]);
if (remote) {
  run('git', ['fetch', 'origin', branch]);
  run('git', ['checkout', '-B', branch, 'FETCH_HEAD']);
} else run('git', ['checkout', '-b', branch]);
await capture(version);
verifyCatalog();
run('git', [
  'add',
  '.collaboration/releases.json',
  `.collaboration/releases/${version}`,
  'docs/site/content/pro/collaboration-versions.mdx',
]);
if (run('git', ['diff', '--cached', '--name-only'])) {
  run('git', ['commit', '-m', `Capture collaboration baseline ${version}`]);
  run('git', ['push', '--set-upstream', 'origin', branch]);
}
const prs = JSON.parse(
  run('gh', ['pr', 'list', '--head', branch, '--state', 'all', '--json', 'number,state'])
);
if (prs.some((pr) => pr.state === 'OPEN' || pr.state === 'MERGED')) {
  console.log('Baseline pull request already exists');
} else {
  const temp = mkdtempSync(resolve(tmpdir(), 'pptx-baseline-pr-'));
  try {
    const file = resolve(temp, 'body.md');
    writeFileSync(
      file,
      `Capture the published ${version} packages and saved collaboration room.\n\nRegistry integrity, source commit, installation lock, and fixture hashes passed verification.\nReview and merge after CI passes. The packages are already published; do not republish.\n`
    );
    run('gh', [
      'pr',
      'create',
      '--head',
      branch,
      '--base',
      'main',
      '--title',
      `Capture collaboration baseline ${version}`,
      '--body-file',
      file,
    ]);
  } finally {
    rmSync(temp, { recursive: true, force: true });
  }
}
