import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
const workflow = (name: string): any =>
  Bun.YAML.parse(readFileSync(new URL(`../../.github/workflows/${name}`, import.meta.url), 'utf8'));

test('stable publication checks downloaded artifacts before publishing', () => {
  const steps = workflow('release.yml').jobs.release.steps;
  const download = steps.findIndex((step: any) => step.name === 'Download publish artifacts');
  const gate = steps.findIndex((step: any) => step.run === 'bun run collaboration:test');
  const publish = steps.findIndex((step: any) => step.id === 'changesets');
  expect(gate).toBeGreaterThan(download);
  expect(publish).toBeGreaterThan(gate);
  expect(steps[gate].if).toContain("needs.mode.outputs.mode == 'publish'");
  expect(steps[gate]['continue-on-error']).toBeUndefined();
  expect(steps.slice(gate + 1, publish).some((step: any) => step.run?.includes('build'))).toBe(
    false
  );
});

test('post-release capture has a manual retry and App-authenticated pull requests', () => {
  const post = workflow('post-release.yml');
  expect(post.jobs['collaboration-baseline'].needs).toBe('source');
  expect(post.jobs['collaboration-baseline'].if).toContain("published == 'true'");
  expect(post.jobs.report.needs).toContain('collaboration-baseline');
  const capture = workflow('collaboration-baseline.yml');
  expect(capture.on.workflow_dispatch.inputs.version.required).toBe(true);
  expect(capture.on.workflow_call.inputs.version.required).toBe(true);
  expect(capture.concurrency['cancel-in-progress']).toBe(false);
  expect(capture.jobs.capture.steps[0].uses).toContain('actions/create-github-app-token@');
});

test('baseline merge receives workflow approval permissions through the reusable call', () => {
  const caller = workflow('post-release.yml').jobs['collaboration-baseline'];
  const baseline = workflow('collaboration-baseline.yml');
  const capture = baseline.jobs.capture;
  const merge = baseline.jobs.merge;
  expect(capture.outputs.pr).toBe('${{ steps.capture.outputs.pr }}');
  expect(capture.steps.find((step: any) => step.id === 'capture').run).toContain(
    'node scripts/collaboration/post-release.mjs'
  );
  expect(merge.needs).toBe('capture');
  expect(merge.if).toBe("needs.capture.outputs.pr != ''");
  expect(merge.env.PR_NUMBER).toBe('${{ needs.capture.outputs.pr }}');
  expect(merge.env.RELEASE_VERSION).toBe('${{ inputs.version }}');
  for (const permissions of [caller.permissions, merge.permissions]) {
    expect(permissions['pull-requests']).toBe('write');
    for (const scope of ['contents', 'actions', 'checks', 'statuses'])
      expect(permissions[scope]).toBe('read');
  }
  const final = merge.steps.at(-1);
  expect(final.env.GH_TOKEN).toBe('${{ github.token }}');
  expect(final.env.MERGE_TOKEN).toBe('${{ steps.token.outputs.token }}');
  expect(final.run).toContain('-f event=APPROVE -f commit_id="$EXPECTED_HEAD"');
  expect(final.run).toContain('GH_TOKEN="$MERGE_TOKEN" gh pr merge');
});
