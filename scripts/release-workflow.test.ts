// After DOCX Editor's scripts/release-workflow.test.ts (commit b9a04a30f), for this
// repository's jobs. Evaluates the workflow's own conditions, so a changed `if:` that lets
// a canary reach the stable path, or a skipped gate reach a publish, fails here.
import { expect, test } from 'bun:test';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';

interface Step {
  name?: string;
  run?: string;
  uses?: string;
  if?: string;
  with?: Record<string, string>;
  env?: Record<string, string>;
  'continue-on-error'?: boolean;
}
interface Job {
  needs?: string | string[];
  if?: string;
  'runs-on'?: string;
  permissions?: Record<string, string>;
  outputs?: Record<string, string>;
  'continue-on-error'?: boolean;
  steps: Step[];
}
const read = (name: string): unknown =>
  Bun.YAML.parse(readFileSync(new URL(`../.github/workflows/${name}`, import.meta.url), 'utf8'));
const workflow = read('release.yml') as {
  on: {
    push?: unknown;
    schedule?: { cron: string }[];
    workflow_dispatch?: { inputs?: Record<string, { options?: string[]; default?: string }> };
  };
  concurrency: { group: string; 'cancel-in-progress': boolean };
  jobs: Record<string, Job>;
};

// Evaluate a workflow expression with the given contexts: `||`, `&&`, `!`, `==`, `!=`,
// parentheses, string literals, context paths, and status calls such as `failure()`. As in
// GitHub Actions, `||` and `&&` answer an operand, and a missing value compares unequal to a
// string. A parser, not `new Function`, so the repository's eval ban holds.
function evaluate(expression: string, contexts: Record<string, unknown>): unknown {
  const source = expression.replace(/^\$\{\{\s*/, '').replace(/\s*\}\}$/, '');
  const tokens = source.match(/'[^']*'|\|\||&&|[!=]=|[!()]|[A-Za-z_][\w.-]*/g) ?? [];
  if (tokens.join('') !== source.replace(/\s+/g, '')) throw new Error(`Cannot parse ${source}`);
  let index = 0;
  const peek = (): string | undefined => tokens[index];
  const take = (expected?: string): string => {
    const token = tokens[index++];
    if (token === undefined || (expected !== undefined && token !== expected))
      throw new Error(`Expected ${expected ?? 'a token'} in ${source}`);
    return token;
  };
  const literals: Record<string, unknown> = { true: true, false: false, null: null };
  const lookup = (path: string): unknown => {
    if (path in literals) return literals[path];
    const [root, ...keys] = path.split('.');
    // A context the test did not supply is a test error, not a missing value.
    if (!(root in contexts)) throw new Error(`No context ${root} for ${source}`);
    return keys.reduce<unknown>(
      (value, key) => (value as Record<string, unknown> | undefined)?.[key],
      contexts[root]
    );
  };
  const primary = (): unknown => {
    const token = take();
    if (token === '!') return !primary();
    if (token === '(') {
      const value = or();
      take(')');
      return value;
    }
    if (token.startsWith("'")) return token.slice(1, -1);
    if (peek() === '(') {
      take('(');
      take(')');
      return (lookup(token) as () => unknown)();
    }
    return lookup(token);
  };
  const equality = (): unknown => {
    let value = primary();
    while (peek() === '==' || peek() === '!=') {
      const operator = take();
      const right = primary();
      value = operator === '==' ? value === right : value !== right;
    }
    return value;
  };
  const and = (): unknown => {
    let value = equality();
    while (peek() === '&&') {
      take();
      const right = equality();
      value = value ? right : value;
    }
    return value;
  };
  const or = (): unknown => {
    let value = and();
    while (peek() === '||') {
      take();
      const right = and();
      value = value ? value : right;
    }
    return value;
  };
  const value = or();
  if (index !== tokens.length) throw new Error(`Unexpected ${peek()} in ${source}`);
  return value;
}
const runs = {
  push: { github: { event_name: 'push', ref: 'refs/heads/main' }, inputs: {} },
  schedule: { github: { event_name: 'schedule', ref: 'refs/heads/main' }, inputs: {} },
  dispatchRelease: {
    github: { event_name: 'workflow_dispatch', ref: 'refs/heads/main' },
    inputs: { channel: 'release' },
  },
  dispatchCanary: {
    github: { event_name: 'workflow_dispatch', ref: 'refs/heads/main' },
    inputs: { channel: 'canary' },
  },
};
function needs(mode: string, results: Record<string, string> = {}) {
  return {
    mode: { result: results.mode ?? 'success', outputs: { mode, commit: 'a'.repeat(40) } },
    checks: { result: results.checks ?? 'success' },
    build: { result: results.build ?? 'success' },
  };
}
const skipped = { checks: 'skipped', build: 'skipped' };
const stable = (job: string, contexts: Record<string, unknown>, cancelled = false): unknown =>
  evaluate(workflow.jobs[job].if!, { ...runs.push, cancelled: () => cancelled, ...contexts });

test('publication requires every prepublish gate to succeed', () => {
  expect(workflow.jobs.release.needs).toEqual(['mode', 'checks', 'build']);
  expect(stable('release', { needs: needs('publish') })).toBe(true);
  for (const gate of ['checks', 'build'])
    for (const result of ['failure', 'cancelled', 'skipped'])
      expect(stable('release', { needs: needs('publish', { [gate]: result }) })).toBe(false);
  expect(stable('release', { needs: needs('publish') }, true)).toBe(false);
  expect(stable('release', { needs: needs('publish', { mode: 'failure' }) })).toBe(false);
});

test('pending changesets take the PR-update path without the prepublish gates', () => {
  expect(stable('release', { needs: needs('pending', skipped) })).toBe(true);
  for (const job of ['checks', 'build']) {
    expect(stable(job, { needs: needs('publish') })).toBe(true);
    expect(stable(job, { needs: needs('pending') })).toBe(false);
    expect(workflow.jobs[job]['continue-on-error']).not.toBe(true);
  }
});

test('no stable job runs outside main, so other branches never publish', () => {
  for (const ref of ['refs/heads/changeset-release/main', 'refs/heads/feature', 'refs/tags/v1']) {
    const github = { event_name: 'workflow_dispatch', ref };
    for (const job of ['checks', 'build'])
      expect(stable(job, { github, needs: needs('publish') })).toBe(false);
    expect(stable('release', { github, needs: needs('publish') })).toBe(false);
    expect(stable('release', { github, needs: needs('pending', skipped) })).toBe(false);
  }
  const refusal = workflow.jobs.mode.steps.find(
    (step) => step.name === 'Refuse stable releases outside main'
  )!;
  expect(refusal.if).toBe(
    "${{ env.RELEASE_CHANNEL == 'release' && github.ref != 'refs/heads/main' }}"
  );
});

test('the canary runs nightly and on demand, in its own concurrency group', () => {
  expect(workflow.on).toHaveProperty('push');
  expect(workflow.on.schedule).toEqual([{ cron: '23 3 * * *' }]);
  expect(workflow.on.workflow_dispatch?.inputs?.channel.options).toEqual(['release', 'canary']);
  expect(workflow.on.workflow_dispatch?.inputs?.channel.default).toBe('release');
  const group = (run: Record<string, unknown>) => evaluate(workflow.concurrency.group, run);
  expect(group(runs.push)).toBe('release');
  expect(group(runs.dispatchRelease)).toBe('release');
  expect(group(runs.schedule)).toBe('release-canary');
  expect(group(runs.dispatchCanary)).toBe('release-canary');
  expect(workflow.concurrency['cancel-in-progress']).toBe(false);
});

test('a canary run can only answer canary or idle', () => {
  const output = workflow.jobs.mode.outputs!.mode;
  const answer = (run: Record<string, unknown>, mode?: string) =>
    evaluate(output, { ...run, steps: { mode: { outputs: { mode } } } });
  for (const run of [runs.schedule, runs.dispatchCanary]) {
    for (const mode of ['pending', 'publish', 'idle', '', undefined])
      expect(answer(run, mode)).toBe('idle');
    expect(answer(run, 'canary')).toBe('canary');
  }
  for (const run of [runs.push, runs.dispatchRelease])
    for (const mode of ['pending', 'publish']) expect(answer(run, mode)).toBe(mode);
});

test('only a canary run with a green commit reaches the canary job', () => {
  const job = workflow.jobs.canary;
  const at = (run: Record<string, unknown>, value: unknown) =>
    evaluate(job.if!, { ...run, needs: value });
  expect(at(runs.schedule, needs('canary'))).toBe(true);
  expect(at(runs.dispatchCanary, needs('canary'))).toBe(true);
  expect(at(runs.push, needs('canary'))).toBe(false);
  expect(at(runs.dispatchRelease, needs('canary'))).toBe(false);
  expect(at(runs.schedule, needs('idle'))).toBe(false);
  expect(at(runs.schedule, needs('canary', { mode: 'failure' }))).toBe(false);
  const noCommit = needs('canary');
  noCommit.mode.outputs.commit = '';
  expect(at(runs.schedule, noCommit)).toBe(false);
  // The stable path never runs for a canary or idle answer.
  for (const mode of ['canary', 'idle']) {
    for (const name of ['checks', 'build'])
      expect(stable(name, { needs: needs(mode) })).toBe(false);
    expect(stable('release', { needs: needs(mode, skipped) })).toBe(false);
  }
});

test('the canary versions before it builds and publishes only to the canary tag', () => {
  const job = workflow.jobs.canary;
  expect(job['runs-on']).toMatch(/^ubuntu-/);
  expect(job.permissions).toEqual({ contents: 'read', actions: 'read', 'id-token': 'write' });
  const index = (run: string) => job.steps.findIndex((step) => step.run === run);
  const version = index('bun run changeset version --snapshot canary');
  const build = index('bun run build:packages');
  const install = index('bun run check:consumer-install');
  const gate = job.steps.findIndex((step) => step.name === 'Confirm no stable release is active');
  const publish = index('bun run changeset publish --tag canary --no-git-tag');
  expect(version).toBeGreaterThan(0);
  expect(build).toBeGreaterThan(version);
  expect(install).toBeGreaterThan(build);
  expect(gate).toBeGreaterThan(install);
  expect(publish).toBe(gate + 1);
  // The gate is the last check before publishing, and only its clear answer publishes.
  expect(job.steps[gate].run).toBe('node scripts/release-mode.mjs');
  expect(job.steps[gate].env?.RELEASE_CHANNEL).toBe('canary-gate');
  expect(job.steps[publish].if).toBe("${{ steps.gate.outputs.clear == 'true' }}");
  // npm refuses provenance from a private repository, and a scheduled event carries no
  // `repository` object, so the visibility comes from a step that reads the API.
  const provenance = job.steps[publish].env?.NPM_CONFIG_PROVENANCE ?? '';
  expect(Object.keys(job.steps[publish].env ?? {})).toEqual(['NPM_CONFIG_PROVENANCE']);
  const visibility = job.steps.findIndex((step) => step.name === 'Read the repository visibility');
  expect(visibility).toBeGreaterThan(install);
  expect(visibility).toBeLessThan(gate);
  const signed = (value: string) =>
    evaluate(provenance, { steps: { visibility: { outputs: { private: value } } } });
  expect(signed('false')).toBe('true');
  expect(signed('true')).toBe('false');
  expect(signed('')).toBe('false');
  for (const step of job.steps) {
    expect(step.uses ?? '').not.toContain('changesets/action');
    expect(step.run ?? '').not.toMatch(/git (?:tag|push)|gh (?:release|workflow run)/);
    expect(step['continue-on-error']).not.toBe(true);
    expect(JSON.stringify(step)).not.toMatch(/slack/i);
  }
  const checkout = job.steps.find((step) => step.uses?.startsWith('actions/checkout@'));
  expect(checkout?.with?.ref).toBe('${{ needs.mode.outputs.commit }}');
});

test('a canary run never reaches the stable failure alert', () => {
  const alert = workflow.jobs['notify-failure'].if!;
  for (const run of [runs.schedule, runs.dispatchCanary])
    expect(evaluate(alert, { ...run, failure: () => true })).toBe(false);
  for (const run of [runs.push, runs.dispatchRelease])
    expect(evaluate(alert, { ...run, failure: () => true })).toBe(true);
});

test('the canary commit comes from a green main CI run in this repository', () => {
  const step = workflow.jobs.mode.steps.find(
    (entry) => entry.name === 'Find the newest main commit that passed CI'
  )!;
  expect(step.run).toContain('"$REPOSITORY" != "eigenpal/pptx-editor"');
  expect(step.run).toContain('"$REF" != "refs/heads/main"');
  expect(step.run).toContain('--workflow ci.yml --branch main --event push');
  expect(step.run).toContain('--status success');
  expect(workflow.jobs.mode.permissions).toEqual({ contents: 'read', actions: 'read' });
});

test('post-release acts only on a run whose stable Release job passed', () => {
  const post = read('post-release.yml') as { jobs: { source: Job } };
  const source = post.jobs.source;
  const run = (event: string, repository = 'eigenpal/pptx-editor') =>
    evaluate(source.if!, {
      github: {
        repository: 'eigenpal/pptx-editor',
        event: { workflow_run: { event, head_repository: { full_name: repository } } },
      },
    });
  expect(run('push')).toBe(true);
  expect(run('workflow_dispatch')).toBe(true);
  expect(run('schedule')).toBe(false);
  expect(run('push', 'fork/pptx-editor')).toBe(false);
  // A manual canary run is a workflow_dispatch: the step checks the Release job by name.
  expect(workflow.jobs.release).toHaveProperty('name', 'Release');
  const identify = source.steps.find(
    (step) => step.name === 'Identify a published and tagged release'
  )!;
  expect(identify.run).toContain('select(.name == "Release")');
  expect(identify.run).toContain('if [ "$release" != "success" ]');
  expect(source.permissions).toEqual({ contents: 'read', actions: 'read' });
});

test('label sync requires an explicit PPTX project and a relevant label', () => {
  const sync = read('sync-labels-to-project.yml') as {
    env: Record<string, string>;
    jobs: { sync: Job };
  };
  expect(sync.env.PROJECT_NUMBER).toBe('${{ vars.PPTX_ROADMAP_PROJECT_NUMBER }}');
  for (const project of ['', '12']) {
    for (const label of ['area:fidelity', 'priority:high', 'bug']) {
      const condition = sync.jobs.sync.if!.replace(
        /startsWith\(github.event.label.name, '([^']+)'\)/g,
        (_, prefix: string) => String(label.startsWith(prefix))
      );
      expect(evaluate(condition, { vars: { PPTX_ROADMAP_PROJECT_NUMBER: project } })).toBe(
        project !== '' && label !== 'bug'
      );
    }
  }
});

for (const project of ['', '12']) {
  test(`release comments work with ${project ? 'an explicit' : 'no'} roadmap project`, () => {
    const post = read('post-release.yml') as { jobs: { announcements: Job } };
    const step = post.jobs.announcements.steps.find((entry) =>
      entry.name?.startsWith('Comment on shipped PRs')
    )!;
    expect(step.env?.PROJECT_NUMBER).toBe('${{ vars.PPTX_ROADMAP_PROJECT_NUMBER }}');
    const directory = mkdtempSync(join(tmpdir(), 'pptx-release-project-'));
    const calls = join(directory, 'calls');
    try {
      // Run the workflow shell itself. The mocks expose one shipped PR and its closed issue.
      // A project lookup returns an existing field, so only this release gets updated.
      const mocks = `
        gh() {
          printf '%s\\n' "$*" >> "$CALLS"
          case "$*" in
            *'projectV2(number:'*) printf '%s' '{"data":{"organization":{"projectV2":{"id":"PROJECT","field":{"id":"FIELD"}}}}}' ;;
            *'closingIssuesReferences'*) printf '%s' '{"data":{"repository":{"pullRequest":{"closingIssuesReferences":{"nodes":[{"number":8,"repository":{"nameWithOwner":"eigenpal/pptx-editor"}}]}}}}}' ;;
            *'projectItems(first:'*) printf '%s' '{"data":{"repository":{"issue":{"projectItems":{"nodes":[{"id":"ITEM","project":{"id":"PROJECT"}}]}}}}}' ;;
          esac
        }
        git() {
          case "$*" in
            'tag '*) printf 'v1.0.0\\nv1.0.1\\n' ;;
            'log --format=%s '*) printf 'fix: example (#7)\\n' ;;
            'log --format=%s %b '*) printf 'fix: example fixes #8\\n' ;;
          esac
        }
      `;
      const result = spawnSync('bash', ['-e', '-o', 'pipefail', '-c', mocks + step.run], {
        env: {
          ...process.env,
          CALLS: calls,
          VERSION: '1.0.1',
          REPO: 'eigenpal/pptx-editor',
          SERVER_URL: 'https://github.com',
          ORG: 'eigenpal',
          PROJECT_NUMBER: project,
          PROJECT_TOKEN: project ? 'test-project-token' : '',
          FIELD_NAME: 'Shipped in',
        },
        encoding: 'utf8',
      });
      expect({ status: result.status, stderr: result.stderr }).toEqual({ status: 0, stderr: '' });
      const recorded = readFileSync(calls, 'utf8');
      expect(recorded).toContain('pr comment 7');
      expect(recorded).toContain('issue comment 8');
      if (project) {
        expect(recorded).toContain('number=12');
        expect(recorded).toContain('updateProjectV2ItemFieldValue');
      } else {
        expect(recorded).not.toMatch(/projectV2|projectItems|ProjectV2/);
      }
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });
}

test('every publishing job builds, writes the notices, and checks the artifacts in that order', () => {
  for (const name of ['build', 'canary']) {
    const steps = workflow.jobs[name].steps;
    const index = (run: string) => steps.findIndex((step) => step.run === run);
    const build = index('bun run build:packages');
    const notices = index('bun run notices:generate');
    const artifacts = index('bun run check:package-artifacts');
    expect(build).toBeGreaterThanOrEqual(0);
    expect(notices).toBeGreaterThan(build);
    expect(artifacts).toBeGreaterThan(notices);
    for (const step of steps.slice(build, artifacts + 1)) {
      expect(step['continue-on-error']).toBeUndefined();
      expect(step.if).toBeUndefined();
    }
  }
});
