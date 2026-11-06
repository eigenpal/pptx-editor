/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The shared scenario catalog on a browser runtime over a live editor, in bun.
//
// Each scenario must answer what it answers on a server runtime, except for the documented
// host difference: the live editor has a selection. `scenario-hosts.ts` holds the checks: a
// whole save that reopens with the same read-back, and one undo step for each sync that wrote.

import { describe, test } from 'bun:test';
import { assets, checkEditor, serverRuntime } from './scenario-hosts.ts';
import { SCENARIOS, type Scenario } from './scenarios/index.ts';
import { missingMembers, runScenario, type Outcome } from './scenarios/runner.ts';

/** What the scenario answers on a server runtime, for the comparison. */
async function onServer(scenario: Scenario): Promise<Outcome> {
  const runtime = await serverRuntime(scenario);
  try {
    return await runScenario(runtime, scenario, assets);
  } finally {
    runtime.dispose();
  }
}

describe('scenarios on a live editor', () => {
  for (const scenario of SCENARIOS) {
    const missing = missingMembers(scenario);
    if (missing.length) {
      test.todo(`${scenario.name}: needs ${missing.join(', ')}`);
      continue;
    }
    test(scenario.name, async () => {
      await checkEditor(scenario, await onServer(scenario));
    });
  }
});
