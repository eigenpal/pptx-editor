/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The shared scenario catalog on the server runtime, over the fixture decks.
//
// Each scenario runs on a fresh runtime. `scenario-hosts.ts` holds the checks: a scenario that
// writes saves a whole package that reopens with the same read-back, and a scenario that does
// not write leaves the package as it was. The outcomes are pinned in a snapshot.
// `scenarios-editor.test.ts` runs the same catalog on a live editor and compares.

import { describe, expect, test } from 'bun:test';
import { checkServer } from './scenario-hosts.ts';
import { SCENARIOS, type ScenarioArea } from './scenarios/index.ts';
import { missingMembers } from './scenarios/runner.ts';

describe('the scenario catalog', () => {
  test('covers every area of the API with at least 30 scenarios', () => {
    const areas: ScenarioArea[] = [
      'presentation',
      'slides',
      'shapes',
      'text',
      'tables',
      'custom-xml',
      'insert',
      'images',
      'selection',
      'load',
      'sync',
      'errors',
    ];
    expect(SCENARIOS.length).toBeGreaterThanOrEqual(30);
    for (const area of areas)
      expect([area, SCENARIOS.some((s) => s.area === area)]).toEqual([area, true]);
    expect(new Set(SCENARIOS.map((s) => s.name)).size).toBe(SCENARIOS.length);
    for (const s of SCENARIOS)
      if (s.writes)
        expect([s.name, !!s.readBack || !!s.host || !!s.requires]).toEqual([s.name, true]);
  });
});

describe('scenarios on the server runtime', () => {
  for (const scenario of SCENARIOS) {
    const missing = missingMembers(scenario);
    if (missing.length) {
      test.todo(`${scenario.name}: needs ${missing.join(', ')}`);
      continue;
    }
    test(scenario.name, async () => {
      expect(await checkServer(scenario)).toMatchSnapshot();
    });
  }
});
