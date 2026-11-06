/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Runs one shared scenario on any runtime and reports a JSON outcome. The server suite and the
// live editor suite both call it, so the two hosts run exactly
// the same steps and their outcomes can be compared as JSON.

import * as api from '../../runtime/public.ts';
import type { PptxEditorRuntime } from '../../runtime/public.ts';
import type { Json, Scenario, ScenarioAssets } from './types.ts';

/** What a scenario did on one host. */
export type Outcome =
  | { readonly status: 'ok'; readonly value: Json; readonly readBack?: Json }
  | {
      readonly status: 'refused';
      readonly code: string;
      readonly inner: string | null;
      /** The proxy path or member the refusal names. */
      readonly location: string | null;
    }
  | { readonly status: 'failed'; readonly message: string }
  | { readonly status: 'todo'; readonly missing: readonly string[] };

/** The members of `scenario.requires` that the API does not have yet. */
export function missingMembers(scenario: Scenario): string[] {
  const classes = api as unknown as Record<string, { prototype?: object } | undefined>;
  return (scenario.requires ?? []).filter((path) => {
    const [owner, member] = path.split('.');
    const proto = owner ? classes[owner]?.prototype : undefined;
    return !proto || !member || !(member in proto);
  });
}

/** The outcome of an error that left a `run`. */
export function outcomeOf(error: unknown): Outcome {
  if (api.isPptxEditorError(error))
    return {
      status: 'refused',
      code: error.code,
      inner: error.debugInfo.innerError?.code ?? null,
      location: error.debugInfo.errorLocation ?? null,
    };
  return { status: 'failed', message: error instanceof Error ? error.message : String(error) };
}

/** Run the script of a scenario, then its read-back on the same runtime. */
export async function runScenario(
  runtime: PptxEditorRuntime,
  scenario: Scenario,
  assets: ScenarioAssets
): Promise<Outcome> {
  const missing = missingMembers(scenario);
  if (missing.length) return { status: 'todo', missing };
  let value: Json;
  try {
    value = await runtime.run((context) => scenario.script(context, assets));
  } catch (error) {
    return outcomeOf(error);
  }
  if (!scenario.readBack) return { status: 'ok', value };
  try {
    return { status: 'ok', value, readBack: await runtime.run(scenario.readBack) };
  } catch (error) {
    const failed = outcomeOf(error);
    return {
      status: 'failed',
      message: `read-back: ${failed.status === 'refused' ? failed.code : JSON.stringify(failed)}`,
    };
  }
}

/** Run only the read-back of a scenario, for a presentation that was saved and reopened. */
export async function runReadBack(runtime: PptxEditorRuntime, scenario: Scenario): Promise<Json> {
  if (!scenario.readBack) return null;
  return runtime.run(scenario.readBack);
}
