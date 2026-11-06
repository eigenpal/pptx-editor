/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The two hosts that run the shared scenarios, in bun: a server runtime over PPTX bytes, and a
// browser runtime over a live editor (`createPptxEditor` with `createBrowserAutomationHost`).
//
// Each check runs one scenario on a fresh runtime and asserts what that host must keep:
// - Both: a scenario that writes saves a whole package (`scripts/fixtures/package-integrity.ts`)
//   that reopens on a server runtime with the same read-back. A scenario that does not write
//   leaves the package as it was opened.
// - The live editor: each sync that writes is one undo step, so undoing steps back through the
//   package after each such sync, in reverse, to the package that was opened.
// - The documented host difference (`HostDifference`): a server runtime refuses selection; the
//   live editor answers it. Pictures answer the same on both: the server draws them with a stub
//   rasterizer, and the live editor with the real paint on a stub canvas.

import { expect } from 'bun:test';
import { readFileSync } from 'node:fs';
import type { AutomationHost } from '@pptx-editor.dev/core/automation';
import {
  createBrowserAutomationHost,
  createPptxEditor,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { canonicalPackageFingerprint, readPptx } from '@pptx-editor.dev/core/store';
import { packageIntegrity } from '../../../../scripts/fixtures/package-integrity.ts';
import { PptxEditor } from '../browser.ts';
import type { PptxEditorServerRuntime } from '../runtime/public.ts';
import { createRuntime } from '../runtime/runtime.ts';
import { FIXTURES, type FixtureName, type Scenario } from './scenarios/index.ts';
import { runReadBack, runScenario, type Outcome } from './scenarios/runner.ts';
import { stubRasterize, withStubDrawing } from './stub-drawing.ts';

const bytes = Object.fromEntries(
  FIXTURES.map((name) => [
    name,
    new Uint8Array(readFileSync(new URL(`../../../../fixtures/${name}.pptx`, import.meta.url))),
  ])
) as Record<FixtureName, Uint8Array>;

export const assets = {
  base64: Object.fromEntries(
    FIXTURES.map((name) => [name, Buffer.from(bytes[name]).toString('base64')])
  ) as Record<FixtureName, string>,
};

export function fingerprint(saved: Uint8Array): string {
  const read = readPptx(saved);
  if (!read.ok) throw new Error('the saved presentation does not read back');
  return canonicalPackageFingerprint(read.pkg);
}

/** The refusal a server runtime gives for a documented host difference. */
export const HOST_REFUSAL = {
  selection: { status: 'refused', code: 'NotSupported', inner: 'host.no-selection' },
} as const;

/** Check that `saved` is whole and reopens on a server runtime with the same read-back. */
async function reopens(saved: Uint8Array, scenario: Scenario, outcome: Outcome): Promise<void> {
  expect(packageIntegrity(saved).issues).toEqual([]);
  // A host refusal stops the script; the syncs before it applied.
  if (outcome.status !== 'ok') return;
  const reopened = await PptxEditor.createServer(saved);
  try {
    expect(await runReadBack(reopened, scenario)).toEqual(outcome.readBack ?? null);
  } finally {
    reopened.dispose();
  }
}

/** A server runtime over the scenario's deck, which draws pictures with the stub rasterizer. */
export function serverRuntime(scenario: Scenario): Promise<PptxEditorServerRuntime> {
  return PptxEditor.createServer(bytes[scenario.deck], { rasterize: stubRasterize });
}

/** Run `scenario` on a server runtime and check it. */
export async function checkServer(scenario: Scenario): Promise<Outcome> {
  const runtime = await serverRuntime(scenario);
  try {
    const before = fingerprint(await runtime.save());
    const outcome = await runScenario(runtime, scenario, assets);
    if (scenario.host === 'selection') {
      expect(outcome).toMatchObject(HOST_REFUSAL.selection);
    } else {
      expect(outcome.status === 'failed' ? outcome.message : outcome.status).toBe('ok');
    }
    const saved = await runtime.save();
    // Nothing was written, and every refused batch changed nothing.
    if (!scenario.writes) expect(fingerprint(saved)).toBe(before);
    else await reopens(saved, scenario, outcome);
    return outcome;
  } finally {
    runtime.dispose();
  }
}

/** The package the live editor holds now. */
function editorPackage(editor: PptxEditorInstance): string {
  const saved = editor.save();
  if (!saved) throw new Error('the editor did not save');
  return fingerprint(saved);
}

/** The package after each sync that changed it, and the history group of that sync. */
interface Step {
  readonly pkg: string;
  readonly group: string | undefined;
}

/**
 * The editor's automation host, recording the package after every batch. The batches of one
 * sync carry one history group, so a sync that splits is still one step.
 */
function recordingHost(editor: PptxEditorInstance, steps: Step[]): AutomationHost {
  const host = createBrowserAutomationHost(editor);
  const recording = Object.create(host) as AutomationHost;
  let revision = host.revision();
  recording.execute = (request) => {
    const response = host.execute(request);
    // Only a batch that moved the revision can have changed the package.
    if (host.revision() === revision) return response;
    revision = host.revision();
    const pkg = editorPackage(editor);
    const last = steps[steps.length - 1]!;
    if (pkg !== last.pkg) {
      const step = { pkg, group: request.historyGroup };
      if (last.group !== undefined && last.group === request.historyGroup)
        steps[steps.length - 1] = step;
      else steps.push(step);
    }
    return response;
  };
  return recording;
}

/** Run `scenario` on a browser runtime over a live editor, and check it against `server`. */
export function checkEditor(scenario: Scenario, server: Outcome): Promise<Outcome> {
  // Pictures need a DOM when the host starts. The paint and the fit are real; the canvas is a
  // stub.
  const check = (): Promise<Outcome> => editorCheck(scenario, server);
  return scenario.area === 'images' ? withStubDrawing(check) : check();
}

async function editorCheck(scenario: Scenario, server: Outcome): Promise<Outcome> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(bytes[scenario.deck].slice())).ok).toBe(true);
  const steps: Step[] = [{ pkg: editorPackage(editor), group: undefined }];
  // `PptxEditor.createBrowser(editor)` over the recording host.
  const runtime = createRuntime({ host: recordingHost(editor, steps), save: false });
  try {
    const outcome = await runScenario(runtime, scenario, assets);
    if (scenario.host === 'selection') {
      expect(outcome.status === 'failed' ? outcome.message : outcome.status).toBe('ok');
    } else {
      expect(outcome).toEqual(server);
    }
    if (!scenario.writes) {
      expect(steps).toHaveLength(1);
      expect(editor.snapshot().canUndo).toBe(false);
      return outcome;
    }
    const saved = editor.save();
    if (!saved) throw new Error('the editor did not save');
    await reopens(saved, scenario, outcome);
    // One undo step for each sync that wrote, back to the package that was opened.
    if (outcome.status === 'ok')
      expect([scenario.name, steps.length > 1]).toEqual([scenario.name, true]);
    for (let i = steps.length - 1; i > 0; i--) {
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect([scenario.name, i, editorPackage(editor)]).toEqual([
        scenario.name,
        i,
        steps[i - 1]!.pkg,
      ]);
    }
    expect(editor.snapshot().canUndo).toBe(false);
    return outcome;
  } finally {
    runtime.dispose();
  }
}
