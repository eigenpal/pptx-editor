/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Every refusal reason the core can give through an automation host maps to a public
// diagnostic with a fix hint. A new reason in the core fails here until `edit-diagnostics.ts`
// (or `errors.ts`) maps it.
import { describe, expect, test } from 'bun:test';
import { refusalReasons } from '../refusal-reasons.mjs';
import { hostFailure } from '../../src/runtime/batch.ts';
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { argumentHint } from '../../src/runtime/argument-hints.ts';
import { diagnosticOf } from '../../src/runtime/errors.ts';

/** The host code an edit refusal reaches the runtime with (`core/src/automation/targets.ts`). */
function refusalCode(reason: string): string {
  if (reason === 'errors.invalidName') return 'invalid-argument';
  return reason.startsWith('disabled.') ? 'unsupported-content' : 'transaction-refused';
}

describe('refusal reasons', () => {
  const { reasons, dynamic } = refusalReasons() as {
    reasons: { kind: string; reason: string; codes: string[]; files: string[] }[];
    dynamic: string[];
  };

  test('the scan finds the edit refusals of the core', () => {
    const names = reasons.map((r) => r.reason);
    expect(names).toContain('disabled.groupPlaceholder');
    expect(names).toContain('disabled.splitUneven');
    expect(names).toContain('errors.coordinateRange');
    expect(reasons.length).toBeGreaterThan(100);
    // Call sites whose reason is not a literal: argument names and host codes, which the
    // fallback by host code covers. A new one is reviewed here.
    expect(dynamic.length).toBeLessThanOrEqual(40);
  });

  test('every edit refusal has a diagnostic of its own with a hint', () => {
    const missing = reasons
      .filter((r) => r.kind === 'refusal')
      .filter((r) => {
        const d = diagnosticOf(r.reason);
        return !d || d.hint.length < 20;
      })
      .map((r) => r.reason);
    expect(missing).toEqual([]);
  });

  test('every host refusal reaches the caller with an innerError', () => {
    const missing: string[] = [];
    for (const r of reasons) {
      const codes = r.kind === 'refusal' ? [refusalCode(r.reason)] : r.codes;
      for (const code of codes) {
        const error = hostFailure({ code: code as never, message: 'm', detail: r.reason });
        if (!error.debugInfo.innerError?.message) missing.push(`${code} ${r.reason}`);
      }
    }
    expect(missing).toEqual([]);
  });

  test('non-argument failures with a detail of their own are mapped, not only by code', () => {
    const generic = reasons
      .filter((r) => r.kind === 'failure')
      .filter((r) => r.codes.some((c) => c !== 'invalid-argument' && c !== 'item-not-found'))
      .filter((r) => !diagnosticOf(r.reason))
      .map((r) => r.reason);
    expect(generic).toEqual([]);
  });
  test('every argument the engine refuses has a hint of its own', () => {
    const generic = reasons
      .filter((r) => r.kind === 'failure' && r.codes.includes('invalid-argument'))
      .filter((r) => !diagnosticOf(r.reason) && !argumentHint(r.reason))
      .map((r) => r.reason);
    expect(generic).toEqual([]);
  });

  test('no argument refusal of the client falls back to the generic text', () => {
    // Every client check of an argument goes through `failArgument`, which says what to pass.
    const src = join(import.meta.dir, '../../src');
    const offenders: string[] = [];
    for (const dir of ['model', 'runtime'])
      for (const name of readdirSync(join(src, dir))) {
        if (!name.endsWith('.ts')) continue;
        const text = readFileSync(join(src, dir, name), 'utf8');
        for (const m of text.matchAll(/fail\(\{\s*code: 'InvalidArgument'[^}]*\}\)/g))
          if (!m[0].includes('diagnostic')) offenders.push(`${dir}/${name}: ${m[0]}`);
      }
    expect(offenders).toEqual([]);
  });
});
