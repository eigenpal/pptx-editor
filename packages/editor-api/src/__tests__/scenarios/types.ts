/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The shape of a shared automation scenario and the plain helpers its scripts use.
//
// A scenario is an Office.js add-in script: `async (context) => { ... await context.sync(); }`.
// The same script runs on a server runtime and on a browser runtime over a live editor, so
// this module and the scenario modules import only types at run time. Values
// such as enum members are written as their string literals, which the API accepts.

import type { RequestContext } from '../../runtime/public.ts';

/** A JSON value: what a scenario answers, so two hosts can compare it. */
export type Json = null | boolean | number | string | Json[] | { [key: string]: Json };

/** The decks under `fixtures/` that scenarios open. */
export type FixtureName = 'slide-layouts' | 'tables' | 'pictures' | 'text-runs';

/** Each fixture a scenario may name, for runners that preload them. */
export const FIXTURES: readonly FixtureName[] = [
  'slide-layouts',
  'tables',
  'pictures',
  'text-runs',
];

/** What a runner hands a script besides the context: decks as base64, as an add-in gets them. */
export interface ScenarioAssets {
  readonly base64: Readonly<Record<FixtureName, string>>;
}

/** The API areas the catalog covers; each has scenarios. */
export type ScenarioArea =
  | 'presentation'
  | 'slides'
  | 'shapes'
  | 'text'
  | 'tables'
  | 'custom-xml'
  | 'insert'
  | 'images'
  | 'selection'
  | 'load'
  | 'sync'
  | 'errors';

/**
 * A documented difference between the hosts:
 * - `selection`: a server runtime has no selection, so it refuses with `NotSupported` and the
 *   diagnostic `host.no-selection`.
 */
export type HostDifference = 'selection';

export interface Scenario {
  /** A stable name: `<area>.<what>`. */
  readonly name: string;
  readonly area: ScenarioArea;
  readonly deck: FixtureName;
  /** The script changes the presentation, so runners save, reopen, and check the package. */
  readonly writes: boolean;
  /** The documented host difference this scenario shows, if any. */
  readonly host?: HostDifference;
  /**
   * Members the scenario needs as `Class.member`. While one is missing, runners report the
   * scenario as to do, not as a pass or a failure.
   */
  readonly requires?: readonly string[];
  /** The add-in script. It throws (through `check`) when the API answers a wrong value. */
  readonly script: (context: RequestContext, assets: ScenarioAssets) => Promise<Json>;
  /**
   * Reads what the script wrote. Runners call it on the edited presentation and again on the
   * saved and reopened presentation; the two answers must be equal.
   */
  readonly readBack?: (context: RequestContext) => Promise<Json>;
}

/** Fail the scenario with a message when a condition does not hold. */
export function check(condition: unknown, message: string): asserts condition {
  if (!condition) throw new Error(`check failed: ${message}`);
}

/** Fail the scenario when two JSON values differ. */
export function same(actual: unknown, expected: unknown, message: string): void {
  const a = JSON.stringify(actual);
  const b = JSON.stringify(expected);
  if (a !== b) throw new Error(`check failed: ${message}: got ${a}, expected ${b}`);
}

/** A refusal as JSON: the Office.js error code and the engine diagnostic. */
export interface Refusal {
  readonly code: string;
  readonly inner: string | null;
  [key: string]: Json;
}

/** Run a step that must be refused, and answer its error code and diagnostic. */
export async function refusal(step: () => Promise<unknown>): Promise<Refusal> {
  try {
    await step();
  } catch (error) {
    const e = error as { code?: unknown; debugInfo?: { innerError?: { code?: unknown } } };
    if (typeof e?.code !== 'string') throw error;
    const inner = e.debugInfo?.innerError?.code;
    return { code: e.code, inner: typeof inner === 'string' ? inner : null };
  }
  throw new Error('check failed: the step was not refused');
}

/** A date as JSON: its ISO text, or `invalid`. */
export function dateJson(value: Date): string {
  return Number.isNaN(value.getTime()) ? 'invalid' : value.toISOString();
}

/** Round a point value to two decimals, so EMU rounding does not hide in the comparison. */
export function pt(value: number): number {
  return Math.round(value * 100) / 100;
}

/** The size of a PNG from its base64 text, read from the IHDR chunk. */
export function pngSize(base64: string): { signature: boolean; width: number; height: number } {
  const head = atob(base64.slice(0, 64));
  const byte = (i: number): number => head.charCodeAt(i);
  const signature = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a].every((b, i) => byte(i) === b);
  const u32 = (i: number): number =>
    ((byte(i) << 24) | (byte(i + 1) << 16) | (byte(i + 2) << 8) | byte(i + 3)) >>> 0;
  return { signature, width: u32(16), height: u32(20) };
}
