/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The shared scenario catalog. The server suite (`scenarios.test.ts`) and the live editor suite
// (`scenarios-editor.test.ts`) run every entry.

import { ERROR_SCENARIOS } from './errors.ts';
import { PLATFORM_SCENARIOS } from './platform.ts';
import { PRESENTATION_SCENARIOS } from './presentation.ts';
import { FORMAT_SCENARIOS } from './format.ts';
import { SHAPE_SCENARIOS } from './shapes.ts';
import { SLIDE_SCENARIOS } from './slides.ts';
import { TABLE_SCENARIOS } from './tables.ts';
import { TEXT_SCENARIOS } from './text.ts';
import type { Scenario } from './types.ts';

export * from './types.ts';

export const SCENARIOS: readonly Scenario[] = [
  ...PRESENTATION_SCENARIOS,
  ...SLIDE_SCENARIOS,
  ...SHAPE_SCENARIOS,
  ...FORMAT_SCENARIOS,
  ...TEXT_SCENARIOS,
  ...TABLE_SCENARIOS,
  ...PLATFORM_SCENARIOS,
  ...ERROR_SCENARIOS,
];

/** A scenario by name. */
export function scenario(name: string): Scenario {
  const found = SCENARIOS.find((s) => s.name === name);
  if (!found) throw new Error(`no scenario ${name}`);
  return found;
}
