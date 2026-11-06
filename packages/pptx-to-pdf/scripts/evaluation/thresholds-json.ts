/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The JSON form of the fidelity thresholds (`scripts/fidelity/thresholds.ts`), for `worker.py`
// and `eigenpal/pptx-eval`. The limits live in one TypeScript file; this
// file is generated from it and checked in, and a test fails when the two disagree.
//
// Regenerate: `bun packages/pptx-to-pdf/scripts/evaluation/thresholds-json.ts --write`

import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { COMPARISON_B, THRESHOLDS } from '../../../../scripts/fidelity/thresholds.ts';
import { PROTOCOL } from './settings.ts';

/** Where the generated file lives, beside `worker.py`. */
export const THRESHOLDS_JSON = join(import.meta.dir, 'thresholds.json');

/** The thresholds as the JSON that `worker.py` and `pptx-eval` read. */
export function thresholdsJson(): string {
  const value = {
    protocol: PROTOCOL,
    source: 'scripts/fidelity/thresholds.ts',
    comparisonA: THRESHOLDS,
    comparisonB: COMPARISON_B,
  };
  return `${JSON.stringify(value, null, 2)}\n`;
}

if (import.meta.main) {
  if (process.argv.includes('--write')) writeFileSync(THRESHOLDS_JSON, thresholdsJson());
  else process.stdout.write(thresholdsJson());
}
