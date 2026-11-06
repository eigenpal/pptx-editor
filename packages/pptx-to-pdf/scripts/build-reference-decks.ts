/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Builds the decks in `src/__tests__/reference/` with the fixture builder, so their themes name
// only open fonts (Carlito). After a rebuild, export each reference PDF again from the new deck.
//
//   bun packages/pptx-to-pdf/scripts/build-reference-decks.ts

import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';

const IN = 914_400;
const REFERENCE = join(import.meta.dir, '../src/__tests__/reference');

/** Four text slides; slides 2 and 4 are hidden. */
const hiddenSlides = buildPptx({
  slides: [1, 2, 3, 4].map((n) => ({
    layout: 'blank' as const,
    hidden: n % 2 === 0,
    shapes: shape({ id: 2, x: IN, y: IN, cx: 6 * IN, cy: IN, text: para(run(`Slide ${n}`)) }),
  })),
});

writeFileSync(join(REFERENCE, 'hidden-slides.pptx'), hiddenSlides);
console.log('hidden-slides.pptx written');
