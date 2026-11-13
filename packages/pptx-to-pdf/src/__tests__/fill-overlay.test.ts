/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `a:fillOverlay` in the PDF export: the overlay fill blends over the shape's fill in an isolated
// transparency group, with the blend mode of the SVG paint's `mix-blend-mode`.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { PDFBool, PDFDict, PDFDocument, PDFName, PDFRawStream } from 'pdf-lib';
import { buildPptx, shape, solid } from '../../../../scripts/fixtures/builder.ts';
import { exportPdf } from '../index.ts';

const FIXTURES = join(import.meta.dir, '../../../../fixtures');
const BEST = { fidelityPolicy: 'best-effort' } as const;
const IN = 914_400;

/** The blend modes of the PDF's graphics states, and how many isolated groups it holds. */
async function blendsOf(bytes: Uint8Array): Promise<{ blends: string[]; isolated: number }> {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false });
  const blends = new Set<string>();
  let isolated = 0;
  for (const [, obj] of doc.context.enumerateIndirectObjects()) {
    const dict = obj instanceof PDFRawStream ? obj.dict : obj instanceof PDFDict ? obj : null;
    if (!dict) continue;
    const bm = dict.lookup(PDFName.of('BM'));
    if (bm instanceof PDFName) blends.add(bm.decodeText());
    const group = dict.lookup(PDFName.of('Group'));
    if (group instanceof PDFDict && group.lookup(PDFName.of('I')) === PDFBool.True) isolated += 1;
  }
  return { blends: [...blends].sort(), isolated };
}

describe('fill overlays', () => {
  test('each overlay of the fixture blends in an isolated group', async () => {
    const deck = readFileSync(join(FIXTURES, 'fill-fidelity.pptx'));
    const result = await exportPdf(deck, BEST);
    const { blends, isolated } = await blendsOf(result.bytes);
    // The fixture's overlays: multiply, screen, and darken.
    expect(blends).toEqual(['Darken', 'Multiply', 'Screen']);
    expect(isolated).toBeGreaterThanOrEqual(3);
    expect(result.diagnostics.map((d) => d.code)).not.toContain('effect.fill-overlay');
  });

  test('a shape without a fill draws no overlay', async () => {
    const effects =
      '<a:effectLst><a:fillOverlay blend="mult"><a:solidFill><a:srgbClr val="808080"/>' +
      '</a:solidFill></a:fillOverlay></a:effectLst>';
    const of = async (fill: string) =>
      blendsOf(
        (
          await exportPdf(
            buildPptx({
              slides: [
                {
                  layout: 'blank',
                  shapes: shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN, fill, effects }),
                },
              ],
            }),
            BEST
          )
        ).bytes
      );
    expect((await of(solid('4472C4'))).blends).toEqual(['Multiply']);
    expect((await of('<a:noFill/>')).blends).toEqual([]);
  });
});
