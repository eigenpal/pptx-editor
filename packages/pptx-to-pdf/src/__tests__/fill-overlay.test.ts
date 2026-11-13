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
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
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

  test('a shape gradient overlay draws as bands of the outline', async () => {
    const shapeGradient =
      '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>' +
      '<a:gs pos="100000"><a:srgbClr val="000080"/></a:gs></a:gsLst>' +
      '<a:path path="shape"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>' +
      '</a:gradFill>';
    const effects = `<a:effectLst><a:fillOverlay blend="mult">${shapeGradient}</a:fillOverlay></a:effectLst>`;
    const fill = solid('4472C4');
    const shapes = shape({
      id: 2,
      x: IN,
      y: IN,
      cx: 2 * IN,
      cy: IN,
      prst: 'ellipse',
      fill,
      effects,
    });
    const result = await exportPdf(buildPptx({ slides: [{ layout: 'blank', shapes }] }), BEST);
    const pdf = await getDocument({ data: result.bytes.slice(), useSystemFonts: false }).promise;
    const { fnArray } = await (await pdf.getPage(1)).getOperatorList();
    const count = (op: number) => fnArray.filter((f) => f === op).length;
    // Bands of the ellipse are many paths; the rectangle construction would be shadings.
    expect(count(OPS.shadingFill)).toBe(0);
    expect(count(OPS.constructPath)).toBeGreaterThan(50);
  });
});
