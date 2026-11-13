/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Blur: the body fades to half its alpha at the painted edge, and a grown blur adds a halo of
// layered rings outside the fill.

import { describe, expect, test } from 'bun:test';
import { decodePDFRawStream, PDFDict, PDFDocument, PDFName, PDFRawStream } from 'pdf-lib';
import { blurSigma } from '@pptx-editor.dev/core/output';
import { buildPptx, shape, solid } from '../../../../scripts/fixtures/builder.ts';
import { softEdgeBands } from '../paint-effects.ts';
import { exportPdf, type PdfExportResult } from '../index.ts';

const BEST = { fidelityPolicy: 'best-effort' } as const;
const IN = 914_400;
const PT = 12_700;

/** The decoded streams of a PDF: luminosity mask groups apart from the other forms. */
async function streams(bytes: Uint8Array): Promise<{ masks: string[]; groups: string[] }> {
  const doc = await PDFDocument.load(bytes);
  const masks: string[] = [];
  const groups: string[] = [];
  for (const [, obj] of doc.context.enumerateIndirectObjects()) {
    if (!(obj instanceof PDFRawStream)) continue;
    const group = obj.dict.lookupMaybe(PDFName.of('Group'), PDFDict);
    if (!group) continue;
    const raw = obj.dict.has(PDFName.of('Filter'))
      ? decodePDFRawStream(obj).decode()
      : obj.contents;
    const text = new TextDecoder('latin1').decode(raw);
    if (group.get(PDFName.of('CS')) === PDFName.of('DeviceGray')) masks.push(text);
    else groups.push(text);
  }
  return { masks, groups };
}

const blurred = (blur: string): Promise<PdfExportResult> =>
  exportPdf(
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 3 * IN,
            cy: 2 * IN,
            fill: solid('4472C4'),
            effects: `<a:effectLst>${blur}</a:effectLst>`,
          }),
        },
      ],
    }),
    BEST
  );

/** The stroke grays of a mask, in order. */
const grays = (mask: string): number[] =>
  [...mask.matchAll(/([\d.]+) G\n/g)].map((m) => Number(m[1]));

describe('blur paint', () => {
  test('the body fades to half its alpha at the edge, as the blurred fill', async () => {
    const result = await blurred('<a:blur rad="127000" grow="0"/>');
    const { masks } = await streams(result.bytes);
    expect(masks).toHaveLength(1);
    const { reach, alpha } = softEdgeBands(0, blurSigma(10 * PT));
    expect(reach[reach.length - 1]).toBeCloseTo(3 * blurSigma(10 * PT));
    // One stroke per band, widest first: the outermost band is the last, near half.
    const g = grays(masks[0]);
    expect(g).toHaveLength(reach.length);
    expect(g[g.length - 1]).toBeCloseTo(alpha[0], 3);
    expect(alpha[0]).toBeGreaterThan(0.5);
    expect(alpha[0]).toBeLessThan(0.6);
    expect(result.diagnostics.find((d) => d.code === 'effect.blur')?.message).toContain(
      'edge fade'
    );
  });

  test('a grown blur adds a halo in the fill color outside the fill', async () => {
    const flat = await streams((await blurred('<a:blur rad="127000" grow="0"/>')).bytes);
    const grown = await streams((await blurred('<a:blur rad="127000"/>')).bytes);
    expect(grown.masks).toHaveLength(1);
    // The halo layers: transparency groups in the fill color (#4472C4).
    const halo = (groups: string[]) =>
      groups.filter((s) => /^0\.2666\d* 0\.4470\d* 0\.7686\d* rg/.test(s)).length;
    expect(halo(grown.groups)).toBeGreaterThan(halo(flat.groups) + 3);
  });

  test('a blur under half a point draws the shape as it is', async () => {
    const { masks } = await streams((await blurred('<a:blur rad="6000"/>')).bytes);
    expect(masks).toHaveLength(0);
  });
});
