/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Soft edges: the body paints as one group through a mask that fades toward the outline.

import { describe, expect, test } from 'bun:test';
import { decodePDFRawStream, PDFDict, PDFDocument, PDFName, PDFRawStream } from 'pdf-lib';
import { blurSigma } from '@pptx-editor.dev/core/output';
import { softEdgeSpread } from '@pptx-editor.dev/pro/drawing';
import { buildPptx, shape, solid } from '../../../../scripts/fixtures/builder.ts';
import { softEdgeBands } from '../paint-effects.ts';
import { exportPdf, type PdfExportResult } from '../index.ts';

const BEST = { fidelityPolicy: 'best-effort' } as const;
const IN = 914_400;
const PT = 12_700;

/** The decoded content of every luminosity soft mask group (a form with a DeviceGray group). */
async function softMasks(bytes: Uint8Array): Promise<string[]> {
  const doc = await PDFDocument.load(bytes);
  const out: string[] = [];
  for (const [, obj] of doc.context.enumerateIndirectObjects()) {
    if (!(obj instanceof PDFRawStream)) continue;
    const group = obj.dict.lookupMaybe(PDFName.of('Group'), PDFDict);
    if (group?.get(PDFName.of('CS')) !== PDFName.of('DeviceGray')) continue;
    const raw = obj.dict.has(PDFName.of('Filter'))
      ? decodePDFRawStream(obj).decode()
      : obj.contents;
    out.push(new TextDecoder('latin1').decode(raw));
  }
  return out;
}

/** The decoded content streams of a PDF, the page's and every form's (pictures left out). */
async function contentText(bytes: Uint8Array): Promise<string> {
  const doc = await PDFDocument.load(bytes);
  let text = '';
  for (const [, obj] of doc.context.enumerateIndirectObjects()) {
    if (!(obj instanceof PDFRawStream)) continue;
    if (obj.dict.get(PDFName.of('Subtype')) === PDFName.of('Image')) continue;
    const raw = obj.dict.has(PDFName.of('Filter'))
      ? decodePDFRawStream(obj).decode()
      : obj.contents;
    text += new TextDecoder('latin1').decode(raw);
  }
  return text;
}

const soft = (cx: number, cy: number, rad: number): Promise<PdfExportResult> =>
  exportPdf(
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            x: IN,
            y: IN,
            cx,
            cy,
            fill: solid('4472C4'),
            effects: `<a:effectLst><a:softEdge rad="${rad}"/></a:effectLst>`,
          }),
        },
      ],
    }),
    BEST
  );

describe('soft edge bands', () => {
  test('the alpha rises from clear on the outline to opaque past the blur', () => {
    const { spread, sigma } = softEdgeSpread(10 * PT);
    const { reach, alpha } = softEdgeBands(-spread, sigma);
    expect(reach.length).toBeGreaterThan(8);
    expect(reach.length).toBeLessThanOrEqual(48);
    expect(reach[reach.length - 1]).toBeCloseTo(-spread + 3 * blurSigma(10 * PT));
    for (let k = 1; k < alpha.length; k += 1) expect(alpha[k]).toBeGreaterThanOrEqual(alpha[k - 1]);
    expect(alpha[0]).toBeLessThan(0.05);
    expect(alpha[alpha.length - 1]).toBeGreaterThan(0.95);
    // The half level is at the inset, 0.84 of the radius inside the outline.
    const half = reach.findIndex((_, k) => alpha[k] >= 0.5);
    expect(reach[half]).toBeGreaterThanOrEqual(-spread);
  });

  test('a blur under half a point is a hard step at the inset', () => {
    const { reach, alpha } = softEdgeBands(10 * PT, 0);
    reach.forEach((r, k) => {
      const middle = ((k ? reach[k - 1] : 0) + r) / 2;
      expect(alpha[k]).toBe(middle >= 10 * PT ? 1 : 0);
    });
  });
});

describe('soft edge paint', () => {
  test('a filled shape paints through a mask of its outline bands', async () => {
    const result = await soft(3 * IN, 2 * IN, 10 * PT);
    const masks = await softMasks(result.bytes);
    expect(masks).toHaveLength(1);
    const { spread, sigma } = softEdgeSpread(10 * PT);
    const { reach } = softEdgeBands(-spread, sigma);
    // One white fill, then one stroke per band, widest first.
    expect(masks[0].match(/\nS\n/g)).toHaveLength(reach.length);
    const widths = [...masks[0].matchAll(/([\d.]+) w/g)].map((m) => Number(m[1]));
    expect(widths[0]).toBeCloseTo(2 * reach[reach.length - 1], 0);
    expect(widths[widths.length - 1]).toBeCloseTo(2 * reach[0], 0);
    expect(await contentText(result.bytes)).toContain('0.4470588');
    expect(result.diagnostics.find((d) => d.code === 'effect.soft-edge')?.message).toContain(
      'distance to the outline'
    );
  });

  test('a soft edge past half the narrower side paints nothing', async () => {
    const result = await soft(3 * IN, 10 * PT, 20 * PT);
    expect(await softMasks(result.bytes)).toHaveLength(0);
    // The fill color (#4472C4) is set nowhere, in the page or in a form.
    const text = await contentText(result.bytes);
    expect(text).not.toContain('0.4470588');
  });
});
