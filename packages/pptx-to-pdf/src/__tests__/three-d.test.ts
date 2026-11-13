/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// 3-D shapes in the PDF: degenerate shapes write no NaN, and the shadow of a turned face turns
// with it, as in the SVG paint.

import { describe, expect, test } from 'bun:test';
import { decodePDFRawStream, PDFDocument, PDFName, PDFRawStream } from 'pdf-lib';
import { buildPptx, shape, solid } from '../../../../scripts/fixtures/builder.ts';
import { exportPdf } from '../index.ts';

const IN = 914_400;
const BEST = { fidelityPolicy: 'best-effort' } as const;

const deckOf = (o: { cx: number; cy: number; prst?: string; effects: string }) =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({ id: 2, x: IN, y: IN, fill: solid('4472C4'), ...o }),
      },
    ],
  });

/** The decoded text of every stream of a PDF, and which of them are forms. */
async function streamsOf(pdf: Uint8Array): Promise<{ text: string; form: boolean }[]> {
  const doc = await PDFDocument.load(pdf);
  const out: { text: string; form: boolean }[] = [];
  for (const [, obj] of doc.context.enumerateIndirectObjects()) {
    if (!(obj instanceof PDFRawStream)) continue;
    const raw = obj.dict.has(PDFName.of('Filter'))
      ? decodePDFRawStream(obj).decode()
      : obj.contents;
    out.push({
      text: new TextDecoder('latin1').decode(raw),
      form: obj.dict.get(PDFName.of('Subtype')) === PDFName.of('Form'),
    });
  }
  return out;
}

/** The height of the path points that the forms trace (their `m` and `l` operands). */
function formPathHeight(streams: readonly { text: string; form: boolean }[]): number {
  const ys: number[] = [];
  for (const s of streams) {
    if (!s.form) continue;
    const tokens = s.text.split(/\s+/);
    tokens.forEach((t, i) => {
      if (t === 'm' || t === 'l') ys.push(Number(tokens[i - 1]));
    });
  }
  return Math.max(...ys) - Math.min(...ys);
}

const SCENE = (rot: string) =>
  `<a:scene3d><a:camera prst="orthographicFront">${rot}</a:camera>` +
  '<a:lightRig rig="threePt" dir="t"/></a:scene3d>';

describe('3-D shapes in the PDF', () => {
  test('a zero-height shape with a bevel writes no NaN', async () => {
    const result = await exportPdf(
      deckOf({
        cx: IN,
        cy: 0,
        prst: 'triangle',
        effects: `${SCENE('')}<a:sp3d><a:bevelT w="76200" h="76200" prst="hardEdge"/></a:sp3d>`,
      }),
      BEST
    );
    for (const s of await streamsOf(result.bytes)) expect(s.text).not.toContain('NaN');
  });

  test('opaque bevel strips have an outline of their own color, as in the SVG paint', async () => {
    const result = await exportPdf(
      deckOf({
        cx: 2 * IN,
        cy: IN,
        effects: `${SCENE('')}<a:sp3d><a:bevelT w="76200" h="76200" prst="angle"/></a:sp3d>`,
      }),
      BEST
    );
    const text = (await streamsOf(result.bytes)).map((s) => s.text).join('\n');
    // Each strip fills and strokes (`B`) with the same color, so neighbors leave no seam.
    const strips = [...text.matchAll(/([\d.]+ [\d.]+ [\d.]+) rg\s+\1 RG\s+[\d.]+ w\s+1 j\s/g)];
    expect(strips.length).toBeGreaterThanOrEqual(4);
    expect(text).toMatch(/\bB\s/);
  });

  test('the shadow of a turned face turns with it', async () => {
    const shadow =
      '<a:effectLst><a:outerShdw dist="0" dir="0" algn="ctr" rotWithShape="0">' +
      '<a:prstClr val="black"/></a:outerShdw></a:effectLst>';
    const height = async (rot: string): Promise<number> => {
      const deck = deckOf({ cx: 2 * IN, cy: IN / 2, effects: shadow + SCENE(rot) });
      return formPathHeight(await streamsOf((await exportPdf(deck, BEST)).bytes));
    };
    const flat = await height('<a:rot lat="0" lon="0" rev="0"/>');
    // A turn of 45 degrees about the view axis: the 2 by 0.5 in face is 1.77 in high.
    const turned = await height('<a:rot lat="0" lon="0" rev="2700000"/>');
    expect(turned / flat).toBeCloseTo((Math.SQRT1_2 * 2.5) / 0.5, 1);
  });
});
