/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Warped text (`a:prstTxWarp`) in the PDF: the text outline strokes the warped glyphs, and a
// picture bullet draws, as in the SVG paint.

import { describe, expect, test } from 'bun:test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { PNG } from 'pngjs';
import { buildPptx, para, run, shape, solid } from '../../../../scripts/fixtures/builder.ts';
import { exportPdf, type PdfExportResult } from '../index.ts';

const IN = 914_400;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const BEST = { fidelityPolicy: 'best-effort' } as const;

function png(): Uint8Array<ArrayBuffer> {
  const image = new PNG({ width: 4, height: 4 });
  image.data.fill(200);
  return new Uint8Array(PNG.sync.write(image));
}

function deck(pPr: string, rPr: string): Uint8Array {
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          x: IN,
          y: IN,
          cx: 4 * IN,
          cy: 2 * IN,
          bodyPr:
            '<a:bodyPr wrap="none" rtlCol="0"><a:prstTxWarp prst="textWave1"><a:avLst/></a:prstTxWarp><a:noAutofit/></a:bodyPr>',
          text: para(
            run(
              'Warp',
              'sz="3600"',
              `${rPr}${solid('1F4E79')}<a:latin typeface="Liberation Sans"/>`
            ),
            pPr
          ),
        }),
        rels: [['rIdBu', R_IMAGE, '../media/image1.png']],
      },
    ],
    media: [{ name: 'image1.png', bytes: png(), contentType: 'image/png' }],
  });
}

/** The operators of page 1, with each path's paint operator in place of `constructPath`. */
async function ops(result: PdfExportResult) {
  const pdf = await getDocument({ data: result.bytes.slice(), useSystemFonts: false }).promise;
  const { fnArray, argsArray } = await (await pdf.getPage(1)).getOperatorList();
  return {
    ops: fnArray.map((op, i) =>
      op === OPS.constructPath ? ((argsArray[i] as [number])[0] ?? op) : op
    ),
    strokeColors: fnArray
      .map((op, i) => (op === OPS.setStrokeRGBColor ? JSON.stringify(argsArray[i]) : null))
      .filter((c): c is string => c !== null),
  };
}

describe('warped text in the PDF', () => {
  test('a solid text outline strokes the warped glyphs', async () => {
    const outline = '<a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>';
    const result = await exportPdf(deck('', outline), BEST);
    const page = await ops(result);
    expect(page.ops).toContain(OPS.stroke);
    expect(page.strokeColors).toContain('["#ff0000"]');
    expect(result.diagnostics.map((d) => d.code)).not.toContain('text.outline');
  });

  test('a picture bullet draws on warped text', async () => {
    const bullet =
      '<a:pPr marL="342900" indent="-342900"><a:buBlip><a:blip r:embed="rIdBu"/></a:buBlip></a:pPr>';
    const page = await ops(await exportPdf(deck(bullet, ''), BEST));
    expect(page.ops).toContain(OPS.paintImageXObject);
    const plain = await ops(await exportPdf(deck('', ''), BEST));
    expect(plain.ops).not.toContain(OPS.paintImageXObject);
  });
});
