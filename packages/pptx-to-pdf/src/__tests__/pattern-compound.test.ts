/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Pattern fills and compound lines, drawn with the SVG paint's bitmaps and bands.

import { describe, expect, test } from 'bun:test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { decodePDFRawStream, PDFDocument, PDFName, PDFRawStream } from 'pdf-lib';
import { buildPptx, shape } from '../../../../scripts/fixtures/builder.ts';
import { exportPdf, type PdfExportResult } from '../index.ts';

const BEST = { fidelityPolicy: 'best-effort' } as const;
const IN = 914_400;

const codes = (result: PdfExportResult): string[] => result.diagnostics.map((d) => d.code);

const patt = (prst: string): string =>
  `<a:pattFill prst="${prst}"><a:fgClr><a:srgbClr val="C00000"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>`;

/** The tiling patterns of a PDF: their matrix and decoded content. */
async function patterns(bytes: Uint8Array): Promise<{ matrix: number[]; content: string }[]> {
  const doc = await PDFDocument.load(bytes);
  const out = [];
  for (const [, obj] of doc.context.enumerateIndirectObjects()) {
    if (!(obj instanceof PDFRawStream) || obj.dict.get(PDFName.of('PatternType')) === undefined)
      continue;
    const raw = obj.dict.has(PDFName.of('Filter'))
      ? decodePDFRawStream(obj).decode()
      : obj.contents;
    const matrix = String(obj.dict.lookup(PDFName.of('Matrix')))
      .replace(/[[\]]/g, ' ')
      .trim()
      .split(/\s+/)
      .map(Number);
    out.push({ matrix, content: new TextDecoder('latin1').decode(raw) });
  }
  return out;
}

/** The number of stroke operators on page 1, through pdf.js. */
async function strokes(result: PdfExportResult): Promise<number> {
  const pdf = await getDocument({ data: result.bytes.slice(), useSystemFonts: false }).promise;
  const { fnArray, argsArray } = await (await pdf.getPage(1)).getOperatorList();
  // pdf.js 5 folds a path and its paint operator into one `constructPath`.
  return fnArray
    .map((op, i) => (op === OPS.constructPath ? ((argsArray[i] as [number])[0] ?? op) : op))
    .filter((op) => op === OPS.stroke).length;
}

describe('pattern fills', () => {
  test('a preset draws as one tile on the slide grid, also on a turned shape', async () => {
    const result = await exportPdf(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes:
              shape({
                id: 2,
                x: IN,
                y: IN,
                cx: 2 * IN,
                cy: IN,
                rot: 2_700_000,
                fill: patt('horz'),
              }) + shape({ id: 3, x: 4 * IN, y: IN, cx: 2 * IN, cy: IN, fill: patt('horz') }),
          },
        ],
      })
    );
    expect(codes(result)).toEqual([]);
    // One pattern per preset and colors on a page, whatever the shape's transform.
    const found = await patterns(result.bytes);
    expect(found).toHaveLength(1);
    // Pattern space is the page's, y down from the slide's top-left corner, 1 pt per pixel.
    expect(found[0].matrix).toEqual([1, 0, 0, -1, 0, 540]);
    // `horz`: the background tile, then one foreground run along the top row.
    expect(found[0].content).toContain('0 0 m\n8 0 l\n8 1 l\n0 1 l\nh');
  });

  test('a pattern line is still reported as not drawn', async () => {
    const result = await exportPdf(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: shape({
              id: 2,
              x: IN,
              y: IN,
              cx: IN,
              cy: IN,
              line: `<a:ln w="12700">${patt('pct50')}</a:ln>`,
            }),
          },
        ],
      }),
      BEST
    );
    expect(codes(result)).toContain('fill.pattern');
  });
});

describe('compound lines', () => {
  const line = (cmpd: string, dash = ''): string =>
    `<a:ln w="76200" cmpd="${cmpd}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>${dash}</a:ln>`;
  const of = (ln: string): Promise<PdfExportResult> =>
    exportPdf(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: 0, prst: 'line', line: ln }),
          },
        ],
      }),
      BEST
    );

  test('a solid compound line draws each band', async () => {
    const result = await of(line('thickThin'));
    expect(codes(result)).not.toContain('line.compound');
    expect(await strokes(result)).toBe(2);
    expect(await strokes(await of(line('tri')))).toBe(3);
  });

  test('a dashed compound line draws as one line, with a diagnostic', async () => {
    const result = await of(line('thickThin', '<a:prstDash val="dash"/>'));
    expect(codes(result)).toContain('line.compound');
    expect(await strokes(result)).toBe(1);
  });
});
