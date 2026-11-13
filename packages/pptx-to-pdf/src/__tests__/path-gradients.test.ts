/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Rectangle and shape path gradients, drawn with the SVG paint's construction and bands.

import { describe, expect, test } from 'bun:test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { buildPptx, shape, solid } from '../../../../scripts/fixtures/builder.ts';
import { shapeGradientBands } from '@pptx-editor.dev/core/output';
import { exportPdf, type PdfExportResult } from '../index.ts';

const BEST = { fidelityPolicy: 'best-effort' } as const;
const IN = 914_400;

/** The operators of page 1 and its fill colors (`"#rrggbb"`), through pdf.js. */
async function page1(result: PdfExportResult): Promise<{ ops: number[]; colors: string[] }> {
  const pdf = await getDocument({ data: result.bytes.slice(), useSystemFonts: false }).promise;
  const { fnArray, argsArray } = await (await pdf.getPage(1)).getOperatorList();
  // pdf.js 5 folds a path and its paint operator into one `constructPath`.
  const ops = fnArray.map((op, i) =>
    op === OPS.constructPath ? ((argsArray[i] as [number])[0] ?? op) : op
  );
  const colors = fnArray
    .map((op, i) => (op === OPS.setFillRGBColor ? String((argsArray[i] as string[])[0]) : null))
    .filter((c): c is string => c !== null);
  return { ops, colors };
}

const count = (ops: number[], op: number): number => ops.filter((o) => o === op).length;
const codes = (result: PdfExportResult): string[] => result.diagnostics.map((d) => d.code);

describe('path gradients', () => {
  const pathFill = (path: string, end = ''): string =>
    '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>' +
    `<a:gs pos="100000"><a:srgbClr val="000080">${end}</a:srgbClr></a:gs></a:gsLst>` +
    `<a:path path="${path}"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>` +
    '</a:gradFill>';
  const triangle = (fill: string) =>
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, prst: 'triangle', fill }),
        },
      ],
    });

  test('a shape path gradient draws the bands of the outline, as the SVG paint', async () => {
    const result = await exportPdf(triangle(pathFill('shape')), BEST);
    expect(codes(result).filter((c) => c.startsWith('fill.'))).toEqual([]);
    const page = await page1(result);
    const stops = [
      { pos: 0, color: { r: 255, g: 255, b: 255, a: 1 } },
      { pos: 1, color: { r: 0, g: 0, b: 128, a: 1 } },
    ];
    const { bands } = shapeGradientBands(
      {
        type: 'path',
        path: 'shape',
        stops,
        fillToRect: { l: 0.5, t: 0.5, r: 0.5, b: 0.5 },
        rotWithShape: true,
      } as unknown as Parameters<typeof shapeGradientBands>[0],
      stops,
      2 * IN,
      IN
    );
    // One filled outline per band in place of the solid fill's one, and no rectangle construction.
    const plain = await page1(await exportPdf(triangle(solid('000080')), BEST));
    expect(count(page.ops, OPS.eoFill)).toBe(count(plain.ops, OPS.eoFill) - 1 + bands.length);
    expect(count(page.ops, OPS.shadingFill)).toBe(0);
    // Outside in: the bands after the background run from near the last stop to near the first.
    const colors = page.colors.slice(-bands.length);
    expect(colors[0]).toMatch(/^#0[0-9a-f]0[0-9a-f][0-9a-f]{2}$/);
    expect(colors[colors.length - 1]).toMatch(/^#f[0-9a-f]f[0-9a-f]f[0-9a-f]$/);
  });

  test('a shape path gradient with stop alpha paints its bands once, through a soft mask', async () => {
    const result = await exportPdf(triangle(pathFill('shape', '<a:alpha val="20000"/>')), BEST);
    const page = await page1(result);
    const pdf = new TextDecoder('latin1').decode(result.bytes);
    expect(pdf).toContain('/Luminosity');
    // The bands paint in one transparency group, so the mask applies once to all of them.
    expect(count(page.ops, OPS.paintFormXObjectBegin)).toBeGreaterThanOrEqual(1);
  });

  test('a rectangular path gradient with stop alpha masks the whole construction once', async () => {
    const result = await exportPdf(triangle(pathFill('rect', '<a:alpha val="20000"/>')), BEST);
    const page = await page1(result);
    const pdf = new TextDecoder('latin1').decode(result.bytes);
    expect(pdf).toContain('/Luminosity');
    expect(count(page.ops, OPS.paintFormXObjectBegin)).toBeGreaterThanOrEqual(1);
    expect(count(page.ops, OPS.shadingFill)).toBeGreaterThanOrEqual(4);
  });
});
