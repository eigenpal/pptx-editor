/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A synthetic italic: an italic run in a family with no italic face leans the regular face, as
// the browser leans it in the editor's paint. Synthetic bold is still reported.
import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { fontConfiguration } from '../../../../scripts/fixtures/font-configuration.ts';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { exportPdf } from '../index.ts';

const FAMILY = 'STIX Two Math';
const fonts = [
  {
    family: FAMILY,
    origin: 'test',
    bold: false,
    italic: false,
    bytes: new Uint8Array(
      readFileSync(new URL('../../../../fixtures/fonts/STIXTwoMath-Regular.otf', import.meta.url))
    ),
  },
];

function deck(attrs: string) {
  return buildPptx({
    slides: [
      {
        shapes: shape({
          id: 2,
          x: 914400,
          y: 914400,
          cx: 7315200,
          cy: 1828800,
          text: para(run('Lean', `sz="2400" ${attrs}`, `<a:latin typeface="${FAMILY}"/>`)),
        }),
      },
    ],
  });
}

/** The shear (`c`) of each text matrix on page 1, and the export's diagnostic codes. */
async function exported(attrs: string): Promise<{ shears: number[]; codes: string[] }> {
  const result = await exportPdf(deck(attrs), {
    fonts: fontConfiguration(fonts),
    fidelityPolicy: 'best-effort',
  });
  const pdf = await getDocument({ data: result.bytes.slice(), useSystemFonts: false }).promise;
  try {
    const { fnArray, argsArray } = await (await pdf.getPage(1)).getOperatorList();
    const shears = fnArray.flatMap((fn, i) =>
      fn === OPS.setTextMatrix ? [Number((argsArray[i] as [ArrayLike<number>])[0][2])] : []
    );
    return { shears, codes: result.diagnostics.map((d) => d.code) };
  } finally {
    await pdf.loadingTask.destroy();
  }
}

test('a synthetic italic leans its glyphs by a shear of 0.25', async () => {
  const { shears, codes } = await exported('i="1"');
  expect(shears.length).toBeGreaterThan(0);
  for (const c of shears) expect(c).toBeCloseTo(0.25, 6);
  expect(codes).not.toContain('text.synthetic-style');
});

test('upright text has no shear, and a synthetic bold is reported', async () => {
  const upright = await exported('');
  expect(upright.shears.length).toBeGreaterThan(0);
  for (const c of upright.shears) expect(c).toBe(0);
  const bold = await exported('b="1"');
  for (const c of bold.shears) expect(c).toBe(0);
  expect(bold.codes).toContain('text.synthetic-style');
});

test('a synthetic bold italic leans and reports the bold', async () => {
  const { shears, codes } = await exported('b="1" i="1"');
  expect(shears.length).toBeGreaterThan(0);
  for (const c of shears) expect(c).toBeCloseTo(0.25, 6);
  expect(codes).toContain('text.synthetic-style');
});
