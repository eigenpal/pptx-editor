/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The paint budget: the deadline, the content bytes, and the costs that a file controls (the
// number of page resources and of gradient functions).

import { describe, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { moveTo, PDFDict, PDFDocument } from 'pdf-lib';
import type { ResolvedFill } from '@pptx-editor.dev/core/cascade';
import { buildPptx, shape } from '../../../../scripts/fixtures/builder.ts';
import { MAX_OUTPUT_BYTES, Work } from '../options.ts';
import { PageOps, SlideDiagnostics } from '../page-ops.ts';
import { paintFill } from '../paint-fill.ts';
import { exportPdf, PdfWorkLimitError } from '../index.ts';

const FIXTURES = join(import.meta.dir, '../../../../fixtures');
const BEST = { fidelityPolicy: 'best-effort' } as const;
const IN = 914_400;

const limitCode = (run: () => void): string | undefined => {
  try {
    run();
  } catch (error) {
    if (error instanceof PdfWorkLimitError) return error.code;
    throw error;
  }
  return undefined;
};

/** A page of a new document, painted under `work`. */
async function pageOf(work: Work): Promise<PageOps> {
  const doc = await PDFDocument.create();
  return new PageOps(
    doc.addPage([960, 540]),
    {
      doc,
      pkg: undefined as never,
      work,
      faces: new Map(),
      fonts: new Map(),
      images: new Map(),
      gradients: new Map(),
    },
    new SlideDiagnostics(0, 0)
  );
}

const full = (): Work => {
  const work = new Work(new AbortController().signal, Infinity);
  work.reserveContent(MAX_OUTPUT_BYTES - 1_000);
  return work;
};

describe('the paint budget', () => {
  test('counting paint work checks the deadline', () => {
    const work = new Work(new AbortController().signal, Date.now() - 1);
    expect(limitCode(() => work.count(1))).toBe('timeout');
  });

  test('content bytes past the budget stop the paint, and the budget grows with the pages', () => {
    const work = new Work(new AbortController().signal, Infinity);
    work.reserveContent(MAX_OUTPUT_BYTES);
    expect(limitCode(() => work.reserveContent(1))).toBe('operations');
    const many = new Work(new AbortController().signal, Infinity);
    // 1000 pages at 192 KiB each: about three times the one-page budget.
    many.allowPages(1000);
    many.reserveContent(MAX_OUTPUT_BYTES);
    expect(limitCode(() => many.reserveContent(1))).toBeUndefined();
  });

  test('page resources are named by a counter, never by pdf-lib uniqueKey', async () => {
    // `uniqueKey` copies every name of the dictionary on each call: the square of the count.
    const spy = spyOn(PDFDict.prototype, 'uniqueKey');
    try {
      for (const name of ['effects', 'gradients', 'pictures', 'text-runs', 'tables'])
        await exportPdf(readFileSync(join(FIXTURES, `${name}.pptx`)), BEST);
      expect(spy).not.toHaveBeenCalled();
    } finally {
      spy.mockRestore();
    }
  }, 60_000);

  test('a rectangular path gradient writes its stop functions once', async () => {
    const stops =
      '<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>' +
      '<a:gs pos="50000"><a:srgbClr val="00FF00"><a:alpha val="50000"/></a:srgbClr></a:gs>' +
      '<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>';
    const fill =
      `<a:gradFill><a:gsLst>${stops}</a:gsLst><a:path path="rect">` +
      '<a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>';
    const result = await exportPdf(
      buildPptx({
        slides: [{ layout: 'blank', shapes: shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN, fill }) }],
      }),
      BEST
    );
    const text = new TextDecoder('latin1').decode(result.bytes);
    const occurrences = (needle: string) => text.split(needle).length - 1;
    // Four trapezoids, each with a color shading and a soft-mask shading...
    expect(occurrences('/ShadingType 2')).toBe(8);
    // ...and one color function and one alpha function for all of them.
    expect(occurrences('/FunctionType 3')).toBe(2);
  });

  test('operators and gradient functions are charged as content', async () => {
    const operators = await pageOf(full());
    const many = Array.from({ length: 200 }, (_, i) => moveTo(i * 1000.25, i * 2000.5));
    expect(limitCode(() => operators.push(...many))).toBe('operations');
    const gradient = await pageOf(full());
    const fill = {
      type: 'linear',
      angle: 0,
      scaled: false,
      rotWithShape: true,
      stops: Array.from({ length: 64 }, (_, i) => ({
        pos: i / 63,
        color: { r: i * 4, g: 0, b: 255 - i * 4, a: 1 },
      })),
    } as unknown as ResolvedFill;
    let code: string | undefined;
    try {
      await paintFill(
        fill,
        IN,
        IN,
        () => {},
        gradient,
        { sourcePart: 'ppt/slides/slide1.xml' },
        'shapes'
      );
    } catch (error) {
      code = error instanceof PdfWorkLimitError ? error.code : String(error);
    }
    expect(code).toBe('operations');
  });

  test('one gradient on many shapes writes its stop functions once', async () => {
    const fill =
      '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>' +
      '<a:gs pos="50000"><a:srgbClr val="00FF00"/></a:gs>' +
      '<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst>' +
      '<a:lin ang="0" scaled="0"/></a:gradFill>';
    const shapes = Array.from({ length: 20 }, (_, i) =>
      shape({ id: i + 2, x: (i % 5) * IN, y: Math.floor(i / 5) * IN, cx: IN / 2, cy: IN / 2, fill })
    ).join('');
    const result = await exportPdf(
      buildPptx({
        slides: [
          { layout: 'blank', shapes },
          { layout: 'blank', shapes },
        ],
      }),
      BEST
    );
    const text = new TextDecoder('latin1').decode(result.bytes);
    expect(text.split('/ShadingType 2').length - 1).toBe(40);
    expect(text.split('/FunctionType 3').length - 1).toBe(1);
  });

  test('the paint lets a caller abort every 64 items', async () => {
    const controller = new AbortController();
    const p = await pageOf(new Work(controller.signal, Infinity));
    for (let i = 1; i < 64; i += 1) await p.step();
    controller.abort(new PdfWorkLimitError('aborted'));
    let code: string | undefined;
    try {
      await p.step();
    } catch (error) {
      code = error instanceof PdfWorkLimitError ? error.code : String(error);
    }
    expect(code).toBe('aborted');
  });
});
