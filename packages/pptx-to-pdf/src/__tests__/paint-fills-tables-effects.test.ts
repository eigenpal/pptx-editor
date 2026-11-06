/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Gradient fills, tables, and outer shadows and glow, drawn from the engine's layout.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { PNG } from 'pngjs';
import { decodePDFRawStream, PDFDocument, PDFName, PDFRawStream } from 'pdf-lib';
import { buildPptx, para, run, shape, solid } from '../../../../scripts/fixtures/builder.ts';
import { apply } from '@pptx-editor.dev/core/layout';
import { blurLayers, innerShadowOffset, layers, phi, rings } from '../paint-effects.ts';
import { MAX_GRADIENT_STOPS } from '../paint-fill.ts';
import { exportPdf, type PdfExportResult } from '../index.ts';

const FIXTURES = join(import.meta.dir, '../../../../fixtures');
const deck = (name: string): Uint8Array => readFileSync(join(FIXTURES, `${name}.pptx`));
const BEST = { fidelityPolicy: 'best-effort' } as const;
const IN = 914_400;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

/** The operator names and the text of each page, through pdf.js. */
async function pages(result: PdfExportResult) {
  const pdf = await getDocument({ data: result.bytes.slice(), useSystemFonts: false }).promise;
  const out = [];
  for (let i = 1; i <= pdf.numPages; i += 1) {
    const page = await pdf.getPage(i);
    const { fnArray, argsArray } = await page.getOperatorList();
    // pdf.js 5 folds a path and its paint operator into one `constructPath`.
    const ops = fnArray.map((op, i) =>
      op === OPS.constructPath ? ((argsArray[i] as [number])[0] ?? op) : op
    );
    const content = await page.getTextContent();
    const words = content.items
      .map((item) => ('str' in item ? item.str : ''))
      .join(' ')
      .split(/\s+/)
      .filter((w) => w !== '');
    out.push({ ops, words });
  }
  return out;
}

const count = (ops: number[], op: number): number => ops.filter((o) => o === op).length;
const codes = (result: PdfExportResult): string[] => result.diagnostics.map((d) => d.code);

/** The fill colors of page 1, as pdf.js reports them (`"#rrggbb"`). */
async function fillColors(result: PdfExportResult): Promise<string> {
  const pdf = await getDocument({ data: result.bytes.slice(), useSystemFonts: false }).promise;
  const { fnArray, argsArray } = await (await pdf.getPage(1)).getOperatorList();
  return fnArray
    .map((op, i) => (op === OPS.setFillRGBColor ? JSON.stringify(argsArray[i]) : null))
    .filter((c): c is string => c !== null)
    .join();
}

/** The points of the path operators in every form XObject, rounded to 100 EMU and sorted. */
async function formPoints(bytes: Uint8Array): Promise<string[]> {
  const doc = await PDFDocument.load(bytes);
  const points: string[] = [];
  for (const [, obj] of doc.context.enumerateIndirectObjects()) {
    if (
      !(obj instanceof PDFRawStream) ||
      obj.dict.get(PDFName.of('Subtype')) !== PDFName.of('Form')
    )
      continue;
    const raw = obj.dict.has(PDFName.of('Filter'))
      ? decodePDFRawStream(obj).decode()
      : obj.contents;
    const tokens = new TextDecoder('latin1').decode(raw).split(/\s+/);
    tokens.forEach((token, i) => {
      const n = token === 'm' || token === 'l' ? 2 : token === 'c' ? 6 : 0;
      const args = tokens.slice(i - n, i).map((t) => Math.round(Number(t) / 100) * 100);
      for (let k = 0; k < n; k += 2) points.push(`${args[k]},${args[k + 1]}`);
    });
  }
  return points.sort();
}

const gradientShape = (stops: string) =>
  shape({
    id: 2,
    x: IN,
    y: IN,
    cx: 3 * IN,
    cy: IN,
    fill: `<a:gradFill><a:gsLst>${stops}</a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>`,
  });

describe('gradient fills', () => {
  test('linear and path gradients draw as shadings, with no diagnostic', async () => {
    const result = await exportPdf(deck('gradients'), BEST);
    expect(codes(result).filter((c) => c.startsWith('fill.'))).toEqual([]);
    const [page] = await pages(result);
    expect(count(page.ops, OPS.shadingFill)).toBeGreaterThanOrEqual(7);
  });

  test('a linear gradient runs along the axis of the SVG paint, through the box center', async () => {
    const stops =
      '<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>';
    const result = await exportPdf(
      buildPptx({ slides: [{ layout: 'blank', shapes: gradientShape(stops) }] }),
      BEST
    );
    // 0°, 3 × 1 inches: from the middle of the left edge to the middle of the right edge.
    const pdf = new TextDecoder('latin1').decode(result.bytes);
    expect(pdf).toContain(`/Coords [ 0 ${IN / 2} ${3 * IN} ${IN / 2} ]`);
  });

  test('a rectangular path gradient is four trapezoids over the last color', async () => {
    const fill =
      '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>' +
      '<a:gs pos="100000"><a:srgbClr val="000080"/></a:gs></a:gsLst>' +
      '<a:path path="rect"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>';
    const result = await exportPdf(
      buildPptx({
        slides: [
          { layout: 'blank', shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, fill }) },
        ],
      }),
      BEST
    );
    const [page] = await pages(result);
    expect(count(page.ops, OPS.shadingFill)).toBe(4);
  });

  test('stops with different alphas draw through a soft mask', async () => {
    const stops =
      '<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>' +
      '<a:gs pos="100000"><a:srgbClr val="0000FF"><a:alpha val="20000"/></a:srgbClr></a:gs>';
    const result = await exportPdf(
      buildPptx({ slides: [{ layout: 'blank', shapes: gradientShape(stops) }] }),
      BEST
    );
    const pdf = new TextDecoder('latin1').decode(result.bytes);
    expect(pdf).toContain('/Luminosity');
    expect(pdf).toContain('/ShadingType 2');
  });

  test('stops at one position make a hard step, not an invalid function', async () => {
    const stops =
      '<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>' +
      '<a:gs pos="50000"><a:srgbClr val="FF0000"/></a:gs>' +
      '<a:gs pos="50000"><a:srgbClr val="0000FF"/></a:gs>' +
      '<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>';
    const result = await exportPdf(
      buildPptx({ slides: [{ layout: 'blank', shapes: gradientShape(stops) }] }),
      BEST
    );
    const pdf = new TextDecoder('latin1').decode(result.bytes);
    // Two segments remain, so the stitching function has one bound, at the step.
    expect(pdf).toMatch(/\/Bounds \[ 0\.5 \]/);
  });
});

describe('gradient text', () => {
  test('clips to the glyphs and keeps the words in the text layer', async () => {
    const gradient =
      '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>' +
      '<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill>';
    const result = await exportPdf(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: shape({
              id: 2,
              x: IN,
              y: IN,
              cx: 6 * IN,
              cy: IN,
              text: para(run('Gradient words', '', gradient)),
            }),
          },
        ],
      }),
      BEST
    );
    expect(codes(result).filter((c) => c.startsWith('fill.'))).toEqual([]);
    const [page] = await pages(result);
    expect(page.words).toEqual(['Gradient', 'words']);
    expect(count(page.ops, OPS.shadingFill)).toBe(1);
  });
});

describe('line ends', () => {
  test('are reported only on open paths, the only paths that show them', async () => {
    const line =
      '<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:tailEnd type="triangle"/></a:ln>';
    const deckWith = (prst: string) =>
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, prst, line }),
          },
        ],
      });
    expect(codes(await exportPdf(deckWith('rect'), BEST))).not.toContain('line.ends');
    expect(codes(await exportPdf(deckWith('line'), BEST))).toContain('line.ends');
  });
});

describe('tables', () => {
  test('cells, borders, and text draw; each word is in the text layer once', async () => {
    const result = await exportPdf(deck('tables'), BEST);
    expect(codes(result).filter((c) => c.startsWith('frame.'))).toEqual([]);
    const [first] = await pages(result);
    expect(first.words.slice(0, 3)).toEqual(['Region', 'Q1', 'Q2']);
    expect(first.words.filter((w) => w === 'North')).toHaveLength(1);
    expect(count(first.ops, OPS.stroke)).toBeGreaterThan(0);
  });
});

describe('effects', () => {
  test('shadows, glow, inner shadows, and reflections draw as transparency groups', async () => {
    const result = await exportPdf(deck('effects'), BEST);
    const all = await pages(result);
    for (const page of all) expect(count(page.ops, OPS.beginGroup)).toBeGreaterThan(0);
    // The blur approximations are reported, never hidden.
    expect(codes(result)).toContain('effect.inner-shadow-blur');
    expect(codes(result)).toContain('effect.reflection-blur');
    // Reflections are drawn now: nothing in the deck is reported as not drawn.
    expect(codes(result)).not.toContain('effect.reflection');
  });

  test('a reflection draws through a luminosity soft mask', async () => {
    const text = new TextDecoder('latin1').decode((await exportPdf(deck('effects'), BEST)).bytes);
    expect(text).toMatch(/\/S\s*\/Luminosity/);
  });

  test('an inner shadow moves in slide space, whatever the rotation of the shape', () => {
    // A shape turned 90° about its center: local +x is slide +y.
    const m = [0, 1, -1, 0, 3 * IN, IN] as const;
    const local = innerShadowOffset(m, 50_800, 0);
    for (const [x, y] of [
      [0, 0],
      [IN, 2 * IN],
    ]) {
      const moved = apply(m, apply(local, x, y).x, apply(local, x, y).y);
      const base = apply(m, x, y);
      // The inverse moves opposite `dir`, so the shadow shows on the side `dir` points to.
      expect(moved.x - base.x).toBeCloseTo(-50_800, 3);
      expect(moved.y - base.y).toBeCloseTo(0, 3);
    }
  });

  test('the drawn inner shadow of a turned shape moves in slide space', async () => {
    // A square turned 90° about its center covers the same area as the square itself, so its
    // inner shadow forms (in slide space) hold the same points. A shadow offset in the shape's
    // local space would move down instead of right.
    const inner =
      '<a:effectLst><a:innerShdw blurRad="0" dist="190500" dir="0"><a:srgbClr val="000000">' +
      '<a:alpha val="50000"/></a:srgbClr></a:innerShdw></a:effectLst>';
    const square = (rot: number) =>
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
              rot,
              fill: '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>',
              effects: inner,
            }),
          },
        ],
      });
    const flat = await formPoints((await exportPdf(square(0), BEST)).bytes);
    const turned = await formPoints((await exportPdf(square(5_400_000), BEST)).bytes);
    expect(turned).toEqual(flat);
    // The shifted square's top-left corner: 0.2 in to the left of the shape's corner.
    expect(flat).toContain(`${IN - 190_500},${IN}`);
  });

  test('the PDF grows in a straight line with the number of shadowed shapes', async () => {
    const effects =
      '<a:effectLst><a:outerShdw blurRad="76200" dist="38100" dir="2700000"><a:srgbClr val="000000">' +
      '<a:alpha val="40000"/></a:srgbClr></a:outerShdw><a:glow rad="63500"><a:srgbClr val="FF0000">' +
      '<a:alpha val="40000"/></a:srgbClr></a:glow></a:effectLst>';
    const deckOf = (count: number) =>
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: Array.from({ length: count }, (_, i) =>
              shape({
                id: i + 2,
                x: (i % 8) * IN,
                y: Math.floor(i / 8) * (IN / 2),
                cx: IN / 2,
                cy: IN / 4,
                fill: '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>',
                effects,
              })
            ).join(''),
          },
        ],
      });
    const small = (await exportPdf(deckOf(8), BEST)).bytes.byteLength;
    const large = (await exportPdf(deckOf(32), BEST)).bytes.byteLength;
    // Each form refers to one shared resource dictionary; a copy in each form grew quadratically.
    expect(large / small).toBeLessThan(6);
  });

  test('a text shadow is outlines: the words are in the text layer once', async () => {
    const shadow =
      '<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl">' +
      '<a:srgbClr val="000000"><a:alpha val="60000"/></a:srgbClr></a:outerShdw></a:effectLst>';
    const result = await exportPdf(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: shape({
              id: 2,
              x: IN,
              y: IN,
              cx: 6 * IN,
              cy: IN,
              effects: shadow,
              text: para(run('Shadowed words')),
            }),
          },
        ],
      }),
      BEST
    );
    const [page] = await pages(result);
    expect(page.words).toEqual(['Shadowed', 'words']);
    expect(count(page.ops, OPS.beginGroup)).toBeGreaterThan(0);
  });

  test('an unfilled shadowed shape with an empty text body draws as one with no body', async () => {
    const shadow =
      '<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl">' +
      '<a:srgbClr val="000000"><a:alpha val="60000"/></a:srgbClr></a:outerShdw></a:effectLst>';
    const box = { id: 2, x: IN, y: IN, cx: 6 * IN, cy: IN, fill: '<a:noFill/>', effects: shadow };
    const forms = async (text?: string): Promise<number> => {
      const slide = { layout: 'blank' as const, shapes: shape(text ? { ...box, text } : box) };
      const { bytes } = await exportPdf(buildPptx({ slides: [slide] }), BEST);
      return (new TextDecoder('latin1').decode(bytes).match(/\/Subtype\s*\/Form/g) ?? []).length;
    };
    // An empty body has no glyphs, so its shadow has no silhouette to draw.
    expect(await forms('<a:p><a:endParaRPr lang="en-US"/></a:p>')).toBe(await forms());
  });

  test('shadowed text draws its outlines once per shape, inside the operation budget', async () => {
    // A blur this large is 12 ring layers. Outlines repeated in each layer would take about
    // 12 slides × 12 layers × 700 glyphs × 25 operators, past the 2.07 million budget.
    const shadow =
      '<a:effectLst><a:outerShdw blurRad="508000" dist="38100" dir="2700000" algn="tl">' +
      '<a:srgbClr val="000000"><a:alpha val="60000"/></a:srgbClr></a:outerShdw></a:effectLst>';
    const words = 'Shadowed words fill the slide '.repeat(24);
    const slide = {
      layout: 'blank' as const,
      shapes: shape({
        id: 2,
        x: IN,
        y: IN,
        cx: 8 * IN,
        cy: 4 * IN,
        effects: shadow,
        text: para(run(words)),
      }),
    };
    const result = await exportPdf(
      buildPptx({ slides: Array.from({ length: 12 }, () => slide) }),
      BEST
    );
    expect(result.pageCount).toBe(12);
  }, 30_000);
});

describe('diagnostic codes', () => {
  test('each approximation has its code on the fixtures that use it', async () => {
    const of = async (name: string) => codes(await exportPdf(deck(name), BEST));
    expect(await of('fills-lines')).toContain('picture.tile');
    expect(await of('pictures')).toContain('picture.color-effects');
    const effects = await of('effects');
    for (const code of [
      'effect.shadow-blur',
      'effect.glow',
      'effect.soft-edge',
      'effect.inner-shadow-blur',
    ])
      expect(effects).toContain(code);
  });

  test('a shape with a group fill draws the fill of its group', async () => {
    const group =
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="2" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      `<p:grpSpPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${IN}" cy="${IN}"/>` +
      `<a:chOff x="${IN}" y="${IN}"/><a:chExt cx="${IN}" cy="${IN}"/></a:xfrm>` +
      '<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></p:grpSpPr>' +
      shape({ id: 3, x: IN, y: IN, cx: IN, cy: IN, fill: '<a:grpFill/>' }) +
      '</p:grpSp>';
    const result = await exportPdf(
      buildPptx({ slides: [{ layout: 'blank', shapes: group }] }),
      BEST
    );
    expect(await fillColors(result)).toContain('#00ff00');
    expect(codes(result)).toEqual([]);
  });

  test('gradients on lines draw flat, and an empty glow is not reported', async () => {
    const line = (kind: string) =>
      '<a:ln w="25400"><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>' +
      `<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst>${kind}</a:gradFill></a:ln>`;
    const effects = '<a:effectLst><a:glow rad="0"><a:srgbClr val="FF0000"/></a:glow></a:effectLst>';
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
                cx: IN,
                cy: IN,
                line: line('<a:path path="circle"/>'),
                effects,
              }) +
              shape({ id: 3, x: 3 * IN, y: IN, cx: IN, cy: IN, line: line('<a:lin ang="0"/>') }),
          },
        ],
      }),
      BEST
    );
    expect(codes(result)).toContain('line.gradient');
    expect(codes(result)).not.toContain('effect.glow');
  });

  test('a picture fill on text is reported, not drawn', async () => {
    // `a:ln` has no picture fill in the schema (CT_LineProperties), so text is the only case.
    const picture =
      '<a:blipFill><a:blip r:embed="rIdMissing"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>';
    const result = await exportPdf(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: shape({
              id: 2,
              x: IN,
              y: IN,
              cx: 3 * IN,
              cy: IN,
              text: para(run('Picture text', 'sz="4000"', picture)),
            }),
          },
        ],
      }),
      BEST
    );
    expect(codes(result)).toContain('fill.picture');
    expect(result.diagnostics.find((d) => d.code === 'fill.picture')?.severity).toBe('unsupported');
  });

  test('the outline of a picture in a shape of several paths is reported only when it draws', async () => {
    const img = new PNG({ width: 2, height: 2 });
    img.data.fill(150);
    const geom =
      '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>' +
      '<a:path w="2" h="1"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="1" y="0"/></a:lnTo><a:lnTo><a:pt x="1" y="1"/></a:lnTo><a:close/></a:path>' +
      '<a:path w="2" h="1"><a:moveTo><a:pt x="1" y="0"/></a:moveTo><a:lnTo><a:pt x="2" y="0"/></a:lnTo><a:lnTo><a:pt x="2" y="1"/></a:lnTo><a:close/></a:path>' +
      '</a:pathLst></a:custGeom>';
    const of = async (ln: string) =>
      codes(
        await exportPdf(
          buildPptx({
            slides: [
              {
                layout: 'blank',
                rels: [['rIdImg', R_IMAGE, '../media/image1.png']],
                shapes:
                  '<p:pic><p:nvPicPr><p:cNvPr id="2" name="P"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
                  '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
                  `<p:spPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${2 * IN}" cy="${IN}"/></a:xfrm>${geom}${ln}</p:spPr></p:pic>`,
              },
            ],
            media: [
              {
                name: 'image1.png',
                bytes: new Uint8Array(PNG.sync.write(img)),
                contentType: 'image/png',
              },
            ],
          }),
          BEST
        )
      );
    expect(
      await of('<a:ln w="12700"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln>')
    ).toContain('picture.outline');
    expect(await of('<a:ln w="12700"><a:noFill/></a:ln>')).not.toContain('picture.outline');
  });

  test('a gradient background draws as a shading', async () => {
    const background =
      '<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>' +
      '<a:gs pos="100000"><a:srgbClr val="000080"/></a:gs></a:gsLst><a:lin ang="5400000"/>' +
      '</a:gradFill><a:effectLst/></p:bgPr></p:bg>';
    const result = await exportPdf(
      buildPptx({ slides: [{ layout: 'blank', background, shapes: '' }] }),
      BEST
    );
    expect(codes(result).filter((c) => c.startsWith('fill.'))).toEqual([]);
    const [page] = await pages(result);
    expect(count(page.ops, OPS.shadingFill)).toBe(1);
  });

  test('a gradient run whose face has no glyph paints no shading', async () => {
    const gradient =
      '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>' +
      '<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill>';
    const result = await exportPdf(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: shape({
              id: 2,
              x: IN,
              y: IN,
              cx: 6 * IN,
              cy: IN,
              text: para(run('\uE000\uE001', '', gradient)),
            }),
          },
        ],
      }),
      BEST
    );
    const [page] = await pages(result);
    // A clip of no glyphs clips nothing in some readers: the gradient would cover the page.
    expect(count(page.ops, OPS.shadingFill)).toBe(0);
  });
});

describe('ring layers', () => {
  test('composite to the Gaussian edge at the middle of each band', () => {
    const sigma = 50_800;
    const a = 0.6;
    const ringList = rings(2.5 * sigma, sigma / 4);
    const profile = (d: number) => a * (1 - phi(d / sigma));
    const list = layers(ringList, a, profile);
    expect(list.length).toBeLessThanOrEqual(13);
    // Coverage at a distance: 1 - product of (1 - alpha) over the layers that reach it.
    const coverage = (d: number) =>
      1 - list.filter((l) => l.grow >= d).reduce((left, l) => left * (1 - l.alpha), 1);
    for (let j = 1; j < ringList.length; j += 1) {
      const mid = (ringList[j - 1]! + ringList[j]!) / 2;
      expect(coverage(mid)).toBeCloseTo(profile(mid), 2);
    }
    expect(coverage(0)).toBeCloseTo(a, 2);
  });
});

describe('limits and line state', () => {
  test('a gradient with very many stops draws at most the stop limit', async () => {
    // The cascade keeps at most 64 stops of a file (`cascade/paint-props.ts`), so the PDF writer's
    // own bound (`MAX_GRADIENT_STOPS`) and its diagnostic are not reached; the size bound holds.
    const stops = Array.from(
      { length: 2_000 },
      (_, i) =>
        `<a:gs pos="${Math.round((i * 100_000) / 1_999)}"><a:srgbClr val="${i % 2 ? 'FF0000' : '0000FF'}"/></a:gs>`
    ).join('');
    const result = await exportPdf(
      buildPptx({ slides: [{ layout: 'blank', shapes: gradientShape(stops) }] }),
      BEST
    );
    // About 2 objects per drawn stop; 2,000 stops would be far larger.
    expect(result.bytes.byteLength).toBeLessThan(MAX_GRADIENT_STOPS * 400 + 40_000);
  });

  test('a mitered line keeps its miter limit, as the SVG paint does', async () => {
    const line =
      '<a:ln w="190500"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:miter lim="100000"/></a:ln>';
    const result = await exportPdf(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN, prst: 'triangle', line }),
          },
        ],
      }),
      BEST
    );
    const pdf = await getDocument({ data: result.bytes.slice(), useSystemFonts: false }).promise;
    const { fnArray, argsArray } = await (await pdf.getPage(1)).getOperatorList();
    const limit = fnArray.indexOf(OPS.setMiterLimit);
    expect(limit).toBeGreaterThan(-1);
    expect((argsArray[limit] as number[])[0]).toBeCloseTo(1, 3);
  });

  test('darkened and lightened preset paths draw the SVG paint colors', async () => {
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
              prst: 'cube',
              fill: '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>',
            }),
          },
        ],
      }),
      BEST
    );
    const colors = await fillColors(result);
    // `cube`: darkenLess scales toward black by 0.8 (0xcc), lightenLess toward white (0x33).
    expect(colors).toContain('#cc0000');
    expect(colors).toContain('#ff3333');
  });
});

describe('blur profiles', () => {
  test('a thin band is lighter inside and casts a fainter edge than an area', () => {
    const sigma = 50_800;
    const area = blurLayers(0.6, sigma, Infinity, 0);
    const line = blurLayers(0.6, sigma, 12_700, 0);
    const inside = (list: { grow: number; alpha: number }[]) =>
      1 - list.reduce((left, l) => left * (1 - l.alpha), 1);
    // A 1 pt band under a 4 pt σ: a(2Φ(W/2σ) − 1), far below the area's level.
    expect(inside(line)).toBeCloseTo(0.6 * (2 * phi(12_700 / (2 * sigma)) - 1), 2);
    expect(inside(line)).toBeLessThan(inside(area) / 3);
  });

  test('inside the edge, a shadow takes the profile at half its offset', () => {
    const sigma = 50_800;
    const list = blurLayers(0.5, sigma, Infinity, 50_800);
    const inside = 1 - list.reduce((left, l) => left * (1 - l.alpha), 1);
    expect(inside).toBeCloseTo(0.5 * phi(0.5), 2);
  });
});

describe('round 3 cases', () => {
  const shadow =
    '<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000">' +
    '<a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:outerShdw></a:effectLst>';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const pictureWith = (bytes: Uint8Array, name: string, type: string) =>
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes:
            `<p:pic><p:nvPicPr><p:cNvPr id="4" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
            `<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
            `<p:spPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${IN}" cy="${IN}"/></a:xfrm>` +
            `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${shadow}</p:spPr></p:pic>`,
          rels: [['rIdImg', `${R}/image`, `../media/${name}`]],
        },
      ],
      media: [{ name, bytes, contentType: type }],
    });

  test('a translucent shape with a shadow reports the opaque silhouette', async () => {
    const fill =
      '<a:solidFill><a:srgbClr val="4472C4"><a:alpha val="50000"/></a:srgbClr></a:solidFill>';
    const result = await exportPdf(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN, fill, effects: shadow }),
          },
        ],
      }),
      BEST
    );
    expect(codes(result)).toContain('effect.silhouette');
  });

  test('a picture that is not drawn casts no shadow', async () => {
    const gif = new TextEncoder().encode('GIF89a\u0001\u0000\u0001\u0000');
    const result = await exportPdf(pictureWith(gif, 'img.gif', 'image/gif'), BEST);
    expect(codes(result)).toContain('picture.format');
    const [page] = await pages(result);
    expect(count(page.ops, OPS.beginGroup)).toBe(0);
  });

  test('a fill that reaches past the box reports it', async () => {
    const fill =
      `<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000">` +
      `<a:srgbClr val="000080"/></a:gs></a:gsLst><a:path path="rect"><a:fillToRect l="50000" ` +
      `t="50000" r="50000" b="50000"/></a:path></a:gradFill>`;
    const result = await exportPdf(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: shape({
              id: 2,
              x: IN,
              y: IN,
              cx: 2 * IN,
              cy: IN,
              prst: 'wedgeRectCallout',
              fill,
            }),
          },
        ],
      }),
      BEST
    );
    expect(codes(result)).toContain('fill.outside-box');
  });

  test('outline-only text reports its outline', async () => {
    const outline =
      '<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln><a:noFill/>';
    const result = await exportPdf(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: shape({
              id: 2,
              x: IN,
              y: IN,
              cx: 6 * IN,
              cy: IN,
              text: para(run('Outline', '', outline)),
            }),
          },
        ],
      }),
      BEST
    );
    expect(codes(result)).toContain('text.outline');
  });
});

describe('round 4 cases', () => {
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const color = (tag: string, attrs: string) =>
    `<a:${tag} ${attrs}><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:${tag}>`;
  const shadow = `<a:effectLst>${color('outerShdw', 'blurRad="50800" dist="38100" dir="2700000"')}</a:effectLst>`;
  const inner = `<a:effectLst>${color('innerShdw', 'blurRad="50800" dist="38100" dir="2700000"')}</a:effectLst>`;
  const codesOf = async (shapes: string) =>
    codes(await exportPdf(buildPptx({ slides: [{ layout: 'blank', shapes }] }), BEST));

  test('a shape whose picture fill is not drawn casts no shadow', async () => {
    const gif = new TextEncoder().encode('GIF89a\u0001\u0000\u0001\u0000');
    const fill =
      '<a:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>';
    const result = await exportPdf(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN, fill, effects: shadow }),
            rels: [['rIdImg', `${R}/image`, '../media/img.gif']],
          },
        ],
        media: [{ name: 'img.gif', bytes: gif, contentType: 'image/gif' }],
      }),
      BEST
    );
    expect(codes(result)).toContain('picture.format');
    const [page] = await pages(result);
    expect(count(page.ops, OPS.beginGroup)).toBe(0);
  });

  test('an inner shadow that the writer starts at the fill is reported', async () => {
    const wide = '<a:ln w="76200"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>';
    const thin = '<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>';
    const square = (fill: string, line: string, prst = 'rect') =>
      shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN, prst, fill, line, effects: inner });
    // Unfilled with a 6 pt outline: the editor shades the line; the writer draws nothing.
    const unfilled = await codesOf(square('<a:noFill/>', wide));
    expect(unfilled).toContain('effect.inner-shadow');
    expect(unfilled).not.toContain('effect.inner-shadow-blur');
    // Filled with a thin outline: one filled path, drawn as the editor draws it.
    const plain = await codesOf(square(solid('4472C4'), thin));
    expect(plain).not.toContain('effect.inner-shadow');
    expect(plain).toContain('effect.inner-shadow-blur');
    // `can` has two filled paths: the even-odd inverse makes the top a hole.
    expect(await codesOf(square(solid('4472C4'), thin, 'can'))).toContain('effect.inner-shadow');
  });

  test('a shadow with nothing to cast is not reported', async () => {
    const empty = shape({
      id: 2,
      x: IN,
      y: IN,
      cx: IN,
      cy: IN,
      fill: '<a:noFill/>',
      line: '<a:ln><a:noFill/></a:ln>',
      effects: shadow,
    });
    expect(await codesOf(empty)).not.toContain('effect.shadow-blur');
  });

  test('a curve past the box is reported, and rounding in the preset formulas is not', async () => {
    const fill =
      `<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000">` +
      `<a:srgbClr val="000080"/></a:gs></a:gsLst><a:path path="rect"><a:fillToRect l="50000" ` +
      `t="50000" r="50000" b="50000"/></a:path></a:gradFill>`;
    const of = (prst: string) =>
      codesOf(shape({ id: 2, x: IN, y: IN, cx: 4 * IN, cy: 4 * IN, prst, fill }));
    // Every end point of `heart` is in the box; its curves bulge past it.
    expect(await of('heart')).toContain('fill.outside-box');
    // `heptagon` reaches a few EMU past the box through rounding.
    expect(await of('heptagon')).not.toContain('fill.outside-box');
  });

  test('a picture with curved geometry is not reported as outside its box', async () => {
    // The editor also draws a `p:pic` as its image clipped to the geometry, so nothing differs.
    const png = new PNG({ width: 2, height: 2 });
    png.data.fill(200);
    const picture =
      `<p:pic><p:nvPicPr><p:cNvPr id="4" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${4 * IN}" cy="${4 * IN}"/></a:xfrm>` +
      `<a:prstGeom prst="heart"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
    const result = await exportPdf(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: picture,
            rels: [['rIdImg', `${R}/image`, '../media/img.png']],
          },
        ],
        media: [{ name: 'img.png', bytes: PNG.sync.write(png), contentType: 'image/png' }],
      }),
      BEST
    );
    expect(codes(result)).toEqual([]);
  });
});
