/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The PDF writer mirrors the editor's gradient geometry and picture placement
// (`packages/core/src/output/paint-fill.ts`). These tests paint the same slide with the editor's
// SVG paint, in a happy-dom window of their own, and with the PDF writer, and fail when the two
// disagree: a change to the editor's rule must reach the PDF writer too.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, test } from 'bun:test';
import { Window } from 'happy-dom';
import { PNG } from 'pngjs';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { decodePDFRawStream, PDFDocument, PDFName, PDFRawStream } from 'pdf-lib';
import { openPresentation } from '@pptx-editor.dev/core/editor';
import { paintSlide } from '@pptx-editor.dev/core/output';
import { installDrawing } from '@pptx-editor.dev/pro/drawing';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { exportPdf } from '../index.ts';

const IN = 914_400;
const BEST = { fidelityPolicy: 'best-effort' } as const;
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** The editor's SVG of slide 1, painted into a document of this test's own window. */
async function svgOf(bytes: Uint8Array): Promise<Element> {
  const opened = await openPresentation(bytes);
  if (!opened.ok) throw new Error('deck');
  const layout = opened.session.layout(0)!;
  const { document } = new Window();
  const root = paintSlide(document as unknown as Document, layout, {
    images: () => 'blob:picture',
    idPrefix: 't',
    hitTargets: false,
  });
  opened.session.dispose();
  return root as unknown as Element;
}

/** The numbers of each `/Coords` array of the PDF's shadings, in object order. */
function coordsOf(pdf: Uint8Array): { type: number; coords: number[] }[] {
  const text = new TextDecoder('latin1').decode(pdf);
  return [
    ...text.matchAll(/\/ShadingType (\d)\s*\/ColorSpace \/DeviceRGB\s*\/Coords \[ ([^\]]+) \]/g),
  ].map((m) => ({ type: Number(m[1]), coords: m[2]!.trim().split(/\s+/).map(Number) }));
}

const attrs = (el: Element | null, names: string[]): number[] => {
  if (!el) throw new Error('missing SVG element');
  return names.map((n) => Number(el.getAttribute(n)));
};

const gradientDeck = (fill: string, cx = 3 * IN, cy = IN) =>
  buildPptx({
    slides: [{ layout: 'blank', shapes: shape({ id: 2, x: IN, y: IN, cx, cy, fill }) }],
  });

const stops =
  '<a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>' +
  '<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst>';

/** The gray `/Coords` of the PDF's shadings: the fades of reflection soft masks. */
function grayCoordsOf(pdf: Uint8Array): number[][] {
  const text = new TextDecoder('latin1').decode(pdf);
  return [
    ...text.matchAll(/\/ShadingType 2\s*\/ColorSpace \/DeviceGray\s*\/Coords \[ ([^\]]+) \]/g),
  ].map((m) => m[1]!.trim().split(/\s+/).map(Number));
}

/** The operands of every `cm` operator in the PDF's form XObjects. */
async function formMatrices(pdf: Uint8Array): Promise<number[][]> {
  const doc = await PDFDocument.load(pdf);
  const out: number[][] = [];
  for (const [, obj] of doc.context.enumerateIndirectObjects()) {
    if (!(obj instanceof PDFRawStream)) continue;
    if (obj.dict.get(PDFName.of('Subtype')) !== PDFName.of('Form')) continue;
    const raw = obj.dict.has(PDFName.of('Filter'))
      ? decodePDFRawStream(obj).decode()
      : obj.contents;
    const tokens = new TextDecoder('latin1').decode(raw).split(/\s+/);
    tokens.forEach((t, i) => {
      if (t === 'cm') out.push(tokens.slice(i - 6, i).map(Number));
    });
  }
  return out;
}

describe('the PDF writer and the SVG paint agree', () => {
  test('on the linear gradient axis, scaled and not', async () => {
    for (const lin of ['<a:lin ang="0" scaled="0"/>', '<a:lin ang="2700000" scaled="1"/>']) {
      const bytes = gradientDeck(`<a:gradFill>${stops}${lin}</a:gradFill>`, 2 * IN, IN);
      const svg = await svgOf(bytes);
      const want = attrs(svg.querySelector('linearGradient'), ['x1', 'y1', 'x2', 'y2']);
      const [shading] = coordsOf((await exportPdf(bytes, BEST)).bytes);
      expect(shading!.type).toBe(2);
      shading!.coords.forEach((v, i) => expect(v).toBeCloseTo(want[i]!, 0));
    }
  });

  test('on the radial gradient focal point, center, and radius', async () => {
    const path =
      '<a:path path="circle"><a:fillToRect l="25000" t="50000" r="75000" b="50000"/></a:path>';
    const bytes = gradientDeck(`<a:gradFill>${stops}${path}</a:gradFill>`);
    const svg = await svgOf(bytes);
    const [cx, cy, r, fx, fy] = attrs(svg.querySelector('radialGradient'), [
      'cx',
      'cy',
      'r',
      'fx',
      'fy',
    ]);
    const [shading] = coordsOf((await exportPdf(bytes, BEST)).bytes);
    expect(shading!.type).toBe(3);
    // The first stop is at the focal point, the last on the circle (the reference PDFs' rule).
    const [x0, y0, r0, x1, y1, r1] = shading!.coords;
    for (const [got, wanted] of [
      [x0, fx],
      [y0, fy],
      [x1, cx],
      [y1, cy],
      [r1, r],
    ] as const)
      expect(got).toBeCloseTo(wanted!, 0);
    expect(r0).toBe(0);
  });

  test('on the trapezoids of a rectangular path gradient', async () => {
    const path =
      '<a:path path="rect"><a:fillToRect l="25000" t="40000" r="50000" b="20000"/></a:path>';
    const bytes = gradientDeck(`<a:gradFill>${stops}${path}</a:gradFill>`);
    const svg = await svgOf(bytes);
    // Each trapezoid's gradient line, to 10 EMU, in sorted order: both painters draw four.
    const line = (values: number[]) => values.map((v) => Math.round(v / 10)).join();
    const want = [...svg.querySelectorAll('linearGradient')]
      .map((g) => line(attrs(g, ['x1', 'y1', 'x2', 'y2'])))
      .sort();
    const got = coordsOf((await exportPdf(bytes, BEST)).bytes)
      .filter((s) => s.type === 2)
      .map((s) => line(s.coords))
      .sort();
    expect(want).toHaveLength(4);
    expect(got).toEqual(want);
  });

  test('on the place and fade of a reflection, turned with the shape and not', async () => {
    // The SVG renderer is the Pro one; opening an export session installs the same one.
    installDrawing();
    for (const [rot, rotWithShape] of [
      [0, 0],
      [1_800_000, 1],
      [1_800_000, 0],
    ] as const) {
      const effects =
        '<a:effectLst><a:reflection blurRad="6350" stA="52000" stPos="10000" endA="300" endPos="35000" ' +
        `dist="12700" dir="5400000" sy="-100000" algn="bl" rotWithShape="${rotWithShape}"/></a:effectLst>`;
      const bytes = buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: shape({
              id: 2,
              x: IN,
              y: IN,
              cx: 2 * IN,
              cy: IN,
              rot,
              fill: '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>',
              effects,
            }),
          },
        ],
      });
      const svg = await svgOf(bytes);
      const gradient = svg.querySelector('linearGradient');
      const [x1, y1, x2, y2] = attrs(gradient, ['x1', 'y1', 'x2', 'y2']);
      const offsets = [...gradient!.querySelectorAll('stop')].map((s) =>
        Number(s.getAttribute('offset'))
      );
      const at = (t: number): number[] => [x1! + (x2! - x1!) * t, y1! + (y2! - y1!) * t];
      const pdf = (await exportPdf(bytes, BEST)).bytes;
      const fades = grayCoordsOf(pdf);
      expect(fades).toHaveLength(1);
      // The PDF fade runs between the two SVG stops, along the same line.
      const want = [...at(offsets[0]!), ...at(offsets[1]!)];
      fades[0]!.forEach((v, i) => expect(Math.abs(v - want[i]!)).toBeLessThan(2));
      // The reflected copy sits under the same transform in both.
      const use = svg.querySelector('[data-effect="reflection"] use')!.getAttribute('transform')!;
      const m = /matrix\(([^)]+)\)/.exec(use)![1]!.split(/\s+/).map(Number);
      const near = (c: number[]): boolean =>
        c.every((v, i) => Math.abs(v - m[i]!) <= (i < 4 ? 1e-2 : 2));
      expect((await formMatrices(pdf)).some(near)).toBe(true);
    }
  });

  for (const [name, blip] of [
    [
      'a cropped picture',
      '<a:srcRect l="10000" t="20000" r="5000"/><a:stretch><a:fillRect/></a:stretch>',
    ],
    [
      'a cropped picture in a fill rectangle',
      '<a:srcRect l="10000" t="20000" r="5000"/>' +
        '<a:stretch><a:fillRect l="10000" t="5000" r="20000" b="15000"/></a:stretch>',
    ],
  ] as const)
    test(`on the place of ${name}`, async () => {
      await pictureAgrees(blip);
    });

  test('on the bars of underline, double underline, and strikethrough', async () => {
    const fill = (rgb: string): string => `<a:solidFill><a:srgbClr val="${rgb}"/></a:solidFill>`;
    const uFill = (rgb: string): string => `<a:uFill>${fill(rgb)}</a:uFill>`;
    const text = para(
      // An underline in its own fill (a:uFill), drawn over the glyphs.
      run('Under', 'sz="2400" u="sng"', fill('FF0000') + uFill('FF00FF')) +
        // A run with no fill draws no bars, even with an underline fill.
        run('Ghost', 'sz="2400" u="sng"', '<a:noFill/>' + uFill('00FFFF')) +
        run('Double', 'sz="2400" u="dbl"', fill('00FF00')) +
        // The trailing space is not struck, in either paint.
        run(' Struck ', 'sz="2400" strike="sngStrike"', fill('0000FF'))
    );
    const bytes = buildPptx({
      slides: [
        { layout: 'blank', shapes: shape({ id: 2, x: IN, y: IN, cx: 6 * IN, cy: IN, text }) },
      ],
    });
    // The SVG bars are in points inside the frame; the PDF bars are in EMU inside the frame.
    const svgBars = [...(await svgOf(bytes)).querySelectorAll('rect')]
      .filter((r) => r.getAttribute('data-kind') !== 'background')
      .map((r) => ({
        fill: r.getAttribute('fill')!,
        box: attrs(r, ['x', 'y', 'width', 'height']),
      }));
    const pdf = await getDocument({
      data: (await exportPdf(bytes, BEST)).bytes.slice(),
      useSystemFonts: false,
    }).promise;
    const { fnArray, argsArray } = await (await pdf.getPage(1)).getOperatorList();
    const pdfBars: { fill: string; box: number[] }[] = [];
    let color = '';
    let firstGlyph = -1;
    let firstBar = -1;
    fnArray.forEach((fn, i) => {
      const args = argsArray[i] as unknown[];
      if (fn === OPS.setFillRGBColor) color = String(args[0]);
      if (fn === OPS.showText && firstGlyph < 0) firstGlyph = i;
      if (fn !== OPS.constructPath || color === '#ffffff') return;
      if (firstBar < 0) firstBar = i;
      // pdf.js path data: 0 moveTo x y, 1 lineTo x y, 4 closePath. One box per subpath.
      const data = Array.from((args[1] as ArrayLike<number>[])[0]!);
      let xs: number[] = [];
      let ys: number[] = [];
      for (let k = 0; k < data.length;) {
        const op = data[k]!;
        if (op === 4) {
          const [x, y] = [Math.min(...xs), Math.min(...ys)];
          pdfBars.push({ fill: color, box: [x, y, Math.max(...xs) - x, Math.max(...ys) - y] });
          [xs, ys] = [[], []];
          k += 1;
        } else {
          xs.push(data[k + 1]!);
          ys.push(data[k + 2]!);
          k += 3;
        }
      }
    });
    expect(pdfBars.map((b) => b.fill)).toEqual(svgBars.map((b) => b.fill));
    expect(pdfBars.map((b) => b.fill)).toEqual(['#ff00ff', '#00ff00', '#00ff00', '#0000ff']);
    expect(firstBar).toBeGreaterThan(firstGlyph);
    for (const [i, bar] of svgBars.entries())
      for (const [k, v] of bar.box.entries())
        expect(pdfBars[i]!.box[k]! / 12_700).toBeCloseTo(v, 1);
  });

  test('on the shapes and text of a diagram saved drawing', async () => {
    const bytes = new Uint8Array(
      readFileSync(join(import.meta.dirname, '../../../../fixtures/diagrams.pptx'))
    );
    const svg = await svgOf(bytes);
    // The frame draws its drawing, not the dashed placeholder.
    expect(svg.querySelector('[stroke-dasharray]')).toBeNull();
    const svgFills = [...svg.querySelectorAll('[data-kind="frame"] path')]
      .map((p) => p.getAttribute('fill')!)
      .filter((f) => f !== 'none');
    const svgText = [...svg.querySelectorAll('[data-kind="frame"] text')]
      .map((t) => t.textContent!.trim())
      .filter(Boolean);
    const result = await exportPdf(bytes, BEST);
    expect(result.diagnostics.some((d) => d.code === 'frame.diagram')).toBe(false);
    const pdf = await getDocument({ data: result.bytes.slice(), useSystemFonts: false }).promise;
    const page = await pdf.getPage(1);
    const { fnArray, argsArray } = await page.getOperatorList();
    const pdfFills: string[] = [];
    let color = '';
    fnArray.forEach((fn, i) => {
      if (fn === OPS.setFillRGBColor) color = String((argsArray[i] as unknown[])[0]);
      if (fn === OPS.constructPath && color) pdfFills.push(color);
    });
    // Nine shapes, each with its fill, in the SVG paint's order. The PDF also strokes and
    // writes text, so its fills hold the SVG's fills in order.
    expect(svgFills).toHaveLength(9);
    let at = 0;
    for (const fill of svgFills) {
      at = pdfFills.indexOf(fill.toLowerCase(), at);
      expect(at).toBeGreaterThanOrEqual(0);
      at += 1;
    }
    const pdfText = (await page.getTextContent()).items
      .map((item) => ('str' in item ? item.str.trim() : ''))
      .filter(Boolean);
    expect(svgText).toEqual([
      'Plan',
      'Scope',
      'Brief',
      'Write',
      'Draft',
      'Art',
      'Ship',
      'Check',
      'Send',
    ]);
    expect(pdfText.slice(1)).toEqual(svgText);
  });
});

async function pictureAgrees(blip: string): Promise<void> {
  {
    const png = new PNG({ width: 2, height: 2 });
    png.data.fill(200);
    const picture =
      `<p:pic><p:nvPicPr><p:cNvPr id="4" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="rIdImg"/>${blip}</p:blipFill>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${IN}" cy="${IN}"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
    const bytes = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: picture,
          rels: [['rIdImg', `${R}/image`, '../media/img.png']],
        },
      ],
      media: [{ name: 'img.png', bytes: PNG.sync.write(png), contentType: 'image/png' }],
    });
    const svg = await svgOf(bytes);
    const [x, y, width, height] = attrs(svg.querySelector('image'), ['x', 'y', 'width', 'height']);
    const pdf = await getDocument({
      data: (await exportPdf(bytes, BEST)).bytes.slice(),
      useSystemFonts: false,
    }).promise;
    const { fnArray, argsArray } = await (await pdf.getPage(1)).getOperatorList();
    const at = fnArray.indexOf(OPS.paintImageXObject);
    // The image space is the unit square: `cm iw 0 0 -ih x y+ih`.
    const [iw, , , negIh, e, f] = argsArray[fnArray.lastIndexOf(OPS.transform, at)] as number[];
    expect(iw!).toBeCloseTo(width!, 0);
    expect(-negIh!).toBeCloseTo(height!, 0);
    expect(e!).toBeCloseTo(x!, 0);
    expect(f! + negIh!).toBeCloseTo(y!, 0);
  }
}
