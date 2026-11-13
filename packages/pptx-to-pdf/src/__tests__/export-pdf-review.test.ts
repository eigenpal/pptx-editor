/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { fontConfiguration } from '../../../../scripts/fixtures/font-configuration.ts';
// Limits, cancellation, metadata, fonts, pictures, and diagnostics placement.

import { missingFontsDeck } from '../../../core/src/editor/__tests__/font-test-helpers.ts';
import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { unzipSync, zipSync } from 'fflate';
import { PDFDocument } from 'pdf-lib';
import { PNG } from 'pngjs';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { openPresentation } from '@pptx-editor.dev/core/editor';
import type { FontSource } from '@pptx-editor.dev/core/layout';
import { readZip, writeZip } from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { buildEot } from '../../../core/src/store/fonts/__tests__/fixtures.ts';
import {
  exportPdf,
  PdfEncodingError,
  PdfFidelityError,
  PdfOutputLimitError,
  PdfWorkLimitError,
  type PdfExportResult,
} from '../index.ts';
import { withPdfWork } from '../options.ts';

const FIXTURES = join(import.meta.dir, '../../../../fixtures');
const deck = (name: string): Uint8Array => readFileSync(join(FIXTURES, `${name}.pptx`));
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const IN = 914_400;
const BEST = { fidelityPolicy: 'best-effort' } as const;

async function pdfText(result: PdfExportResult) {
  const pdf = await getDocument({ data: result.bytes.slice(), useSystemFonts: false }).promise;
  const pages = [];
  for (let i = 1; i <= pdf.numPages; i += 1) {
    const content = await (await pdf.getPage(i)).getTextContent();
    pages.push(content.items.filter((item) => 'str' in item && item.str.trim() !== ''));
  }
  return pages as { str: string; transform: number[] }[][];
}

/** A substitute face from the engine, as raw TrueType bytes. */
async function substituteFace(): Promise<FontSource> {
  const opened = await openPresentation(deck('text-runs'));
  if (!opened.ok) throw new Error('fixture');
  const face = opened.session.faces.find((f) => f.origin === 'substitute' && !f.bold && !f.italic);
  if (!face) throw new Error('no substitute face');
  return { family: face.family, bold: false, italic: false, bytes: face.bytes, origin: 'test' };
}

/** A copy of a TrueType face with its OS/2 fsType replaced. */
function withFsType(bytes: Uint8Array, fsType: number): Uint8Array {
  const out = bytes.slice();
  const view = new DataView(out.buffer);
  const tables = view.getUint16(4);
  for (let i = 0; i < tables; i += 1) {
    const rec = 12 + i * 16;
    if (view.getUint32(rec) !== 0x4f532f32) continue;
    view.setUint16(view.getUint32(rec + 8) + 8, fsType);
    return out;
  }
  throw new Error('no OS/2 table');
}

const textSlide = (text: string, runChildren = '') => ({
  layout: 'blank' as const,
  shapes: shape({
    id: 2,
    x: IN,
    y: IN,
    cx: 6 * IN,
    cy: IN,
    text: para(run(text, '', runChildren)),
  }),
});

function tinyPng(): Uint8Array {
  const png = new PNG({ width: 2, height: 2 });
  png.data.fill(200);
  return PNG.sync.write(png);
}

/** A PNG header that declares 30000 × 30000 pixels and holds no image data. */
function hugePngHeader(): Uint8Array {
  const out = new Uint8Array(8 + 25 + 12);
  out.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const v = new DataView(out.buffer);
  v.setUint32(8, 13);
  out.set([0x49, 0x48, 0x44, 0x52], 12); // IHDR
  v.setUint32(16, 30_000);
  v.setUint32(20, 30_000);
  out.set([8, 6, 0, 0, 0], 24);
  v.setUint32(33, 0);
  out.set([0x49, 0x45, 0x4e, 0x44], 37); // IEND
  return out;
}

const picture = (srcRect = '', blipChildren = '', ln = '') =>
  `<p:pic><p:nvPicPr><p:cNvPr id="4" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
  `<p:blipFill><a:blip r:embed="rIdImg">${blipChildren}</a:blip>${srcRect}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${IN}" cy="${IN}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${ln}</p:spPr></p:pic>`;

const pictureDeck = (bytes: Uint8Array, srcRect = '', blipChildren = '', ln = '') =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: picture(srcRect, blipChildren, ln),
        rels: [['rIdImg', `${R}/image`, '../media/img.png']],
      },
    ],
    media: [{ name: 'img.png', bytes, contentType: 'image/png' }],
  });

/** A pattern line, which this writer does not draw yet. */
const patternLine =
  '<a:ln w="12700"><a:pattFill prst="pct50"><a:fgClr><a:srgbClr val="FF0000"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill></a:ln>';

describe('cancellation and limits', () => {
  test('an engine failure inside the work is PdfEncodingError, never a bare TypeError', async () => {
    const run = withPdfWork({}, async () => {
      throw new TypeError('engine bug');
    });
    await expect(run).rejects.toBeInstanceOf(PdfEncodingError);
    await expect(run).rejects.toMatchObject({ code: 'encodingFailed' });
  });

  test('an aborted signal stops before the deck opens, with its reason as the cause', async () => {
    const controller = new AbortController();
    const reason = new Error('user cancelled');
    controller.abort(reason);
    const error = await exportPdf(deck('text-runs'), { ...BEST, signal: controller.signal }).catch(
      (e: unknown) => e
    );
    expect(error).toBeInstanceOf(PdfWorkLimitError);
    expect((error as PdfWorkLimitError).code).toBe('aborted');
    expect((error as PdfWorkLimitError).cause).toBe(reason);
  });

  test('a signal aborted during the conversion stops it', async () => {
    const controller = new AbortController();
    const pending = exportPdf(deck('preset-shapes'), { ...BEST, signal: controller.signal });
    controller.abort();
    await expect(pending).rejects.toMatchObject({ code: 'aborted' });
  });

  test('timeoutMs stops a conversion that runs past it', async () => {
    const run = exportPdf(deck('preset-shapes'), { ...BEST, timeoutMs: 1 });
    await expect(run).rejects.toMatchObject({ name: 'PdfWorkLimitError', code: 'timeout' });
  });

  test('maxOutputBytes refuses a PDF that is too large', async () => {
    const run = exportPdf(deck('text-runs'), { ...BEST, maxOutputBytes: 1000 });
    await expect(run).rejects.toBeInstanceOf(PdfOutputLimitError);
  });

  test('fonts accepts one source as well as an array', async () => {
    const face = await substituteFace();
    const one = await exportPdf(deck('text-runs'), { ...BEST, fonts: fontConfiguration([face]) });
    const many = await exportPdf(deck('text-runs'), {
      ...BEST,
      fonts: [fontConfiguration([face])],
    });
    expect(one.bytes.byteLength).toBe(many.bytes.byteLength);
  });
});

describe('metadata', () => {
  test('title and author come from the deck core properties; dates are fixed', async () => {
    const result = await exportPdf(deck('text-runs'), BEST);
    const pdf = await PDFDocument.load(result.bytes, { updateMetadata: false });
    expect(pdf.getTitle()).toBe('Fixture');
    expect(pdf.getAuthor()).toBe('PPTX Editor fixtures');
    expect(pdf.getProducer()).toBe('pptx-editor.dev');
    expect(pdf.getCreationDate()?.toISOString()).toBe('2020-01-01T00:00:00.000Z');
  });
});

describe('fonts', () => {
  test('a deck with an embedded font draws its text with that font', async () => {
    const face = await substituteFace();
    const fontData = buildEot(face.bytes, { family: 'Embedded Test', flags: 0 });
    const bytes = buildPptx({
      slides: [textSlide('Embedded words here', '<a:latin typeface="Embedded Test"/>')],
      presentationLists:
        '<p:embeddedFontLst><p:embeddedFont><p:font typeface="Embedded Test"/><p:regular r:id="rIdFont"/></p:embeddedFont></p:embeddedFontLst>',
      presentationRels: [['rIdFont', `${R}/font`, 'fonts/font1.fntdata']],
      binaryParts: [
        {
          name: 'ppt/fonts/font1.fntdata',
          bytes: fontData,
          contentType: 'application/x-fontdata',
        },
      ],
    });
    const result = await exportPdf(bytes, BEST);
    expect(result.fontResolution.faces).toContainEqual({
      family: 'Embedded Test',
      origin: 'embedded',
    });
    expect(result.fontResolution.substitutions.map((s) => s.family)).not.toContain('Embedded Test');
    const [page] = await pdfText(result);
    expect(page.map((i) => i.str).join('')).toContain('Embedded');
  });

  test('a font that forbids embedding is not embedded, and strict refuses the deck', async () => {
    const face = await substituteFace();
    const restricted = { ...face, family: 'Locked Font', bytes: withFsType(face.bytes, 0x0002) };
    const bytes = buildPptx({
      slides: [textSlide('Locked words', '<a:latin typeface="Locked Font"/>')],
    });
    const result = await exportPdf(bytes, { ...BEST, fonts: fontConfiguration([restricted]) });
    expect(result.diagnostics.map((d) => d.code)).toContain('font.not-embedded');
    const [page] = await pdfText(result);
    expect(page.map((i) => i.str).join('')).not.toContain('Locked');
    await expect(
      exportPdf(bytes, { fonts: fontConfiguration([restricted]) })
    ).rejects.toBeInstanceOf(PdfFidelityError);
  });

  test('rotated text is drawn rotated; flipped text is not mirrored', async () => {
    const bytes = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes:
            shape({
              id: 2,
              x: IN,
              y: IN,
              cx: 3 * IN,
              cy: IN,
              rot: 5_400_000,
              text: para(run('Rotated')),
            }) +
            shape({
              id: 3,
              x: IN,
              y: 4 * IN,
              cx: 3 * IN,
              cy: IN,
              flipH: true,
              text: para(run('Flipped')),
            }),
        },
      ],
    });
    const [page] = await pdfText(await exportPdf(bytes, BEST));
    const rotated = page.find((i) => i.str.startsWith('R'));
    const flipped = page.find((i) => i.str.startsWith('F'));
    if (!rotated || !flipped) throw new Error('text not found');
    // pdf.js text transform [a, b, c, d]: a 90° turn puts the scale in b and c.
    expect(Math.abs(rotated.transform[0])).toBeLessThan(0.01);
    expect(Math.abs(rotated.transform[1])).toBeGreaterThan(1);
    expect(flipped.transform[0]).toBeGreaterThan(0);
    expect(Math.abs(flipped.transform[1])).toBeLessThan(0.01);
  });
});

describe('pictures', () => {
  test('an OLE object draws its picture (ECMA-376 Part 1, 19.3.1.36), and a chart draws', async () => {
    const result = await exportPdf(deck('slide-structure'), BEST);
    const codes = result.diagnostics.map((d) => d.code);
    expect(codes).not.toContain('frame.chart');
    expect(codes).not.toContain('frame.ole');
    const pdf = await getDocument({ data: result.bytes.slice(), useSystemFonts: false }).promise;
    const { fnArray } = await (await pdf.getPage(4)).getOperatorList();
    expect(fnArray).toContain(OPS.paintImageXObject);
  });

  test('a chart type outside the drawing module keeps the placeholder note', async () => {
    const zip = readZip(deck('charts'));
    if (!zip.ok) throw new Error('charts.pptx');
    const entries = new Map(zip.entries);
    const name = 'ppt/charts/chart1.xml';
    const xml = new TextDecoder().decode(entries.get(name) ?? entries.get(`/${name}`));
    const threeD = xml.replaceAll('c:barChart>', 'c:bar3DChart>');
    expect(threeD).not.toBe(xml);
    entries.set(entries.has(name) ? name : `/${name}`, new TextEncoder().encode(threeD));
    const result = await exportPdf(writeZip(entries), BEST);
    const charts = result.diagnostics.filter((d) => d.code === 'frame.chart');
    expect(charts.map((d) => d.slideNumber)).toEqual([1]);
  });

  test('a cropped picture is drawn cropped, inside its box', async () => {
    const result = await exportPdf(pictureDeck(tinyPng(), '<a:srcRect l="10000"/>'), BEST);
    expect(result.diagnostics.filter((d) => d.severity !== 'information')).toEqual([]);
    const pdf = await getDocument({ data: result.bytes.slice(), useSystemFonts: false }).promise;
    const { fnArray, argsArray } = await (await pdf.getPage(1)).getOperatorList();
    const image = fnArray.indexOf(OPS.paintImageXObject);
    expect(image).toBeGreaterThan(0);
    // The picture is clipped to its geometry, then drawn 1/0.9 as wide, 10 % of that to the left.
    expect(fnArray.slice(0, image)).toContain(OPS.clip);
    const placed = argsArray[fnArray.lastIndexOf(OPS.transform, image)] as number[];
    expect(placed[0]).toBeCloseTo(IN / 0.9, 0);
    expect(placed[4]).toBeCloseTo(-0.1 * (IN / 0.9), 0);
  });

  test('a picture transparency from the file is clamped to the 0–1 that PDF allows', async () => {
    const negative = await exportPdf(
      pictureDeck(tinyPng(), '', '<a:alphaModFix amt="-50000"/>'),
      BEST
    );
    const pdf = new TextDecoder('latin1').decode(negative.bytes);
    expect(pdf).not.toMatch(/\/ca -/);
    expect(pdf).toMatch(/\/ca 0\b/);
  });

  test('an SVG picture is drawn from its PNG fallback and says so', async () => {
    const svg =
      '<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
      '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rIdImg"/>' +
      '</a:ext></a:extLst>';
    const result = await exportPdf(pictureDeck(tinyPng(), '', svg), BEST);
    const fallback = result.diagnostics.find((d) => d.code === 'picture.svg-fallback');
    // The deck's own raster fallback of the SVG: reported, but strict does not refuse it.
    expect(fallback?.severity).toBe('information');
    expect(new TextDecoder('latin1').decode(result.bytes)).toContain('/Subtype /Image');
  });

  test('a picture border is drawn outside the picture', async () => {
    const ln = '<a:ln w="38100"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln>';
    const result = await exportPdf(pictureDeck(tinyPng(), '', '', ln), BEST);
    const pdf = await getDocument({ data: result.bytes.slice(), useSystemFonts: false }).promise;
    const { fnArray, argsArray } = await (await pdf.getPage(1)).getOperatorList();
    // An even-odd clip to the outside of the frame, then a stroke twice the line width.
    const clip = fnArray.indexOf(OPS.eoClip);
    expect(clip).toBeGreaterThan(-1);
    const widths = fnArray.flatMap((op, i) =>
      op === OPS.setLineWidth && i > clip ? [argsArray[i][0] as number] : []
    );
    expect(widths.length).toBe(1);
    // 3 pt in EMU, doubled; the page matrix maps EMU to points.
    expect(widths[0]).toBeCloseTo(76_200, 0);
  });

  test('a picture cropped to a shape is clipped to it; its effects are drawn', async () => {
    const deck = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: picture('', '<a:grayscl/>').replace('prst="rect"', 'prst="ellipse"'),
          rels: [['rIdImg', `${R}/image`, '../media/img.png']],
        },
      ],
      media: [{ name: 'img.png', bytes: tinyPng(), contentType: 'image/png' }],
    });
    const result = await exportPdf(deck, BEST);
    expect(result.diagnostics.map((d) => d.code)).not.toContain('picture.color-effects');
    const pdf = await getDocument({ data: result.bytes.slice(), useSystemFonts: false }).promise;
    const { fnArray } = await (await pdf.getPage(1)).getOperatorList();
    // A nonzero clip to the ellipse comes before the image.
    const image = fnArray.indexOf(OPS.paintImageXObject);
    const clip = fnArray.indexOf(OPS.clip);
    expect(clip).toBeGreaterThan(-1);
    expect(clip).toBeLessThan(image);
  });

  test('a PNG that declares a huge size is refused from its header', async () => {
    const before = process.memoryUsage().rss;
    const result = await exportPdf(pictureDeck(hugePngHeader()), BEST);
    expect(result.diagnostics.map((d) => d.code)).toContain('picture.too-large');
    expect(new TextDecoder('latin1').decode(result.bytes)).not.toContain('/Subtype /Image');
    // Decoding 30000 × 30000 RGBA would take 3.6 GB.
    expect(process.memoryUsage().rss - before).toBeLessThan(500 * 1024 * 1024);
  });
});

describe('diagnostics placement', () => {
  const deckWithHidden = () =>
    buildPptx({
      slides: [
        textSlide('First'),
        { ...textSlide('Hidden'), hidden: true },
        {
          layout: 'blank',
          shapes: shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN, line: patternLine }),
        },
      ],
    });

  test('slide numbers stay the deck numbers; page numbers follow the PDF after a hidden slide', async () => {
    const result = await exportPdf(deckWithHidden(), BEST);
    expect(result.pageCount).toBe(2);
    const fill = result.diagnostics.find((d) => d.code === 'fill.pattern');
    expect(fill).toMatchObject({ slideIndex: 2, slideNumber: 3, pageIndex: 1, pageNumber: 2 });
  });

  test('single and double underline and strikethrough pass strict; other underline styles are approximate', async () => {
    const deckOf = (runs: string): Uint8Array =>
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: shape({ id: 2, x: IN, y: IN, cx: 6 * IN, cy: IN, text: para(runs) }),
          },
        ],
      });
    const exact = await exportPdf(
      deckOf(
        run('One', 'u="sng"') +
          run('Two', 'u="dbl"') +
          run('Gone', 'strike="sngStrike"') +
          run('Twice', 'strike="dblStrike"')
      )
    );
    expect(exact.diagnostics.filter((d) => d.severity !== 'information')).toEqual([]);
    const wavy = await exportPdf(deckOf(run('Wavy', 'u="wavy"')), BEST);
    expect(wavy.diagnostics.map((d) => [d.code, d.severity])).toContainEqual([
      'text.underline-style',
      'approximation',
    ]);
  });

  test('the font resolution lists no substitute failures when the packaged faces load', async () => {
    const result = await exportPdf(missingFontsDeck(), BEST);
    expect(result.fontResolution.substituteFailures).toEqual([]);
    expect(result.diagnostics.map((d) => d.code)).not.toContain('font.substitute-failed');
  });

  test('the strict error names the slide and the feature', async () => {
    const error = await exportPdf(deckWithHidden()).catch((e: unknown) => e);
    expect(error).toBeInstanceOf(PdfFidelityError);
    expect((error as Error).message).toContain('slide 3: fill.pattern');
  });

  test('a diagram without its saved drawing says why it is not drawn', async () => {
    const files = unzipSync(deck('diagrams'));
    delete files['ppt/diagrams/drawing1.xml'];
    const result = await exportPdf(zipSync(files), BEST);
    const missing = result.diagnostics.filter((d) => d.code === 'frame.diagram');
    expect(missing.map((d) => [d.slideNumber, d.message])).toEqual([
      [1, 'A diagram is not drawn: it has no saved drawing.'],
    ]);
  });

  test('a subscript run reads as its own text item, below the baseline', async () => {
    const bytes = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 6 * IN,
            cy: IN,
            text: para(
              run('H', 'sz="1800"') +
                run('2', 'sz="1800" baseline="-25000"') +
                run('O', 'sz="1800"')
            ),
          }),
        },
      ],
    });
    const [page] = await pdfText(await exportPdf(bytes, BEST));
    expect(page.map((i) => i.str)).toEqual(['H', '2', 'O']);
    const [h, two, o] = page;
    // pdf.js y is up: the subscript sits lower, and the runs around it share one baseline.
    expect(two.transform[5]).toBeLessThan(h.transform[5] - 1);
    expect(o.transform[5]).toBeCloseTo(h.transform[5], 3);
    // The subscript is drawn at its smaller size, left to right after the H.
    expect(two.transform[3]).toBeLessThan(h.transform[3]);
    expect(two.transform[4]).toBeGreaterThan(h.transform[4]);
    expect(o.transform[4]).toBeGreaterThan(two.transform[4]);
  });

  test('stacked vertical letters draw upright, one below the other', async () => {
    const bytes = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            x: IN,
            y: IN,
            cx: IN,
            cy: 4 * IN,
            bodyPr: '<a:bodyPr vert="wordArtVert" wrap="square"/>',
            text: para(run('ABC', 'sz="2400"')),
          }),
        },
      ],
    });
    const result = await exportPdf(bytes, BEST);
    expect(result.diagnostics.map((d) => d.code)).not.toContain('text.upright');
    const [page] = await pdfText(result);
    const letters = ['A', 'B', 'C'].map((c) => page.find((i) => i.str === c));
    for (const item of letters) {
      if (!item) throw new Error('letter not found');
      // Upright: no turn in the text transform.
      expect(item.transform[0]).toBeGreaterThan(0);
      expect(Math.abs(item.transform[1])).toBeLessThan(0.01);
      expect(Math.abs(item.transform[2])).toBeLessThan(0.01);
      expect(item.transform[3]).toBeGreaterThan(0);
    }
    const [a, b, c] = letters as { transform: number[] }[];
    // One column: the same center line, and each letter lower than the one before.
    expect(b.transform[5]).toBeLessThan(a.transform[5]);
    expect(c.transform[5]).toBeLessThan(b.transform[5]);
    expect(Math.abs(b.transform[4] - a.transform[4])).toBeLessThan(5);
  });

  test('diagnostic codes are fixed strings, not file text', async () => {
    const result = await exportPdf(missingFontsDeck(), BEST);
    for (const d of result.diagnostics) expect(d.code).toMatch(/^[a-z]+(\.[a-z-]+)+$/);
  });
});
