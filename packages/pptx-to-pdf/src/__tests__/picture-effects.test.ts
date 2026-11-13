/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Picture color effects and tiled pictures in the PDF: the color function computes the SVG
// filter's colors (`blipColorSteps`), a viewer applies it to the decoded pixels, and the tile
// pattern places the tiles of `tileGeometry`.

import { describe, expect, test } from 'bun:test';
import { PNG } from 'pngjs';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { decodePDFRawStream, PDFArray, PDFDict, PDFDocument, PDFName, PDFRawStream } from 'pdf-lib';
import type { BlipEffect } from '@pptx-editor.dev/core/cascade';
import { applyBlipColorSteps, blipColorSteps } from '@pptx-editor.dev/core/output';
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';
import { exportPdf } from '../index.ts';
import { colorStepsCode, MAX_COLOR_STEPS } from '../paint-picture.ts';

const IN = 914_400;
const BEST = { fidelityPolicy: 'best-effort' } as const;
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/**
 * A PostScript calculator (ISO 32000-1, 7.10.5) with the operators that `colorStepsCode`
 * writes, apart from the writer, so a test can run its code.
 */
function runCalculator(code: string, inputs: number[]): number[] {
  const tokens = code.match(/\{|\}|[^\s{}]+/g) ?? [];
  let at = 0;
  type Proc = (string | Proc)[];
  const parse = (): Proc => {
    const out: Proc = [];
    while (at < tokens.length) {
      const t = tokens[at++]!;
      if (t === '{') out.push(parse());
      else if (t === '}') return out;
      else out.push(t);
    }
    return out;
  };
  at = 1; // The outer brace.
  const program = parse();
  const stack = [...inputs];
  const exec = (proc: Proc): void => {
    for (let i = 0; i < proc.length; i += 1) {
      const t = proc[i]!;
      if (Array.isArray(t)) {
        const next = proc[i + 1];
        if (next !== 'if') throw new Error('only if');
        if (stack.pop()) exec(t);
        i += 1;
        continue;
      }
      const b = (): number => stack.pop()!;
      switch (t) {
        case 'mul':
          stack.push(b() * b());
          break;
        case 'add':
          stack.push(b() + b());
          break;
        case 'exch': {
          const y = b();
          const x = b();
          stack.push(y, x);
          break;
        }
        case 'dup':
          stack.push(stack[stack.length - 1]!);
          break;
        case 'pop':
          b();
          break;
        case 'lt': {
          const y = b();
          stack.push(b() < y ? 1 : 0);
          break;
        }
        case 'gt': {
          const y = b();
          stack.push(b() > y ? 1 : 0);
          break;
        }
        case 'roll': {
          const j = b();
          const n = b();
          const top = stack.splice(stack.length - n, n);
          const k = ((j % n) + n) % n;
          stack.push(...top.slice(n - k), ...top.slice(0, n - k));
          break;
        }
        default:
          if (!/^-?\d+(\.\d+)?$/.test(t)) throw new Error(`operator ${t}`);
          stack.push(Number(t));
      }
    }
  };
  exec(program);
  return stack;
}

const dark = { r: 0x38, g: 0x57, b: 0x23, a: 1 };
const light = { r: 255, g: 255, b: 255, a: 1 };
const EFFECTS: Record<string, BlipEffect[]> = {
  grayscale: [{ kind: 'grayscale' }],
  duotone: [{ kind: 'duotone', dark, light }],
  lum: [{ kind: 'lum', bright: 0.2, contrast: -0.3 }],
  contrast: [{ kind: 'lum', bright: -0.1, contrast: 0.6 }],
  biLevel: [{ kind: 'biLevel', threshold: 0.5 }],
  chain: [
    { kind: 'lum', bright: 0.4, contrast: 0.5 },
    { kind: 'duotone', dark, light },
    { kind: 'alphaModFix', amount: 0.5 },
  ],
};

describe('the color function of picture effects', () => {
  for (const [name, effects] of Object.entries(EFFECTS))
    test(`computes the SVG filter's colors: ${name}`, () => {
      const steps = blipColorSteps(effects);
      const code = colorStepsCode(steps, 3);
      for (const rgb of [
        [0, 0, 0],
        [1, 1, 1],
        [1, 0, 0],
        [0.2, 0.7, 0.4],
        [0.49, 0.51, 0.5],
        [0.9, 0.1, 0.6],
      ] as [number, number, number][]) {
        const want = applyBlipColorSteps(steps, rgb);
        const got = runCalculator(code, rgb);
        expect(got).toHaveLength(3);
        got.forEach((v, i) => expect(v).toBeCloseTo(want[i]!, 5));
        // A gray image reads one component, as all three.
        const gray = runCalculator(colorStepsCode(steps, 1), [rgb[0]]);
        const wantGray = applyBlipColorSteps(steps, [rgb[0], rgb[0], rgb[0]]);
        gray.forEach((v, i) => expect(v).toBeCloseTo(wantGray[i]!, 5));
      }
    });

  test('writes plain decimal numbers, also for values from the file out of range', () => {
    const code = colorStepsCode(
      blipColorSteps([{ kind: 'lum', bright: 1e30, contrast: Number.NaN }]),
      3
    );
    expect(code).not.toMatch(/e[+-]?\d|NaN|Infinity/);
  });
});

/** A 2 × 1 PNG: a red and a teal pixel, opaque. */
function colorPng(): Uint8Array {
  const png = new PNG({ width: 2, height: 1 });
  png.data.set([255, 0, 0, 255, 32, 160, 128, 255]);
  return PNG.sync.write(png);
}

const pictureDeck = (blip: string, fill = '<a:stretch><a:fillRect/></a:stretch>') =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          `<p:pic><p:nvPicPr><p:cNvPr id="4" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
          `<p:blipFill><a:blip r:embed="rIdImg">${blip}</a:blip>${fill}</p:blipFill>` +
          `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${IN}" cy="${IN}"/></a:xfrm>` +
          `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`,
        rels: [['rIdImg', `${R}/image`, '../media/img.png']],
      },
    ],
    media: [{ name: 'img.png', bytes: colorPng(), contentType: 'image/png' }],
  });

/** The RGB pixels of the first image that pdf.js draws on page 1, decoded through its color space. */
async function drawnPixels(pdf: Uint8Array): Promise<number[][]> {
  const doc = await getDocument({ data: pdf.slice(), useSystemFonts: false }).promise;
  const page = await doc.getPage(1);
  const { fnArray, argsArray } = await page.getOperatorList();
  const at = fnArray.indexOf(OPS.paintImageXObject);
  expect(at).toBeGreaterThan(-1);
  const name = (argsArray[at] as [string])[0];
  const image = (await new Promise((resolve) => page.objs.get(name, resolve))) as {
    width: number;
    height: number;
    data: Uint8ClampedArray;
    kind: number;
  };
  const channels = image.data.length / (image.width * image.height);
  const out: number[][] = [];
  for (let i = 0; i < image.width * image.height; i += 1)
    out.push([...image.data.slice(i * channels, i * channels + 3)]);
  return out;
}

describe('picture color effects in the PDF', () => {
  const cases: [string, string, BlipEffect[]][] = [
    ['grayscale', '<a:grayscl/>', EFFECTS.grayscale!],
    [
      'duotone',
      '<a:duotone><a:srgbClr val="385723"/><a:srgbClr val="FFFFFF"/></a:duotone>',
      EFFECTS.duotone!,
    ],
    ['brightness and contrast', '<a:lum bright="20000" contrast="-30000"/>', EFFECTS.lum!],
    ['black and white', '<a:biLevel thresh="50000"/>', EFFECTS.biLevel!],
  ];
  for (const [name, blip, effects] of cases)
    test(`a viewer draws the effect's colors: ${name}`, async () => {
      const result = await exportPdf(pictureDeck(blip), BEST);
      expect(result.diagnostics.map((d) => d.code)).not.toContain('picture.color-effects');
      const pixels = await drawnPixels(result.bytes);
      const steps = blipColorSteps(effects);
      const source = [
        [255, 0, 0],
        [32, 160, 128],
      ];
      pixels.forEach((px, i) => {
        const want = applyBlipColorSteps(
          steps,
          source[i]!.map((v) => v / 255) as [number, number, number]
        );
        px.forEach((v, c) => expect(Math.abs(v - want[c]! * 255)).toBeLessThanOrEqual(1));
      });
    });

  test('at most MAX_COLOR_STEPS steps are drawn, and more are reported', async () => {
    const result = await exportPdf(pictureDeck('<a:grayscl/>'.repeat(MAX_COLOR_STEPS + 1)), BEST);
    expect(result.diagnostics.map((d) => d.code)).toContain('picture.color-effects');
    const pixels = await drawnPixels(result.bytes);
    // Grayscale still applies.
    expect(pixels[0]![0]).toBe(pixels[0]![1]);
  });

  test('an effect that the SVG paint does not draw is reported', async () => {
    const result = await exportPdf(
      pictureDeck('<a:clrChange><a:clrFrom><a:srgbClr val="FF0000"/></a:clrFrom></a:clrChange>'),
      BEST
    );
    expect(result.diagnostics.map((d) => d.code)).toContain('picture.color-effects');
  });

  test('the recolored picture shares the original data, once per image and effects', async () => {
    const result = await exportPdf(pictureDeck('<a:grayscl/>'), BEST);
    const doc = await PDFDocument.load(result.bytes);
    const spaces: string[] = [];
    for (const [, obj] of doc.context.enumerateIndirectObjects()) {
      if (!(obj instanceof PDFRawStream)) continue;
      if (obj.dict.get(PDFName.of('Subtype')) !== PDFName.of('Image')) continue;
      const space = obj.dict.get(PDFName.of('ColorSpace'));
      spaces.push(space instanceof PDFArray ? String(space.get(0)) : String(space));
    }
    // The original (with no user here, still embedded), its alpha, and the DeviceN copy.
    expect(spaces).toContain('/DeviceN');
    expect(spaces.filter((s) => s === '/DeviceN')).toHaveLength(1);
  });
});

describe('tiled pictures in the PDF', () => {
  /** The tiling patterns that draw a picture: their steps and matrix. */
  async function tilePatterns(pdf: Uint8Array) {
    const doc = await PDFDocument.load(pdf);
    const out: { step: number[]; matrix: number[]; draws: number }[] = [];
    for (const [, obj] of doc.context.enumerateIndirectObjects()) {
      if (!(obj instanceof PDFRawStream)) continue;
      if (obj.dict.get(PDFName.of('PatternType'))?.toString() !== '1') continue;
      const content = new TextDecoder().decode(
        obj.dict.has(PDFName.of('Filter')) ? decodePDFRawStream(obj).decode() : obj.contents
      );
      if (!content.includes('/Im Do')) continue;
      const num = (key: string): number => Number(obj.dict.get(PDFName.of(key))!.toString());
      const matrix = obj.dict.lookup(PDFName.of('Matrix'), PDFArray).asArray().map(Number);
      out.push({
        step: [num('XStep'), num('YStep')],
        matrix,
        draws: content.match(/\/Im Do/g)!.length,
      });
    }
    return out;
  }

  test('tiles at the natural size × scale, placed at the alignment point and offset', async () => {
    // 2 × 1 pixels at 72 DPI (no pHYs): 25400 × 12700 EMU; 75 % scale; centered; mirrored.
    const tile = '<a:tile tx="12700" ty="-6350" sx="75000" sy="75000" flip="xy" algn="ctr"/>';
    const result = await exportPdf(pictureDeck('', tile), BEST);
    expect(result.diagnostics.map((d) => d.code)).not.toContain('picture.tile');
    const [pattern, ...rest] = await tilePatterns(result.bytes);
    expect(rest).toHaveLength(0);
    const tw = 19_050;
    const th = 9_525;
    // A mirrored cell holds 2 × 2 tiles.
    expect(pattern!.step).toEqual([2 * tw, 2 * th]);
    expect(pattern!.draws).toBe(4);
    const [a, b, c, d, e, f] = pattern!.matrix;
    expect([a, b, c, d]).toEqual([1, 0, 0, 1]);
    expect(e).toBeCloseTo(0.5 * IN - 0.5 * tw + 12_700, 3);
    expect(f).toBeCloseTo(0.5 * IN - 0.5 * th - 6_350, 3);
  });

  test('one tile per cell without mirroring, from the top-left corner', async () => {
    const tile = '<a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/>';
    const [pattern] = await tilePatterns((await exportPdf(pictureDeck('', tile), BEST)).bytes);
    expect(pattern!.step).toEqual([25_400, 12_700]);
    expect(pattern!.draws).toBe(1);
    expect(pattern!.matrix).toEqual([1, 0, 0, 1, 0, 0]);
  });

  test('a tiled picture with effects tiles the recolored picture', async () => {
    const tile = '<a:tile tx="0" ty="0" sx="100000" sy="100000" flip="x" algn="tl"/>';
    const result = await exportPdf(pictureDeck('<a:grayscl/>', tile), BEST);
    const codes = result.diagnostics.map((d) => d.code);
    expect(codes).not.toContain('picture.tile');
    expect(codes).not.toContain('picture.color-effects');
    const doc = await PDFDocument.load(result.bytes);
    let tiled = false;
    for (const [, obj] of doc.context.enumerateIndirectObjects()) {
      if (!(obj instanceof PDFRawStream)) continue;
      if (obj.dict.get(PDFName.of('PatternType'))?.toString() !== '1') continue;
      const xobjects = obj.dict
        .lookup(PDFName.of('Resources'), PDFDict)
        .lookup(PDFName.of('XObject'), PDFDict);
      const image = doc.context.lookup(xobjects.get(PDFName.of('Im')));
      if (!(image instanceof PDFRawStream)) throw new Error('no tile image');
      const space = image.dict.lookup(PDFName.of('ColorSpace'), PDFArray);
      expect(space.get(0)).toBe(PDFName.of('DeviceN'));
      tiled = true;
    }
    expect(tiled).toBe(true);
  });
});
