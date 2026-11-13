/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Tiling pattern fills under a normal transform. A viewer inverts the current transform before
// it fills with a tiling pattern, and Poppler refuses the fill when the determinant is below
// 1e-6. Slide EMU (1/12700 pt per unit) is far below that, so the content that fills with a
// pattern is written in a space of about one point per unit.

import { describe, expect, test } from 'bun:test';
import { PNG } from 'pngjs';
import {
  decodePDFRawStream,
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFName,
  PDFRawStream,
  PDFRef,
} from 'pdf-lib';
import { buildPptx, shape } from '../../../../scripts/fixtures/builder.ts';
import { exportPdf } from '../index.ts';

const IN = 914_400;
const BEST = { fidelityPolicy: 'best-effort' } as const;
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MIN_DET = 1e-6;

type Matrix = [number, number, number, number, number, number];
const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];

/** `m` then `n`: the PDF order of `cm` (ISO 32000-1, 8.3.4). */
const times = (m: Matrix, n: Matrix): Matrix => [
  m[0] * n[0] + m[1] * n[2],
  m[0] * n[1] + m[1] * n[3],
  m[2] * n[0] + m[3] * n[2],
  m[2] * n[1] + m[3] * n[3],
  m[4] * n[0] + m[5] * n[2] + n[4],
  m[4] * n[1] + m[5] * n[3] + n[5],
];
const det = (m: Matrix): number => m[0] * m[3] - m[1] * m[2];

/** The operands and operators of a content stream, without strings, arrays, and dictionaries. */
function tokens(src: string): string[] {
  const out: string[] = [];
  let i = 0;
  while (i < src.length) {
    const ch = src[i]!;
    if (/\s/.test(ch)) i += 1;
    else if (ch === '%') while (i < src.length && src[i] !== '\n') i += 1;
    else if (ch === '(') {
      let depth = 0;
      for (; i < src.length; i += 1) {
        if (src[i] === '\\') i += 1;
        else if (src[i] === '(') depth += 1;
        else if (src[i] === ')' && --depth === 0) break;
      }
      i += 1;
      out.push('(str)');
    } else if (src.startsWith('<<', i) || src.startsWith('>>', i)) {
      out.push(src.slice(i, i + 2));
      i += 2;
    } else if (ch === '<') {
      i = src.indexOf('>', i) + 1;
      out.push('(str)');
    } else if (ch === '[' || ch === ']' || ch === '{' || ch === '}') {
      out.push(ch);
      i += 1;
    } else {
      const m = /^\/?[^\s()<>[\]{}/%]*/.exec(src.slice(i))![0];
      out.push(m);
      i += m.length || 1;
    }
  }
  return out;
}

function decoded(stream: PDFRawStream): string {
  const raw = stream.dict.has(PDFName.of('Filter'))
    ? decodePDFRawStream(stream).decode()
    : stream.contents;
  return new TextDecoder('latin1').decode(raw);
}

function matrixOf(dict: PDFDict, key: string): Matrix {
  const arr = dict.lookupMaybe(PDFName.of(key), PDFArray);
  return arr ? (arr.asArray().map((n) => Number(n.toString())) as Matrix) : IDENTITY;
}

/**
 * Every fill with a tiling pattern on page 1, through form XObjects: the current transform is
 * tracked through `q`, `Q`, `cm`, and `Do`.
 */
async function patternFills(bytes: Uint8Array): Promise<Matrix[]> {
  const doc = await PDFDocument.load(bytes);
  const page = doc.getPages()[0]!;
  const fills: Matrix[] = [];
  const walk = (content: string, resources: PDFDict | undefined, start: Matrix): void => {
    const stack: { ctm: Matrix; pattern: boolean }[] = [];
    let ctm = start;
    let pattern = false;
    let operands: string[] = [];
    for (const t of tokens(content)) {
      if (/^[-+.\d]/.test(t) || t.startsWith('/') || t === '(str)' || /^[[\]{}]$|^<<|>>$/.test(t)) {
        operands.push(t);
        continue;
      }
      if (t === 'q') stack.push({ ctm, pattern });
      else if (t === 'Q') ({ ctm, pattern } = stack.pop()!);
      else if (t === 'cm') ctm = times(operands.slice(-6).map(Number) as Matrix, ctm);
      else if (t === 'scn' || t === 'SCN' || t === 'sc' || t === 'rg' || t === 'g' || t === 'k') {
        const name = operands.at(-1);
        const ref = name?.startsWith('/')
          ? resources?.lookupMaybe(PDFName.of('Pattern'), PDFDict)?.get(PDFName.of(name.slice(1)))
          : undefined;
        const obj = ref instanceof PDFRef ? doc.context.lookup(ref) : undefined;
        pattern = obj instanceof PDFRawStream;
      } else if (['f', 'F', 'f*', 'B', 'B*', 'b', 'b*'].includes(t) && pattern) fills.push(ctm);
      else if (t === 'Do') {
        const name = operands.at(-1)!.slice(1);
        const ref = resources?.lookupMaybe(PDFName.of('XObject'), PDFDict)?.get(PDFName.of(name));
        const form = ref instanceof PDFRef ? doc.context.lookup(ref) : undefined;
        if (
          form instanceof PDFRawStream &&
          String(form.dict.get(PDFName.of('Subtype'))) === '/Form'
        )
          walk(
            decoded(form),
            form.dict.lookupMaybe(PDFName.of('Resources'), PDFDict),
            times(matrixOf(form.dict, 'Matrix'), ctm)
          );
      }
      operands = [];
    }
  };
  const contents = page.node.Contents();
  const streams =
    contents instanceof PDFArray
      ? contents.asArray().map((r) => doc.context.lookup(r) as PDFRawStream)
      : [contents as unknown as PDFRawStream];
  walk(streams.map(decoded).join('\n'), page.node.Resources(), IDENTITY);
  return fills;
}

function png(): Uint8Array {
  const img = new PNG({ width: 2, height: 1 });
  img.data.set([255, 0, 0, 255, 0, 0, 255, 255]);
  return new Uint8Array(PNG.sync.write(img));
}

describe('pattern fills under a normal transform', () => {
  const expectNormal = (fills: Matrix[]): void => {
    expect(fills.length).toBeGreaterThan(0);
    for (const ctm of fills) expect(Math.abs(det(ctm))).toBeGreaterThanOrEqual(MIN_DET);
  };

  test('a tiled picture fills under a transform of about one point per unit', async () => {
    const tile = '<a:tile tx="0" ty="0" sx="100000" sy="100000" flip="xy" algn="tl"/>';
    const deck = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes:
            `<p:pic><p:nvPicPr><p:cNvPr id="4" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
            `<p:blipFill><a:blip r:embed="rIdImg"/>${tile}</p:blipFill>` +
            `<p:spPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${IN}" cy="${IN}"/></a:xfrm>` +
            `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`,
          rels: [['rIdImg', `${R}/image`, '../media/img.png']],
        },
      ],
      media: [{ name: 'img.png', bytes: png(), contentType: 'image/png' }],
    });
    expectNormal(await patternFills((await exportPdf(deck, BEST)).bytes));
  });

  test('a preset pattern fills under a transform of about one point per unit', async () => {
    const patt =
      '<a:pattFill prst="dkDnDiag"><a:fgClr><a:srgbClr val="C00000"/></a:fgClr>' +
      '<a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>';
    const deck = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 2 * IN,
            cy: IN,
            rot: 2_700_000,
            fill: patt,
            effects:
              '<a:effectLst><a:reflection blurRad="0" stA="50000" endA="0" endPos="90000" ' +
              'dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/></a:effectLst>',
          }),
        },
      ],
    });
    // The page draws the shape, and the reflection form draws its copy.
    const fills = await patternFills((await exportPdf(deck, BEST)).bytes);
    expect(fills.length).toBeGreaterThanOrEqual(2);
    expectNormal(fills);
  });
});
