/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Soft edges: the body paints as one group through a mask that fades toward the outline.

import { describe, expect, test } from 'bun:test';
import { decodePDFRawStream, PDFDict, PDFDocument, PDFName, PDFRawStream } from 'pdf-lib';
import { blurSigma } from '@pptx-editor.dev/core/output';
import { softEdgeSpread } from '@pptx-editor.dev/pro/drawing';
import { buildPptx, shape, solid } from '../../../../scripts/fixtures/builder.ts';
import { softEdgeBands } from '../paint-effects.ts';
import { exportPdf, type PdfExportResult } from '../index.ts';

const BEST = { fidelityPolicy: 'best-effort' } as const;
const IN = 914_400;
const PT = 12_700;

/** The decoded content of every luminosity soft mask group (a form with a DeviceGray group). */
async function softMasks(bytes: Uint8Array): Promise<string[]> {
  const doc = await PDFDocument.load(bytes);
  const out: string[] = [];
  for (const [, obj] of doc.context.enumerateIndirectObjects()) {
    if (!(obj instanceof PDFRawStream)) continue;
    const group = obj.dict.lookupMaybe(PDFName.of('Group'), PDFDict);
    if (group?.get(PDFName.of('CS')) !== PDFName.of('DeviceGray')) continue;
    const raw = obj.dict.has(PDFName.of('Filter'))
      ? decodePDFRawStream(obj).decode()
      : obj.contents;
    out.push(new TextDecoder('latin1').decode(raw));
  }
  return out;
}

/** The decoded content streams of a PDF, the page's and every form's (pictures left out). */
async function contentText(bytes: Uint8Array): Promise<string> {
  const doc = await PDFDocument.load(bytes);
  let text = '';
  for (const [, obj] of doc.context.enumerateIndirectObjects()) {
    if (!(obj instanceof PDFRawStream)) continue;
    if (obj.dict.get(PDFName.of('Subtype')) === PDFName.of('Image')) continue;
    const raw = obj.dict.has(PDFName.of('Filter'))
      ? decodePDFRawStream(obj).decode()
      : obj.contents;
    text += new TextDecoder('latin1').decode(raw);
  }
  return text;
}

const soft = (
  cx: number,
  cy: number,
  rad: number,
  extra: { fill?: string; line?: string; prst?: string } = {}
): Promise<PdfExportResult> =>
  exportPdf(
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            x: IN,
            y: IN,
            cx,
            cy,
            fill: solid('4472C4'),
            ...extra,
            effects: `<a:effectLst><a:softEdge rad="${rad}"/></a:effectLst>`,
          }),
        },
      ],
    }),
    BEST
  );

describe('soft edge bands', () => {
  test('the alpha rises from clear on the outline to opaque past the blur', () => {
    const { spread, sigma } = softEdgeSpread(10 * PT);
    const { reach, alpha } = softEdgeBands(-spread, sigma);
    expect(reach.length).toBeGreaterThan(8);
    expect(reach.length).toBeLessThanOrEqual(48);
    expect(reach[reach.length - 1]).toBeCloseTo(-spread + 3 * blurSigma(10 * PT));
    for (let k = 1; k < alpha.length; k += 1) expect(alpha[k]).toBeGreaterThanOrEqual(alpha[k - 1]);
    expect(alpha[0]).toBeLessThan(0.05);
    expect(alpha[alpha.length - 1]).toBeGreaterThan(0.95);
    // The half level is at the inset, 0.84 of the radius inside the outline.
    const half = reach.findIndex((_, k) => alpha[k] >= 0.5);
    expect(reach[half]).toBeGreaterThanOrEqual(-spread);
  });

  test('a blur under half a point is a hard step at the inset', () => {
    const { reach, alpha } = softEdgeBands(10 * PT, 0);
    reach.forEach((r, k) => {
      const middle = ((k ? reach[k - 1] : 0) + r) / 2;
      expect(alpha[k]).toBe(middle >= 10 * PT ? 1 : 0);
    });
  });
});

/** The line widths and stroke grays of a mask, in order. */
const strokesOf = (mask: string): { width: number; gray: number }[] => {
  const out: { width: number; gray: number }[] = [];
  let gray = 0;
  let width = 0;
  for (const m of mask.matchAll(/([\d.]+) (G|w)\n|\nS\n/g)) {
    if (m[2] === 'G') gray = Number(m[1]);
    else if (m[2] === 'w') width = Number(m[1]);
    else out.push({ width, gray });
  }
  return out;
};

const LINE_12PT = `<a:ln w="${12 * PT}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>`;

describe('soft edge paint', () => {
  test('a 12 pt outline fades from its outer edge, as the SVG filter reads the stroke', async () => {
    const result = await soft(3 * IN, 2 * IN, 2.5 * PT, { line: LINE_12PT });
    const masks = await softMasks(result.bytes);
    expect(masks).toHaveLength(1);
    const { spread, sigma } = softEdgeSpread(2.5 * PT);
    const { reach } = softEdgeBands(-spread, sigma);
    const extent = reach[reach.length - 1];
    // The whole fade lies on the outline: it is wider than the soft edge reaches.
    expect(extent).toBeLessThan(6 * PT);
    // The outline is clipped to the outside of the fill, then stroked from its full width in.
    expect(masks[0]).toContain('W*\nn');
    const strokes = strokesOf(masks[0]);
    expect(strokes).toHaveLength(reach.length + 1);
    expect(strokes[0].width).toBeCloseTo(12 * PT, 0);
    expect(strokes[0].gray).toBeLessThan(0.05);
    for (let k = 1; k < strokes.length; k += 1) {
      expect(strokes[k].width).toBeLessThan(strokes[k - 1].width);
      expect(strokes[k].gray).toBeGreaterThanOrEqual(strokes[k - 1].gray);
    }
    // Past the fade, the inner part of the outline and the fill stay opaque.
    expect(strokes[strokes.length - 1]).toEqual({
      width: expect.closeTo(12 * PT - 2 * extent, 0),
      gray: 1,
    });
  });

  test('an outline with no fill fades from both of its edges', async () => {
    const result = await soft(3 * IN, 2 * IN, 2.5 * PT, {
      fill: '<a:noFill/>',
      line: LINE_12PT,
    });
    const masks = await softMasks(result.bytes);
    expect(masks).toHaveLength(1);
    expect(masks[0]).not.toContain('W*\nn');
    expect(strokesOf(masks[0])[0].width).toBeCloseTo(12 * PT, 0);
  });

  test('a filled shape paints through a mask of its outline bands', async () => {
    const result = await soft(3 * IN, 2 * IN, 10 * PT);
    const masks = await softMasks(result.bytes);
    expect(masks).toHaveLength(1);
    const { spread, sigma } = softEdgeSpread(10 * PT);
    const { reach } = softEdgeBands(-spread, sigma);
    // One white fill, then one stroke per band, widest first.
    expect(masks[0].match(/\nS\n/g)).toHaveLength(reach.length);
    const widths = [...masks[0].matchAll(/([\d.]+) w/g)].map((m) => Number(m[1]));
    expect(widths[0]).toBeCloseTo(2 * reach[reach.length - 1], 0);
    expect(widths[widths.length - 1]).toBeCloseTo(2 * reach[0], 0);
    expect(await contentText(result.bytes)).toContain('0.4470588');
    expect(result.diagnostics.find((d) => d.code === 'effect.soft-edge')?.message).toContain(
      'distance to the painted edge'
    );
  });

  test('a soft edge past half the narrower side paints nothing', async () => {
    const result = await soft(3 * IN, 10 * PT, 20 * PT);
    expect(await softMasks(result.bytes)).toHaveLength(0);
    // The fill color (#4472C4) is set nowhere, in the page or in a form.
    const text = await contentText(result.bytes);
    expect(text).not.toContain('0.4470588');
  });
});

/** The end points of every stroked subpath of a mask, in slide EMU. */
const strokedPoints = (mask: string): { x: number; y: number }[] => {
  const out: { x: number; y: number }[] = [];
  let pending: { x: number; y: number }[] = [];
  for (const line of mask.split('\n')) {
    const t = line.trim().split(' ');
    const op = t[t.length - 1];
    if (op === 'm' || op === 'l' || op === 'c')
      pending.push({ x: Number(t[t.length - 3]), y: Number(t[t.length - 2]) });
    else if (op === 'S') {
      out.push(...pending);
      pending = [];
    } else if (op === 'f*' || op === 'f' || op === 'n') pending = [];
  }
  return out;
};

describe('soft edge of overlapping filled paths', () => {
  test('a smiley face fades only its outer edge, with no rings around its eyes', async () => {
    const result = await soft(2 * IN, 2 * IN, 5 * PT, { prst: 'smileyFace' });
    const masks = await softMasks(result.bytes);
    expect(masks).toHaveLength(1);
    const { spread, sigma } = softEdgeSpread(5 * PT);
    const { reach } = softEdgeBands(-spread, sigma);
    // The face and the eyes fill the mask one path at a time, so the eyes stay white.
    expect(masks[0].match(/\nf\*\n/g)).toHaveLength(2);
    // Only the face outline takes bands: the eyes lie inside the face.
    expect(masks[0].match(/\nS\n/g)).toHaveLength(reach.length);
    const cx = IN + IN;
    const cy = IN + IN;
    for (const pt of strokedPoints(masks[0]))
      expect(Math.hypot(pt.x - cx, pt.y - cy)).toBeGreaterThan(0.9 * IN);
  });

  test('a can fades no seam where its top ellipse joins its body', async () => {
    const result = await soft(2 * IN, 3 * IN, 5 * PT, { prst: 'can' });
    const masks = await softMasks(result.bytes);
    expect(masks).toHaveLength(1);
    const points = strokedPoints(masks[0]);
    expect(points.length).toBeGreaterThan(0);
    // The top ellipse spans 0.5 in from the top, the bottom arc 0.5 in from the bottom. Between
    // them, the outer edge is the two sides; the lower half of the top ellipse is inside.
    const tol = 0.02 * IN;
    for (const pt of points) {
      if (pt.y <= IN + 0.25 * IN + tol || pt.y >= IN + 2.75 * IN - tol) continue;
      expect(Math.min(Math.abs(pt.x - IN), Math.abs(pt.x - 3 * IN))).toBeLessThan(tol);
    }
  });
});

/** A soft-edged shape of custom geometry: `paths` are `a:path` elements, in a 4 in square. */
const softCustom = (paths: string, rad: number): Promise<PdfExportResult> =>
  exportPdf(
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes:
            '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Freeform"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
            `<p:spPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${4 * IN}" cy="${4 * IN}"/></a:xfrm>` +
            '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>' +
            `<a:pathLst>${paths}</a:pathLst></a:custGeom>` +
            '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>' +
            `<a:effectLst><a:softEdge rad="${rad}"/></a:effectLst></p:spPr></p:sp>`,
        },
      ],
    }),
    BEST
  );

const pt = (x: number, y: number): string => `<a:pt x="${x}" y="${y}"/>`;
/** A closed `a:path` through `points`, in a 100,000 unit square. */
const polygon = (points: [number, number][]): string =>
  '<a:path w="100000" h="100000">' +
  `<a:moveTo>${pt(...points[0])}</a:moveTo>` +
  points
    .slice(1)
    .map((q) => `<a:lnTo>${pt(...q)}</a:lnTo>`)
    .join('') +
  '<a:close/></a:path>';

describe('soft edge of a complex shape', () => {
  test('30,000 small paths export with a hard edge, in bounded time and size', async () => {
    const paths: string[] = [];
    for (let i = 0; i < 30_000; i += 1) {
      const x = (i % 200) * 500;
      const y = Math.floor(i / 200) * 600;
      paths.push(
        polygon([
          [x, y],
          [x + 300, y],
          [x + 300, y + 300],
          [x, y + 300],
        ])
      );
    }
    const start = performance.now();
    const result = await softCustom(paths.join(''), 63_500);
    expect(performance.now() - start).toBeLessThan(15_000);
    expect(result.bytes.length).toBeLessThan(8 * 1024 * 1024);
    expect(result.diagnostics.map((d) => d.code)).toContain('effect.soft-edge-complex');
    const masks = await softMasks(result.bytes);
    expect(masks).toHaveLength(1);
    expect(masks[0].match(/\nS\n/g)).toBeNull();
  }, 30_000);

  test('300 concentric squares keep a solid body: no bands inside the union', async () => {
    const paths: string[] = [];
    for (let i = 0; i < 300; i += 1) {
      const a = i * 150;
      const b = 100_000 - a;
      paths.push(
        polygon([
          [a, a],
          [b, a],
          [b, b],
          [a, b],
        ])
      );
    }
    const result = await softCustom(paths.join(''), 5 * PT);
    expect(result.diagnostics.map((d) => d.code)).toContain('effect.soft-edge-complex');
    const masks = await softMasks(result.bytes);
    expect(masks).toHaveLength(1);
    // Every square fills white; no band strokes an edge inside the outer square.
    expect(masks[0].match(/\nf\*\n/g)).toHaveLength(300);
    expect(masks[0].match(/\nS\n/g)).toBeNull();
  });

  test('a path of 45,000 points fades with fewer bands, and the PDF stays small', async () => {
    const n = 45_000;
    const points = Array.from({ length: n }, (_, i): [number, number] => {
      const t = (2 * Math.PI * i) / n;
      return [Math.round(50_000 + 45_000 * Math.cos(t)), Math.round(50_000 + 45_000 * Math.sin(t))];
    });
    const result = await softCustom(polygon(points), 20 * PT);
    expect(result.bytes.length).toBeLessThan(4 * 1024 * 1024);
    const diag = result.diagnostics.find((d) => d.code === 'effect.soft-edge-complex');
    expect(diag?.message).toContain('fewer bands');
    const masks = await softMasks(result.bytes);
    expect(masks).toHaveLength(1);
    const strokes = masks[0].match(/\nS\n/g) ?? [];
    expect(strokes.length).toBeGreaterThanOrEqual(4);
    expect(strokes.length).toBeLessThan(10);
  }, 30_000);
});
