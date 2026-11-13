/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The effects of a group (`p:grpSpPr/a:effectLst`) in the PDF export, against the editor's SVG
// paint: one shadow of all children together, under the same transform as the SVG shadow copy,
// placed by the group's box under its turn, flips, and child space.

import { describe, expect, test } from 'bun:test';
import { Window } from 'happy-dom';
import { decodePDFRawStream, PDFDict, PDFDocument, PDFName, PDFRawStream } from 'pdf-lib';
import { openPresentation } from '@pptx-editor.dev/core/editor';
import { apply, type LayoutItem, type Matrix } from '@pptx-editor.dev/core/layout';
import { paintSlide } from '@pptx-editor.dev/core/output';
import { installDrawing } from '@pptx-editor.dev/pro/drawing';
import { buildPptx, group, shape, solid } from '../../../../scripts/fixtures/builder.ts';
import { exportPdf, type PdfExportResult } from '../index.ts';

// The export installs the Pro drawing renderers; the editor's paint takes them too.
installDrawing();

const IN = 914_400;
const BEST = { fidelityPolicy: 'best-effort' } as const;
const DIST = 152_400;

const shadowXml = (attrs = '') =>
  `<a:outerShdw blurRad="0" dist="${DIST}" dir="0"${attrs}>` +
  '<a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:outerShdw>';

/** Two overlapping rectangles in a group scaled 2× by its child space. */
const deckOf = (effects: string, xfrm: { rot?: number; flipH?: boolean; flipV?: boolean } = {}) =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: group(
          {
            id: 10,
            x: 2 * IN,
            y: 1.5 * IN,
            cx: 4 * IN,
            cy: 2 * IN,
            chX: 0,
            chY: 0,
            chCx: 2 * IN,
            chCy: IN,
            effects: effects ? `<a:effectLst>${effects}</a:effectLst>` : '',
            ...xfrm,
          },
          shape({ id: 11, x: 0, y: 0, cx: IN, cy: 0.5 * IN, fill: solid('C00000') }) +
            shape({
              id: 12,
              x: 0.75 * IN,
              y: 0.25 * IN,
              cx: IN,
              cy: 0.75 * IN,
              fill: solid('0070C0'),
            })
        ),
      },
    ],
  });

/** The editor's layout of slide 1 and the transform of its group's shadow copy. */
async function editorShadow(bytes: Uint8Array): Promise<{ children: LayoutItem[]; m: Matrix }> {
  const opened = await openPresentation(bytes);
  if (!opened.ok) throw new Error('deck');
  const layout = opened.session.layout(0)!;
  const { document } = new Window();
  const root = paintSlide(document as unknown as Document, layout, {
    images: () => undefined,
    idPrefix: 't',
    hitTargets: false,
  }) as unknown as Element;
  opened.session.dispose();
  const item = layout.items[0];
  if (item?.kind !== 'group') throw new Error('no group');
  const uses = root.querySelectorAll('[data-effect="shadow"] use');
  expect(uses.length).toBe(1);
  const t = /matrix\(([^)]+)\)/.exec(uses[0]!.getAttribute('transform') ?? '')!;
  const m = t[1]!.trim().split(/\s+/).map(Number) as unknown as Matrix;
  return { children: [...item.children], m };
}

/** The forms of the PDF: their path points (slide EMU) and whether each is a transparency group. */
async function forms(bytes: Uint8Array): Promise<{ points: number[][]; group: boolean }[]> {
  const doc = await PDFDocument.load(bytes);
  const out: { points: number[][]; group: boolean }[] = [];
  for (const [, obj] of doc.context.enumerateIndirectObjects()) {
    if (!(obj instanceof PDFRawStream)) continue;
    if (obj.dict.get(PDFName.of('Subtype')) !== PDFName.of('Form')) continue;
    const raw = obj.dict.has(PDFName.of('Filter'))
      ? decodePDFRawStream(obj).decode()
      : obj.contents;
    const tokens = new TextDecoder('latin1').decode(raw).split(/\s+/);
    const points: number[][] = [];
    tokens.forEach((token, i) => {
      if (token !== 'm' && token !== 'l') return;
      points.push(tokens.slice(i - 2, i).map(Number));
    });
    const groupDict = obj.dict.lookupMaybe(PDFName.of('Group'), PDFDict);
    const transparency = groupDict?.get(PDFName.of('S')) === PDFName.of('Transparency');
    out.push({ points, group: transparency });
  }
  return out;
}

const codes = (result: PdfExportResult): string[] => result.diagnostics.map((d) => d.code);

/** The slide-space corners of each child's box under the shadow transform `m`. */
const shadowCorners = (children: readonly LayoutItem[], m: Matrix) =>
  children.flatMap((c) => {
    const { cx, cy } = c.element.xfrm;
    return [
      [0, 0],
      [cx, 0],
      [cx, cy],
      [0, cy],
    ].map(([x, y]) => {
      const p = apply(c.matrix, x!, y!);
      return apply(m, p.x, p.y);
    });
  });

describe('group effects in the PDF export', () => {
  for (const [name, xfrm, attrs] of [
    ['plain', {}, ''],
    ['turned and flipped', { rot: 1_800_000, flipH: true }, ''],
    ['flipped, slide direction', { rot: 2_700_000, flipV: true }, ' rotWithShape="0"'],
  ] as const)
    test(`one shadow of both children, placed as the SVG copy: ${name}`, async () => {
      const bytes = deckOf(shadowXml(attrs), xfrm);
      const { children, m } = await editorShadow(bytes);
      expect(children.length).toBe(2);
      const result = await exportPdf(bytes, BEST);
      expect(codes(result)).toContain('effect.group-silhouette');
      const all = await forms(result.bytes);
      // One layer of one shadow: both children paint in one transparency group, so the overlap
      // does not darken.
      expect(all.filter((f) => f.group).length).toBe(1);
      const points = all.flatMap((f) => f.points);
      for (const corner of shadowCorners(children, m)) {
        const hit = points.some((p) => Math.hypot(p[0]! - corner.x, p[1]! - corner.y) < 200);
        expect(hit).toBe(true);
      }
    });

  test('a group without effects draws no forms', async () => {
    const result = await exportPdf(deckOf(''), BEST);
    expect((await forms(result.bytes)).length).toBe(0);
    expect(codes(result).filter((c) => c.startsWith('effect.'))).toEqual([]);
  });

  test('glow, reflection, and the soft edge of a group draw and are reported', async () => {
    const effects =
      '<a:glow rad="63500"><a:srgbClr val="FF0000"/></a:glow>' +
      '<a:innerShdw blurRad="63500" dist="50800" dir="5400000"><a:srgbClr val="000000"/></a:innerShdw>' +
      '<a:reflection blurRad="0" stA="50000" endPos="55000" dist="0" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>' +
      '<a:softEdge rad="12700"/>';
    const result = await exportPdf(deckOf(effects), BEST);
    const found = codes(result);
    expect(found).toContain('effect.glow');
    expect(found).toContain('effect.group-soft-edge');
    // The soft edge applies to the group, not to each shape; a group's inner shadow draws nothing.
    expect(found).not.toContain('effect.soft-edge');
    expect(found).not.toContain('effect.inner-shadow');
    // The reflection fades through a gray soft mask.
    const text = new TextDecoder('latin1').decode(result.bytes);
    expect(text).toMatch(/\/ShadingType 2\s*\/ColorSpace \/DeviceGray/);
  });

  test("a group soft edge reports that its shapes' own shadows show only over them", async () => {
    const own =
      '<a:effectLst><a:outerShdw blurRad="0" dist="152400" dir="0"><a:srgbClr val="000000"/>' +
      '</a:outerShdw></a:effectLst>';
    const deck = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: group(
            {
              id: 10,
              x: IN,
              y: IN,
              cx: 2 * IN,
              cy: IN,
              effects: '<a:effectLst><a:softEdge rad="63500"/></a:effectLst>',
            },
            shape({ id: 11, x: IN, y: IN, cx: IN, cy: IN, fill: solid('C00000'), effects: own })
          ),
        },
      ],
    });
    expect(codes(await exportPdf(deck, BEST))).toContain('effect.group-soft-edge-effects');
    expect(codes(await exportPdf(deckOf('<a:softEdge rad="63500"/>'), BEST))).not.toContain(
      'effect.group-soft-edge-effects'
    );
  });

  test('a group soft edge masks all bodies together and paints each text once', async () => {
    const labelled = (effects: string) =>
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: group(
              { id: 10, x: 2 * IN, y: 1.5 * IN, cx: 4 * IN, cy: 2 * IN, effects },
              shape({
                id: 11,
                x: 2 * IN,
                y: 1.5 * IN,
                cx: 2.5 * IN,
                cy: IN,
                fill: solid('C00000'),
                text: '<a:p><a:r><a:rPr lang="en-US"/><a:t>Front</a:t></a:r></a:p>',
              }) +
                shape({
                  id: 12,
                  x: 3.5 * IN,
                  y: 2.5 * IN,
                  cx: 2.5 * IN,
                  cy: IN,
                  prst: 'ellipse',
                  fill: solid('0070C0'),
                  text: '<a:p><a:r><a:rPr lang="en-US"/><a:t>Back</a:t></a:r></a:p>',
                })
            ),
          },
        ],
      });
    const soft = await exportPdf(
      labelled('<a:effectLst><a:softEdge rad="63500"/></a:effectLst>'),
      BEST
    );
    const plain = await exportPdf(labelled(''), BEST);
    const streams = await contentText(soft.bytes);
    // One fade per shape, lightest wins: the mask draws them with the Lighten blend mode.
    expect(new TextDecoder('latin1').decode(soft.bytes)).toMatch(/\/BM \/Lighten/);
    // The text paints over the faded bodies, once.
    const textObjects = (all: string[]) =>
      all
        .join(' ')
        .split(/\s+/)
        .filter((t) => t === 'BT').length;
    expect(textObjects(streams)).toBe(textObjects(await contentText(plain.bytes)));
    expect(textObjects(streams)).toBeGreaterThan(0);
  });
});

/** The decoded content of the page and of every form, as text. */
async function contentText(bytes: Uint8Array): Promise<string[]> {
  const doc = await PDFDocument.load(bytes);
  const { Contents } = doc.getPages()[0]!.node.normalizedEntries();
  const pageRefs = new Set((Contents?.asArray() ?? []).map((r) => r.toString()));
  const out: string[] = [];
  for (const [ref, obj] of doc.context.enumerateIndirectObjects()) {
    if (!(obj instanceof PDFRawStream)) continue;
    const form = obj.dict.get(PDFName.of('Subtype')) === PDFName.of('Form');
    if (!form && !pageRefs.has(ref.toString())) continue;
    const raw = obj.dict.has(PDFName.of('Filter'))
      ? decodePDFRawStream(obj).decode()
      : obj.contents;
    out.push(new TextDecoder('latin1').decode(raw));
  }
  return out;
}
