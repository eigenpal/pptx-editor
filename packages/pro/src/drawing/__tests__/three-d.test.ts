/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Bevels, contours, extrusion, and cameras as polygons, and their SVG paint with and without
// the Pro drawing renderers.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, shape, solid } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '@pptx-editor.dev/core/store';
import { resolveSlide, type ResolvedEffects } from '@pptx-editor.dev/core/cascade';
import {
  createFixedMeasurer,
  installDrawingRenderers,
  layoutSlide,
  type GeometryPath,
} from '@pptx-editor.dev/core/layout';
import { paintSlide } from '@pptx-editor.dev/core/output';
import { installDrawing } from '../index.ts';
import { threeDPaint } from '../three-d.ts';

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const IN = 914_400;
const RECT: GeometryPath[] = [
  {
    commands: [
      { op: 'M', x: 0, y: 0 },
      { op: 'L', x: 2 * IN, y: 0 },
      { op: 'L', x: 2 * IN, y: IN },
      { op: 'L', x: 0, y: IN },
      { op: 'Z' },
    ],
    fill: 'norm',
    stroke: true,
  },
];

const effects = (threeD: Partial<NonNullable<ResolvedEffects['threeD']>>): ResolvedEffects => ({
  unsupported: [],
  threeD: {
    camera: 'orthographicFront',
    lightRig: 'threePt',
    lightDir: 't',
    extrusionH: 0,
    contourW: 0,
    material: 'warmMatte',
    ...threeD,
  },
});

const input = (fx: ResolvedEffects) => ({
  effects: fx,
  paths: RECT,
  matrix: [1, 0, 0, 1, IN, IN] as const,
  width: 2 * IN,
  height: IN,
  color: { r: 0x44, g: 0x72, b: 0xc4, a: 1 },
});

describe('3-D polygons', () => {
  test('a circle bevel is eight strips on each edge, lit from the top', () => {
    const paint = threeDPaint(
      input(effects({ bevelTop: { w: 76_200, h: 76_200, preset: 'circle' } }))
    )!;
    expect(paint.over).toHaveLength(4 * 8);
    expect(paint.behind).toHaveLength(0);
    expect(paint.front).toEqual([1, 0, 0, 1, 0, 0]);
    // The outermost strip of each edge: top, right, bottom, left.
    const [top, right, bottom, left] = [0, 1, 2, 3].map((i) => paint.over[i]!.color.g);
    expect(right).toBeLessThan(left!);
    expect(bottom).toBeLessThan(top!);
    // The outer right strip is close to the measured 0x17.
    expect(Math.abs(right! - 0x17)).toBeLessThan(12);
  });

  test('a front camera hides the extrusion; a turned one turns the face and shows sides', () => {
    expect(threeDPaint(input(effects({ extrusionH: 228_600 })))?.behind ?? []).toHaveLength(0);
    const turned = threeDPaint(input(effects({ camera: 'isometricTopUp', extrusionH: 228_600 })))!;
    expect(turned.front).not.toEqual([1, 0, 0, 1, 0, 0]);
    // Four sides and the front cap.
    expect(turned.behind).toHaveLength(5);
    const oblique = threeDPaint(input(effects({ camera: 'obliqueTopLeft', extrusionH: 228_600 })))!;
    expect(oblique.front).toEqual([1, 0, 0, 1, 0, 0]);
    // The top and left sides face the depth, and the front cap.
    expect(oblique.behind).toHaveLength(3);
  });

  test('a contour is a ring outside the outline; a flat material does not shade', () => {
    const contour = threeDPaint(
      input(effects({ contourW: 38_100, contourColor: { r: 192, g: 0, b: 0, a: 1 } }))
    )!;
    expect(contour.under).toHaveLength(4);
    expect(contour.under[0]!.points.some((p) => p.y < IN)).toBe(true);
    const flat = threeDPaint(
      input(effects({ material: 'flat', bevelTop: { w: 76_200, h: 76_200, preset: 'angle' } }))
    )!;
    expect(new Set(flat.over.map((f) => f.color.g))).toEqual(new Set([0x72]));
  });

  test('an outline past the point limit draws flat', () => {
    const many: GeometryPath[] = [
      {
        commands: [
          { op: 'M', x: 0, y: 0 },
          ...Array.from({ length: 5000 }, (_, i) => ({
            op: 'L' as const,
            x: i * 10,
            y: (i % 2) * 1000,
          })),
          { op: 'Z' },
        ],
        fill: 'norm',
        stroke: true,
      },
    ];
    const fx = effects({ bevelTop: { w: 76_200, h: 76_200, preset: 'circle' } });
    expect(threeDPaint({ ...input(fx), paths: many })).toBeNull();
  });
});

describe('3-D paint', () => {
  const deck = shape({
    id: 2,
    x: IN,
    y: IN,
    cx: 2 * IN,
    cy: IN,
    fill: solid('4472C4'),
    effects:
      '<a:scene3d><a:camera prst="isometricTopUp"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>' +
      '<a:sp3d extrusionH="228600"><a:bevelT w="76200" h="76200"/></a:sp3d>',
  });
  const paint = (): SVGSVGElement => {
    const r = readPptx(buildPptx({ slides: [{ shapes: deck, layout: 'blank' }] }));
    if (!r.ok) throw new Error('open failed');
    const layout = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer());
    return paintSlide(doc, layout, { images: () => undefined, idPrefix: 't' });
  };

  test('paints only with the drawing renderers', () => {
    installDrawingRenderers({ effects: undefined });
    expect(paint().querySelectorAll('path').length).toBe(1);
    installDrawing();
    const svg = paint();
    // The geometry, 32 bevel strips, four sides, and the front cap.
    expect(svg.querySelectorAll('path').length).toBe(1 + 32 + 5);
    expect(svg.querySelector('g[data-kind="shape"] g g[transform]')).not.toBeNull();
  });
});
