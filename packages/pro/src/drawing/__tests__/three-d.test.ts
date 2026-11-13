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
  type ThreeDFacet,
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

/** Every polygon of a list of facets, each with its facet's color. */
const polygons = (facets: readonly ThreeDFacet[]) =>
  facets.flatMap((f) => f.polygons.map((points) => ({ points, color: f.color })));

const input = (fx: ResolvedEffects) => ({
  effects: fx,
  paths: RECT,
  matrix: [1, 0, 0, 1, IN, IN] as const,
  width: 2 * IN,
  height: IN,
  color: { r: 0x44, g: 0x72, b: 0xc4, a: 1 },
});

describe('3-D polygons', () => {
  test('a circle bevel is twelve strips on each edge, lit from the top', () => {
    const paint = threeDPaint(
      input(effects({ bevelTop: { w: 76_200, h: 76_200, preset: 'circle' } }))
    )!;
    const strips = polygons(paint.over);
    expect(strips).toHaveLength(4 * 12);
    expect(paint.behind).toHaveLength(0);
    expect(paint.front).toEqual([1, 0, 0, 1, 0, 0]);
    expect(paint.showsFace).toBe(true);
    // The outermost strip of each edge: top, right, bottom, left.
    const [top, right, bottom, left] = [0, 1, 2, 3].map((i) => strips[i]!.color.g);
    expect(right).toBeLessThan(left!);
    // The outer bottom strip catches light from below, as in the reference; the fifth strips
    // show the light from the top.
    expect(bottom).toBeGreaterThan(top!);
    expect(strips[4 * 4 + 2]!.color.g).toBeLessThan(strips[4 * 4]!.color.g);
    // The outer right strip is close to the measured 0x17.
    expect(Math.abs(right! - 0x17)).toBeLessThan(12);
  });

  test('a front camera hides the extrusion; a turned one turns the face and shows sides', () => {
    expect(threeDPaint(input(effects({ extrusionH: 228_600 })))?.behind ?? []).toHaveLength(0);
    const turned = threeDPaint(input(effects({ camera: 'isometricTopUp', extrusionH: 228_600 })))!;
    expect(turned.front).not.toEqual([1, 0, 0, 1, 0, 0]);
    // Four sides and the front cap.
    expect(polygons(turned.behind)).toHaveLength(5);
    const oblique = threeDPaint(input(effects({ camera: 'obliqueTopLeft', extrusionH: 228_600 })))!;
    expect(oblique.front).toEqual([1, 0, 0, 1, 0, 0]);
    // The top and left sides face the depth, and the front cap.
    expect(polygons(oblique.behind)).toHaveLength(3);
  });

  test('a contour is a ring outside the outline; a flat material does not shade', () => {
    const contour = threeDPaint(
      input(effects({ contourW: 38_100, contourColor: { r: 192, g: 0, b: 0, a: 1 } }))
    )!;
    // One color: the four sides of the ring are one facet.
    expect(contour.under).toHaveLength(1);
    expect(polygons(contour.under)).toHaveLength(4);
    // Half of the contour shows outside the outline (the top edge lies at y = 0).
    const top = Math.min(...contour.under[0]!.polygons.flat().map((p) => p.y));
    expect(top).toBeCloseTo(-38_100 / 2, 3);
    const flat = threeDPaint(
      input(effects({ material: 'flat', bevelTop: { w: 76_200, h: 76_200, preset: 'angle' } }))
    )!;
    expect(new Set(flat.over.map((f) => f.color.g))).toEqual(new Set([0x72]));
  });

  test('the light shades the face, and the highlight follows the view', () => {
    const face = (fx: ResolvedEffects) => threeDPaint(input(fx))!.faceColor!;
    const bevel = { bevelTop: { w: 76_200, h: 76_200, preset: 'angle' as const } };
    // The face of a lit shape is brighter than its fill: the reference shows 4B7BD0.
    const front = face(effects(bevel));
    expect(Math.abs(front.r - 0x4b)).toBeLessThan(6);
    expect(Math.abs(front.g - 0x7b)).toBeLessThan(6);
    expect(Math.abs(front.b - 0xd0)).toBeLessThan(6);
    // A flat material keeps the fill.
    expect(face(effects({ ...bevel, material: 'flat' }))).toEqual(input(effects({})).color);
    // Two cameras with the same face normal on screen: only the highlight tells them apart
    // (the reference shows 4878CD and 7BABFF).
    const up = face(effects({ camera: 'isometricTopUp', extrusionH: 228_600 }));
    const down = face(effects({ camera: 'isometricTopDown', extrusionH: 228_600 }));
    expect(down.g - up.g).toBeGreaterThan(40);
    // The harsh rig is darker than the three-point rig.
    expect(face(effects({ ...bevel, lightRig: 'harsh' })).g).toBeLessThan(front.g - 20);
  });

  test('a material changes the highlights, not the diffuse light', () => {
    const bevel = { bevelTop: { w: 76_200, h: 76_200, preset: 'circle' as const } };
    // The strips of one side, outer to inner: the side of each strip is its index modulo 4.
    const side = (material: string, k: number) =>
      polygons(threeDPaint(input(effects({ ...bevel, material })))!.over).filter(
        (_, i) => i % 4 === k
      );
    const top = (material: string) => side(material, 0).map((q) => q.color);
    const right = (material: string) => side(material, 1).map((q) => q.color);
    // Matte takes no highlight. The right side, away from the light, shows the same diffuse
    // light in every material.
    const matte = top('matte');
    const warm = top('warmMatte');
    const metal = top('metal');
    expect(Math.max(...warm.map((c) => c.r))).toBeGreaterThan(Math.max(...matte.map((c) => c.r)));
    // Metal: a strong highlight (white in the reference), and a rim at the grazing outer strip
    // on every side, also the dark one (the reference shows 95989E there).
    expect(Math.max(...metal.map((c) => c.r))).toBeGreaterThan(240);
    expect(right('metal')[0]!.r).toBeGreaterThan(right('warmMatte')[0]!.r + 60);
    expect(right('metal')[11]!.g).toBeCloseTo(right('warmMatte')[11]!.g, -1);
    // Plastic has stronger highlights than warm matte.
    const peak = (cs: readonly { r: number }[]) => Math.max(...cs.map((c) => c.r));
    expect(peak(top('plastic'))).toBeGreaterThan(peak(warm) + 10);
  });

  test('the sides take the light of their normals in the shape axes', () => {
    const paint = threeDPaint(
      input(effects({ camera: 'isometricOffAxis2Left', extrusionH: 228_600 }))
    )!;
    // Four sides, far ones first, and the front cap. The top side (normal up) and the right
    // side (normal right) show: the reference shows 335592 on top and 0D131D on the right.
    const greens = polygons(paint.behind).map((q) => q.color.g);
    expect(greens).toHaveLength(5);
    expect(greens.some((g) => Math.abs(g - 0x55) < 12)).toBe(true);
    expect(greens.some((g) => Math.abs(g - 0x13) < 12)).toBe(true);
  });

  test('a turned extrusion has its back at the turned depth axis', () => {
    const turned = threeDPaint(
      input(effects({ cameraRot: { lat: 0, lon: 30, rev: 0 }, extrusionH: IN }))
    )!;
    // A turn of 30 degrees about the vertical axis moves the back by sin 30 of the depth.
    expect(turned.backOffset!.x).toBeCloseTo(Math.sin(Math.PI / 6) * IN, 3);
    expect(turned.backOffset!.y).toBeCloseTo(0, 6);
    expect(threeDPaint(input(effects({ extrusionH: IN })))?.backOffset).toBeUndefined();
    const flatTurn = threeDPaint(input(effects({ cameraRot: { lat: 0, lon: 30, rev: 0 } })))!;
    expect(flatTurn.backOffset).toBeUndefined();
  });

  test('a measured profile rises and falls by its slopes', () => {
    // The soft round bevel rises at the outer edge, then falls toward the face: its inner strips
    // face inward, so the top edge's inner strips are darker than its outer ones.
    const paint = threeDPaint(
      input(effects({ bevelTop: { w: 76_200, h: 76_200, preset: 'softRound' } }))
    )!;
    const strips = polygons(paint.over);
    expect(strips).toHaveLength(4 * 12);
    // Strip k of the top edge is polygon 4k.
    expect(strips[4 * 6]!.color.g).toBeLessThan(strips[4]!.color.g);
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
    // The geometry, 48 bevel strips, four sides, and the front cap, as subpaths: polygons of
    // one color share a path.
    const subpaths = [...svg.querySelectorAll('path')].map(
      (p) => (p.getAttribute('d') ?? '').split('M').length - 1
    );
    expect(subpaths.reduce((a, b) => a + b, 0)).toBe(1 + 48 + 5);
    expect(svg.querySelector('g[data-kind="shape"] g g[transform]')).not.toBeNull();
  });
});

describe('degenerate 3-D shapes', () => {
  const line: GeometryPath[] = [
    {
      commands: [
        { op: 'M', x: IN / 2, y: 0 },
        { op: 'L', x: IN, y: 0 },
        { op: 'L', x: 0, y: 0 },
        { op: 'Z' },
      ],
      fill: 'norm',
      stroke: true,
    },
  ];
  const square: GeometryPath[] = [
    {
      commands: [
        { op: 'M', x: 0, y: 0 },
        { op: 'L', x: IN, y: 0 },
        { op: 'L', x: IN, y: IN },
        { op: 'L', x: 0, y: IN },
        { op: 'Z' },
      ],
      fill: 'norm',
      stroke: true,
    },
  ];
  const finite = (facets: readonly ThreeDFacet[]) =>
    polygons(facets).every(
      (f) =>
        [f.color.r, f.color.g, f.color.b, f.color.a].every(Number.isFinite) &&
        f.points.every((p) => Number.isFinite(p.x) && Number.isFinite(p.y))
    );

  test('a bevel on an outline with no inside has no strips and no NaN', () => {
    const fx = effects({ bevelTop: { w: 76_200, h: 76_200, preset: 'hardEdge' } });
    const paint = threeDPaint({ ...input(fx), paths: line, width: IN, height: 0 });
    expect(paint?.over ?? []).toHaveLength(0);
    expect(finite(paint?.under ?? []) && finite(paint?.behind ?? [])).toBe(true);
  });

  test('a zero-height shape with a bevel paints no NaN', () => {
    installDrawing();
    const deck = shape({
      id: 2,
      x: IN,
      y: IN,
      cx: IN,
      cy: 0,
      prst: 'triangle',
      fill: solid('C00000'),
      effects:
        '<a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>' +
        '<a:sp3d><a:bevelT w="76200" h="76200" prst="hardEdge"/></a:sp3d>',
    });
    const r = readPptx(buildPptx({ slides: [{ shapes: deck, layout: 'blank' }] }));
    if (!r.ok) throw new Error('open failed');
    const layout = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer());
    const svg = paintSlide(doc, layout, { images: () => undefined, idPrefix: 'z' });
    expect(svg.outerHTML).not.toContain('NaN');
  });

  test('a face turned edge-on hides; a half turn mirrors it', () => {
    const edge = threeDPaint({
      ...input(effects({ cameraRot: { lat: 0, lon: 90, rev: 0 }, extrusionH: IN })),
      paths: square,
      width: IN,
      height: IN,
    })!;
    expect(edge.showsFace).toBe(false);
    // The extrusion still shows: its side toward the viewer.
    expect(polygons(edge.behind).length).toBeGreaterThan(0);
    expect(finite(edge.behind)).toBe(true);
    const back = threeDPaint({
      ...input(effects({ cameraRot: { lat: 0, lon: 180, rev: 0 } })),
      paths: square,
      width: IN,
      height: IN,
    })!;
    expect(back.showsFace).toBe(true);
    // A turn about the vertical axis: x mirrors, y stays.
    expect(back.front[0]).toBeCloseTo(-1, 9);
    expect(back.front[3]).toBeCloseTo(1, 9);
  });

  test('an edge-on face paints only its extrusion', () => {
    installDrawing();
    const deck = shape({
      id: 2,
      x: IN,
      y: IN,
      cx: IN,
      cy: IN,
      fill: solid('4472C4'),
      effects:
        '<a:effectLst><a:outerShdw dist="76200" dir="2700000"><a:prstClr val="black"/></a:outerShdw></a:effectLst>' +
        '<a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="5400000" rev="0"/></a:camera>' +
        '<a:lightRig rig="threePt" dir="t"/></a:scene3d><a:sp3d extrusionH="228600"/>',
    });
    const r = readPptx(buildPptx({ slides: [{ shapes: deck, layout: 'blank' }] }));
    if (!r.ok) throw new Error('open failed');
    const layout = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer());
    const svg = paintSlide(doc, layout, { images: () => undefined, idPrefix: 'e' });
    const item = svg.querySelector('g[data-kind="shape"]')!;
    // No face, no shadow copy: the extrusion paths alone.
    expect(item.querySelector('use')).toBeNull();
    expect(item.querySelector('g[transform]')).toBeNull();
    expect(item.querySelectorAll('path').length).toBeGreaterThan(0);
  });

  test('the blur region holds the turned face', () => {
    installDrawing();
    // A camera turned 45 degrees about the view axis: the 2 by 0.5 in face reaches 0.88 in
    // above and below its center, past the unturned box.
    const deck = shape({
      id: 2,
      x: IN,
      y: 2 * IN,
      cx: 2 * IN,
      cy: IN / 2,
      fill: solid('4472C4'),
      effects:
        '<a:effectLst><a:blur rad="12700" grow="0"/></a:effectLst>' +
        '<a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="2700000"/></a:camera>' +
        '<a:lightRig rig="threePt" dir="t"/></a:scene3d>',
    });
    const r = readPptx(buildPptx({ slides: [{ shapes: deck, layout: 'blank' }] }));
    if (!r.ok) throw new Error('open failed');
    const layout = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer());
    const svg = paintSlide(doc, layout, { images: () => undefined, idPrefix: 'b' });
    const filter = svg.querySelector('filter[id*="blur"]')!;
    const y = Number(filter.getAttribute('y'));
    const h = Number(filter.getAttribute('height'));
    const center = 2.25 * IN;
    const reach = (Math.SQRT1_2 * (2 * IN + IN / 2)) / 2;
    expect(y).toBeLessThanOrEqual(center - reach);
    expect(y + h).toBeGreaterThanOrEqual(center + reach);
  });

  test('a long outline is thinned, and neighboring polygons of one color share a path', () => {
    const N = 4000;
    const commands: GeometryPath['commands'][number][] = [{ op: 'M', x: 2 * IN, y: IN }];
    for (let i = 1; i < N; i += 1) {
      const a = (i / N) * 2 * Math.PI;
      commands.push({ op: 'L', x: IN + Math.cos(a) * IN, y: IN + Math.sin(a) * IN });
    }
    commands.push({ op: 'Z' });
    const fx = effects({
      bevelTop: { w: 76_200, h: 76_200, preset: 'circle' },
      contourW: 12_700,
      cameraRot: { lat: 20, lon: 30, rev: 0 },
      extrusionH: IN,
    });
    const paint = threeDPaint({
      ...input(fx),
      paths: [{ commands, fill: 'norm', stroke: true }],
      width: 2 * IN,
      height: 2 * IN,
    })!;
    const facets = paint.over.length + paint.under.length + paint.behind.length;
    const all = polygons(paint.over).length + polygons(paint.under).length;
    // The 3-D work follows every fourth point: 1,000 points, 12 strips and a contour side each.
    expect(all).toBe(1000 * 12 + 1000);
    // Neighboring polygons of one color share a path: about 5,200 paths for 13,000 polygons.
    expect(facets).toBeLessThan(6000);
  });
});

describe('pattern fills under a turned camera', () => {
  test('the tiles stay square to the slide: the transforms above the fill cancel', () => {
    installDrawing();
    const deck = shape({
      id: 2,
      x: IN,
      y: 2 * IN,
      cx: 2 * IN,
      cy: IN / 2,
      fill: '<a:pattFill prst="horz"><a:fgClr><a:srgbClr val="C00000"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>',
      effects:
        '<a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="2700000"/></a:camera>' +
        '<a:lightRig rig="threePt" dir="t"/></a:scene3d>',
    });
    const r = readPptx(buildPptx({ slides: [{ shapes: deck, layout: 'blank' }] }));
    if (!r.ok) throw new Error('open failed');
    const layout = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer());
    const svg = paintSlide(doc, layout, { images: () => undefined, idPrefix: 'p' });
    const pattern = [...svg.querySelectorAll('pattern')].find((p) => p.id.includes('patt'))!;
    const filled = [...svg.querySelectorAll('path')].find(
      (p) => p.getAttribute('fill') === `url(#${pattern.id})`
    )!;
    let m = [1, 0, 0, 1, 0, 0];
    let turned = false;
    for (let el: Element | null = filled; el && el !== svg; el = el.parentElement) {
      const t = el.getAttribute('transform');
      if (!t) continue;
      const [a, b, c, d, e, f] = t
        .replace(/^matrix\(|\)$/g, '')
        .split(' ')
        .map(Number);
      if (Math.abs(b) > 0.01) turned = true;
      // The ancestor's transform applies after the ones below it.
      m = [
        a * m[0] + c * m[1],
        b * m[0] + d * m[1],
        a * m[2] + c * m[3],
        b * m[2] + d * m[3],
        a * m[4] + c * m[5] + e,
        b * m[4] + d * m[5] + f,
      ];
    }
    expect(turned).toBe(true);
    m.slice(0, 4).forEach((v, i) => expect(v).toBeCloseTo([1, 0, 0, 1][i], 6));
    expect(Math.abs(m[4]) + Math.abs(m[5])).toBeLessThan(1);
  });
});
