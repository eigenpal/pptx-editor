// Fill fidelity: gradient geometry, the two-stop color rule,
// the gradient tile rectangle, the fill overlay, tiled picture recoloring, and the picture mask.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '../../store/index.ts';
import {
  resolveSlide,
  type GradientStop,
  type ResolvedEffects,
  type ResolvedFill,
} from '../../cascade/index.ts';
import { createFixedMeasurer, layoutSlide } from '../../layout/index.ts';
import { MAX_OVERLAY_PATHS as REPORTED_OVERLAY_PATHS } from '../../layout/approximations.ts';
import { paintSlide } from '../index.ts';
import { MAX_OVERLAY_PATHS } from '../paint.ts';
import { MAX_TILE_IMAGES, tilePlacements } from '../paint-fill.ts';
import {
  colorAt,
  linearEnds,
  paintStops,
  tileBox,
  twoStopColor,
  bends,
  TWO_STOP_STEPS,
  MAX_SHAPE_BANDS,
  MIN_SHAPE_BANDS,
  rectGeometry,
  shapeBands,
  ShapeBandBudget,
  shapeGradientBands,
} from '../paint-gradient.ts';

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

function paint(shapes: string, rels: [string, string, string][] = []): SVGSVGElement {
  const r = readPptx(buildPptx({ slides: [{ shapes, layout: 'blank', rels }] }));
  if (!r.ok) throw new Error('open failed');
  const layout = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer());
  return paintSlide(doc, layout, {
    images: () => 'blob:test',
    imageSize: () => ({ cx: 1_219_200, cy: 914_400 }),
    idPrefix: 't',
  });
}

function resolvedFill(fill: string): ResolvedFill {
  const r = readPptx(
    buildPptx({
      slides: [{ shapes: shape({ id: 2, x: 0, y: 0, cx: 1, cy: 1, fill }), layout: 'blank' }],
    })
  );
  if (!r.ok) throw new Error('open failed');
  const el = resolveSlide(r.pkg, 0)!.elements.find((e) => e.kind === 'shape');
  if (!el || el.kind !== 'shape') throw new Error('no shape');
  return el.fill;
}

const black = { r: 0, g: 0, b: 0, a: 1 };
const white = { r: 255, g: 255, b: 255, a: 1 };
const twoStops: GradientStop[] = [
  { pos: 0, color: black },
  { pos: 1, color: white },
];
describe('linear gradient ends (design D1)', () => {
  const box = { x: 0, y: 0, w: 16, h: 9 };

  test('unscaled: the ends are the projections of the farthest corners', () => {
    const e = linearEnds(0, false, box);
    expect(e).toEqual({ x1: 0, y1: 4.5, x2: 16, y2: 4.5 });
  });

  test('scaled: a 315° gradient keeps the top-left to bottom-right diagonal on one isoline', () => {
    const e = linearEnds(315, true, box);
    const dx = e.x2 - e.x1;
    const dy = e.y2 - e.y1;
    // The diagonal (16, 9) is perpendicular to the gradient direction.
    expect(16 * dx + 9 * dy).toBeCloseTo(0, 6);
    // The bottom-left corner is the start, the top-right corner the end.
    const t = (x: number, y: number): number =>
      ((x - e.x1) * dx + (y - e.y1) * dy) / (dx * dx + dy * dy);
    expect(t(0, 9)).toBeCloseTo(0, 6);
    expect(t(16, 0)).toBeCloseTo(1, 6);
    expect(t(0, 0)).toBeCloseTo(0.5, 6);
  });

  test('scaled and unscaled agree on axis-aligned angles', () => {
    for (const a of [0, 90, 180, 270]) {
      const s = linearEnds(a, true, box);
      const u = linearEnds(a, false, box);
      expect(s.x1).toBeCloseTo(u.x1, 6);
      expect(s.y2).toBeCloseTo(u.y2, 6);
    }
  });
});

describe('gradient colors (design D2)', () => {
  const blue = { r: 0x2e, g: 0x4b, b: 0xe6, a: 1 };
  const orange = { r: 0xf2, g: 0xa2, b: 0x3a, a: 1 };

  test('two stops bend each channel toward its larger end value', () => {
    // Measured in the reference PDF (fixtures/fill-fidelity.pdf, slide 1): (193, 143, 183) at u = 0.51.
    const mid = twoStopColor(blue, orange, 0.5);
    expect(Math.round(mid.r)).toBe(193);
    expect(Math.round(mid.g)).toBe(140);
    expect(Math.round(mid.b)).toBe(187);
    expect(twoStopColor(blue, orange, 0)).toEqual(blue);
    expect(twoStopColor(blue, orange, 1)).toEqual(orange);
  });

  test('only two stops at 0 and 100 % bend; other stop lists stay linear', () => {
    const at = (a: number, b: number) => [
      { pos: a, color: blue },
      { pos: b, color: orange },
    ];
    expect(bends(at(0, 1))).toBe(true);
    expect(bends(at(0, 0.88))).toBe(false);
    expect(bends(at(0.2, 1))).toBe(false);
    expect(colorAt(at(0, 0.5), 0.25)).toEqual(
      colorAt([...at(0, 0.5), { pos: 1, color: orange }], 0.25)
    );
  });

  test('alpha stays linear', () => {
    const faded = { ...white, a: 0.2 };
    expect(twoStopColor(white, faded, 0.5).a).toBeCloseTo(0.6, 9);
  });

  test('white to a color is the color ramp at u²', () => {
    const red = { r: 192, g: 0, b: 0, a: 1 };
    const c = twoStopColor(white, red, 0.5);
    expect(c.g).toBeCloseTo(255 * (1 - 0.25), 9);
    expect(
      colorAt(
        [
          { pos: 0, color: white },
          { pos: 1, color: red },
        ],
        0.5
      )
    ).toEqual(c);
  });

  test('three or more stops interpolate linearly in sRGB', () => {
    const three: GradientStop[] = [
      { pos: 0, color: black },
      { pos: 0.5, color: { r: 100, g: 100, b: 100, a: 1 } },
      { pos: 1, color: white },
    ];
    expect(colorAt(three, 0.25)).toEqual({ r: 50, g: 50, b: 50, a: 1 });
    expect(paintStops(three)).toEqual(three);
  });

  test('a two-stop gradient is resampled at a fixed count; hard stops are not', () => {
    const out = paintStops([
      { pos: 0, color: black },
      { pos: 1, color: white },
    ]);
    expect(out).toHaveLength(TWO_STOP_STEPS + 1);
    expect(out[0].pos).toBe(0);
    expect(out[out.length - 1].pos).toBeCloseTo(1, 12);
    const inner: GradientStop[] = [
      { pos: 0.2, color: black },
      { pos: 0.6, color: white },
    ];
    expect(paintStops(inner)).toEqual(inner);
    const hard: GradientStop[] = [
      { pos: 0.5, color: black },
      { pos: 0.5, color: white },
    ];
    expect(paintStops(hard)).toEqual(hard);
    expect(colorAt(twoStops, -1)).toEqual(black);
    expect(colorAt(twoStops, 2)).toEqual(white);
  });

  test('shape gradient bands follow the steepest color change, within fixed bounds', () => {
    const gray = (v: number): GradientStop['color'] => ({ r: v, g: v, b: v, a: 1 });
    expect(
      shapeBands([
        { pos: 0, color: gray(100) },
        { pos: 1, color: gray(110) },
      ])
    ).toBe(MIN_SHAPE_BANDS);
    expect(shapeBands(twoStops)).toBe(MAX_SHAPE_BANDS);
    const three: GradientStop[] = [
      { pos: 0, color: gray(0) },
      { pos: 0.5, color: gray(60) },
      { pos: 1, color: gray(100) },
    ];
    expect(shapeBands(three)).toBe(60);
  });

  test('shape gradient bands run outside in about the focus center (shared with pptx-to-pdf)', () => {
    const gray = (v: number): GradientStop['color'] => ({ r: v, g: v, b: v, a: 1 });
    const stops: GradientStop[] = [
      { pos: 0, color: gray(100) },
      { pos: 1, color: gray(110) },
    ];
    const fill = {
      type: 'path',
      path: 'shape',
      stops,
      fillToRect: { l: 0.25, t: 0.5, r: 0.75, b: 0.5 },
    } as unknown as Extract<ResolvedFill, { type: 'path' }>;
    const { cx, cy, bands } = shapeGradientBands(fill, stops, 100, 40);
    expect([cx, cy]).toEqual([25, 20]);
    expect(bands).toHaveLength(MIN_SHAPE_BANDS);
    expect(bands[0].scale).toBe(1);
    expect(bands[0].color).toEqual(colorAt(stops, 1 - 0.5 / MIN_SHAPE_BANDS));
    expect(bands[bands.length - 1].scale).toBeCloseTo(1 / MIN_SHAPE_BANDS);
  });

  test('the shape band budget bounds paths, commands, and the work of a slide', () => {
    const budget = new ShapeBandBudget();
    expect(budget.allow(17, 10)).toBe(false);
    expect(budget.allow(1, 2001)).toBe(false);
    // 512,000 commands × bands per slide: two shapes of 2,000 commands fit, a third does not.
    expect(budget.allow(1, 2000)).toBe(true);
    expect(budget.allow(1, 2000)).toBe(true);
    expect(budget.allow(1, 2000)).toBe(false);
    expect(budget.allow(1, 1)).toBe(false);
  });

  test('a rectangle gradient reaches the tile rectangle grown to hold the focus', () => {
    const fill = {
      type: 'path',
      path: 'rect',
      stops: [],
      fillToRect: { l: 0.5, t: 0.5, r: 0.5, b: 0.5 },
      tileRect: { l: 0.1, t: 0, r: 0.1, b: 0 },
    } as unknown as Extract<ResolvedFill, { type: 'path' }>;
    expect(rectGeometry(fill, 100, 40)).toEqual({
      l: 50,
      t: 20,
      r: 50,
      b: 20,
      x0: 10,
      y0: 0,
      x1: 90,
      y1: 40,
    });
  });

  test('a tile rectangle with negative insets extends the box; an empty one falls back', () => {
    expect(tileBox({ l: 0, t: 0, r: -1, b: -1 }, 10, 20)).toEqual({ x: 0, y: 0, w: 20, h: 40 });
    expect(tileBox({ l: -1, t: 0, r: 0, b: -1 }, 10, 20)).toEqual({ x: -10, y: 0, w: 20, h: 40 });
    expect(tileBox({ l: 0.6, t: 0, r: 0.6, b: 0 }, 10, 20)).toEqual({ x: 0, y: 0, w: 10, h: 20 });
    expect(tileBox(undefined, 10, 20)).toEqual({ x: 0, y: 0, w: 10, h: 20 });
  });
});

describe('cascade', () => {
  const overlayOf = (effects: string, style = ''): ResolvedEffects | undefined => {
    const r = readPptx(
      buildPptx({
        slides: [
          {
            shapes: shape({ id: 2, x: 0, y: 0, cx: 1, cy: 1, fill: '<a:noFill/>', effects, style }),
            layout: 'blank',
          },
        ],
      })
    );
    if (!r.ok) throw new Error('open failed');
    const el = resolveSlide(r.pkg, 0)!.elements.find((e) => e.kind === 'shape');
    return el && el.kind === 'shape' ? el.effects : undefined;
  };

  test('a fill overlay with an unknown blend or a group fill is reported, not drawn', () => {
    const bad = overlayOf(
      '<a:effectLst><a:fillOverlay blend="burn"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:fillOverlay></a:effectLst>'
    );
    expect(bad?.fillOverlay).toBeUndefined();
    expect(bad?.unsupported).toEqual(['fillOverlay']);
    const group = overlayOf(
      '<a:effectLst><a:fillOverlay blend="mult"><a:grpFill/></a:fillOverlay></a:effectLst>'
    );
    expect(group?.fillOverlay).toBeUndefined();
    const ok = overlayOf(
      '<a:effectLst><a:fillOverlay blend="screen"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:fillOverlay></a:effectLst>'
    );
    expect(ok?.fillOverlay).toEqual({
      blend: 'screen',
      fill: { type: 'solid', color: { r: 0, g: 255, b: 0, a: 1 } },
    });
  });

  test('a gradient keeps at most 64 stops, the first and the last among them', () => {
    const gs = Array.from(
      { length: 1000 },
      (_, i) =>
        `<a:gs pos="${Math.round((i / 999) * 100000)}"><a:srgbClr val="${i === 999 ? 'FFFFFF' : '000000'}"/></a:gs>`
    ).join('');
    const fill = resolvedFill(`<a:gradFill><a:gsLst>${gs}</a:gsLst><a:lin ang="0"/></a:gradFill>`);
    if (fill.type !== 'linear') throw new Error('not linear');
    expect(fill.stops).toHaveLength(64);
    expect(fill.stops[0].pos).toBe(0);
    expect(fill.stops[63]).toEqual({ pos: 1, color: { r: 255, g: 255, b: 255, a: 1 } });
  });

  test('a gradient keeps its tile rectangle', () => {
    const fill = resolvedFill(
      '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="000000"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst>' +
        '<a:path path="circle"><a:fillToRect r="100000" b="100000"/></a:path><a:tileRect r="-100000" b="-100000"/></a:gradFill>'
    );
    expect(fill).toMatchObject({ type: 'path', tileRect: { l: 0, t: 0, r: -1, b: -1 } });
  });
});

const STOPS =
  '<a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="C00000"/></a:gs></a:gsLst>';

describe('paint', () => {
  test('a two-stop circle gradient carries the resampled stops, linear in the distance', () => {
    const root = paint(
      shape({
        id: 2,
        x: 0,
        y: 0,
        cx: 2_000_000,
        cy: 1_000_000,
        fill: `<a:gradFill>${STOPS}<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>`,
      })
    );
    const radial = root.querySelector('radialGradient')!;
    expect(radial.querySelectorAll('stop').length).toBe(TWO_STOP_STEPS + 1);
    expect(Number(radial.getAttribute('r'))).toBeCloseTo(Math.hypot(1_000_000, 500_000), 0);
    expect(radial.getAttribute('fx')).toBe(radial.getAttribute('cx'));
  });

  test('a corner gradient spans the tile rectangle with its focal point at the corner', () => {
    const root = paint(
      shape({
        id: 2,
        x: 0,
        y: 0,
        cx: 2_000_000,
        cy: 1_000_000,
        fill: `<a:gradFill>${STOPS}<a:path path="circle"><a:fillToRect r="100000" b="100000"/></a:path><a:tileRect r="-100000" b="-100000"/></a:gradFill>`,
      })
    );
    const radial = root.querySelector('radialGradient')!;
    // The circle through the corners of the doubled box; the focal point stays at the corner.
    expect(radial.getAttribute('cx')).toBe('2000000');
    expect(radial.getAttribute('cy')).toBe('1000000');
    // The focus is on the circle: the circle grows to 1.005 times the focus distance.
    expect(Number(radial.getAttribute('r'))).toBeCloseTo(
      1.005 * Math.hypot(2_000_000, 1_000_000),
      -1
    );
    expect(radial.getAttribute('fx')).toBe('0');
    expect(radial.getAttribute('fy')).toBe('0');
  });

  test('a focus outside the circle stays, and the circle grows to contain it', () => {
    const root = paint(
      shape({
        id: 2,
        x: 0,
        y: 0,
        cx: 2_000_000,
        cy: 1_000_000,
        fill: `<a:gradFill>${STOPS}<a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill>`,
      })
    );
    const radial = root.querySelector('radialGradient')!;
    expect(radial.getAttribute('fx')).toBe('1000000');
    expect(radial.getAttribute('fy')).toBe('-800000');
    expect(Number(radial.getAttribute('r'))).toBeCloseTo(1.005 * 1_300_000, -1);
  });

  test('a shape path gradient draws bands of the outline, not trapezoids', () => {
    const root = paint(
      shape({
        id: 2,
        x: 0,
        y: 0,
        cx: 2_000_000,
        cy: 1_000_000,
        prst: 'ellipse',
        fill: `<a:gradFill>${STOPS}<a:path path="shape"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>`,
      })
    );
    const pattern = [...root.querySelectorAll('pattern')].find((p) => p.id.includes('pathshape'))!;
    expect(pattern.querySelectorAll('use').length).toBe(MAX_SHAPE_BANDS);
    expect(root.querySelectorAll('polygon').length).toBe(0);
  });

  test('a fill overlay multiplies over the fill in an isolated group, under the outline', () => {
    const root = paint(
      shape({
        id: 2,
        x: 0,
        y: 0,
        cx: 2_000_000,
        cy: 1_000_000,
        fill: '<a:solidFill><a:srgbClr val="808080"/></a:solidFill>',
        line: '<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>',
        effects: `<a:effectLst><a:fillOverlay blend="mult"><a:gradFill>${STOPS}<a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillOverlay></a:effectLst>`,
      })
    );
    const group = [...root.querySelectorAll('g')].find(
      (g) => g.getAttribute('style') === 'isolation: isolate'
    )!;
    const [base, overlay] = [...group.querySelectorAll('path')];
    expect(base.getAttribute('fill')).toBe('#808080');
    expect(overlay.getAttribute('style')).toBe('mix-blend-mode: multiply');
    expect(overlay.getAttribute('fill')).toMatch(/^url\(#t-lin/);
    // The outline draws after the group, with no fill of its own.
    const outline = group.nextElementSibling!;
    expect(outline.getAttribute('fill')).toBe('none');
    expect(outline.getAttribute('stroke')).toBe('#000000');
  });

  test('a fill overlay on a picture blends with the image inside the picture', () => {
    const root = paint(
      '<p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
        '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
        '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>' +
        '<a:effectLst><a:fillOverlay blend="screen"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:fillOverlay></a:effectLst></p:spPr></p:pic>',
      [['rIdImg', IMAGE_REL, '../media/image1.png']]
    );
    const group = [...root.querySelectorAll('g')].find(
      (g) => g.getAttribute('style') === 'isolation: isolate'
    )!;
    expect(group.firstElementChild!.tagName).toBe('image');
    expect(group.querySelector('path')!.getAttribute('style')).toBe('mix-blend-mode: screen');
    // The ellipse needs a clip: the isolated group sits inside the mask group.
    expect(group.parentElement!.getAttribute('mask')).toMatch(/^url\(#t-clip/);
  });

  test('a fill overlay with a compound outline fills once and keeps the bands', () => {
    const root = paint(
      shape({
        id: 2,
        x: 0,
        y: 0,
        cx: 2_000_000,
        cy: 1_000_000,
        fill: '<a:solidFill><a:srgbClr val="808080"/></a:solidFill>',
        line: '<a:ln w="38100" cmpd="dbl"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>',
        effects:
          '<a:effectLst><a:fillOverlay blend="darken"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:fillOverlay></a:effectLst>',
      })
    );
    const filled = [...root.querySelectorAll('path')].filter(
      (p) => p.getAttribute('fill') === '#808080'
    );
    expect(filled).toHaveLength(1);
    expect(root.querySelector('[data-compound="dbl"]')).not.toBeNull();
    // Inner bands keep to the inside of the path with a mask, not a clip path.
    expect(root.querySelectorAll('clipPath').length).toBe(0);
  });

  test('a hostile tile scale stays within the largest slide side', () => {
    const root = paint(
      shape({
        id: 2,
        x: 0,
        y: 0,
        cx: 914_400,
        cy: 457_200,
        fill: '<a:blipFill><a:blip r:embed="rIdImg"/><a:tile sx="2000000000" sy="1" algn="tl"/></a:blipFill>',
      }),
      [['rIdImg', IMAGE_REL, '../media/image1.png']]
    );
    const tile = root.querySelector('image')!;
    expect(Number(tile.getAttribute('width'))).toBe(51_206_400);
    expect(Number(tile.getAttribute('height'))).toBe(9525);
  });

  test('rectangle and shape gradients carry stop alpha in a mask, not twice', () => {
    const faded = `<a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:gs><a:gs pos="100000"><a:srgbClr val="C00000"><a:alpha val="20000"/></a:srgbClr></a:gs></a:gsLst>`;
    for (const [kind, prst] of [
      ['rect', 'rect'],
      ['shape', 'ellipse'],
    ]) {
      const root = paint(
        shape({
          id: 2,
          x: 0,
          y: 0,
          cx: 2_000_000,
          cy: 1_000_000,
          prst,
          fill: `<a:gradFill>${faded}<a:path path="${kind}"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>`,
        })
      );
      const mask = [...root.querySelectorAll('mask')].find((m) => m.id.includes('alpha'))!;
      expect(mask).toBeDefined();
      // The painted construction is opaque; the alpha lives only in the mask.
      const pattern = [...root.querySelectorAll('pattern')].find((p) => p.id.includes('path'))!;
      expect(pattern.querySelector('g')!.getAttribute('mask')).toBe(`url(#${mask.id})`);
      expect(pattern.querySelectorAll('[stop-opacity]').length).toBe(0);
    }
  });

  test('a rectangle gradient reaches the edge of its tile rectangle', () => {
    const root = paint(
      shape({
        id: 2,
        x: 0,
        y: 0,
        cx: 2_000_000,
        cy: 1_000_000,
        fill: `<a:gradFill>${STOPS}<a:path path="rect"><a:fillToRect r="100000" b="100000"/></a:path><a:tileRect r="-100000" b="-100000"/></a:gradFill>`,
      })
    );
    const points = [...root.querySelectorAll('polygon')].map((p) => p.getAttribute('points'));
    expect(points.some((p) => p!.includes('4000000,2000000'))).toBe(true);
  });

  test('a shape with many paths falls back to the rectangle construction', () => {
    const paths = Array.from(
      { length: 20 },
      (_, i) =>
        `<a:path w="100" h="100"><a:moveTo><a:pt x="${i}" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:lnTo><a:pt x="0" y="100"/></a:lnTo><a:close/></a:path>`
    ).join('');
    const root = paint(
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Many"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm>' +
        `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>${paths}</a:pathLst></a:custGeom>` +
        `<a:gradFill>${STOPS}<a:path path="shape"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></p:spPr></p:sp>`
    );
    expect(root.querySelectorAll('use').length).toBe(0);
  });

  test('prototype names from the file fall back: tile algn and preset dash', () => {
    for (const name of ['toString', 'constructor', '__proto__', 'hasOwnProperty']) {
      const root = paint(
        shape({
          id: 2,
          x: 0,
          y: 0,
          cx: 2_000_000,
          cy: 1_000_000,
          fill: `<a:blipFill><a:blip r:embed="rIdImg"/><a:tile sx="50000" sy="50000" algn="${name}"/></a:blipFill>`,
          line: `<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="${name}"/></a:ln>`,
        }),
        [['rIdImg', IMAGE_REL, '../media/image1.png']]
      );
      const tile = root.querySelector('image')!;
      expect([tile.getAttribute('x'), tile.getAttribute('y')]).toEqual(['0', '0']);
      expect(root.querySelector('path[stroke]')!.hasAttribute('stroke-dasharray')).toBe(false);
    }
  });

  test('one shape shares one paint server per fill mode across its paths', () => {
    const paths = Array.from(
      { length: 3 },
      (_, i) =>
        `<a:path w="100" h="100"><a:moveTo><a:pt x="${i * 30}" y="0"/></a:moveTo><a:lnTo><a:pt x="${i * 30 + 20}" y="100"/></a:lnTo><a:lnTo><a:pt x="${i * 30}" y="100"/></a:lnTo><a:close/></a:path>`
    ).join('');
    const root = paint(
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Three"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm>' +
        `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>${paths}</a:pathLst></a:custGeom>` +
        `<a:gradFill>${STOPS}<a:lin ang="0" scaled="0"/></a:gradFill>` +
        '<a:effectLst><a:fillOverlay blend="mult"><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="000000"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillOverlay></a:effectLst></p:spPr></p:sp>'
    );
    expect(root.querySelectorAll('linearGradient').length).toBe(2);
    expect(root.querySelectorAll('[style="isolation: isolate"]').length).toBe(3);
  });

  test('an overlay on a shape with many paths uses no outline bands', () => {
    const paths = Array.from(
      { length: 20 },
      (_, i) =>
        `<a:path w="100" h="100"><a:moveTo><a:pt x="${i}" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:lnTo><a:pt x="0" y="100"/></a:lnTo><a:close/></a:path>`
    ).join('');
    const root = paint(
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Many"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm>' +
        `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>${paths}</a:pathLst></a:custGeom>` +
        '<a:solidFill><a:srgbClr val="808080"/></a:solidFill>' +
        `<a:effectLst><a:fillOverlay blend="mult"><a:gradFill>${STOPS}<a:path path="shape"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:fillOverlay></a:effectLst></p:spPr></p:sp>`
    );
    expect(root.querySelectorAll('use').length).toBe(0);
  });

  test('every shape band keeps a distinct scale', () => {
    const root = paint(
      shape({
        id: 2,
        x: 0,
        y: 0,
        cx: 2_000_000,
        cy: 1_000_000,
        prst: 'ellipse',
        fill: `<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="000000"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:path path="shape"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>`,
      })
    );
    const scales = [...root.querySelectorAll('use')].map(
      (u) => /scale\(([^)]+)\)/.exec(u.getAttribute('transform') ?? '')![1]
    );
    expect(scales).toHaveLength(MAX_SHAPE_BANDS);
    expect(new Set(scales).size).toBe(MAX_SHAPE_BANDS);
  });

  test('path fill modes get their own paint servers; group fills band the child outline', () => {
    const cube = paint(
      shape({
        id: 2,
        x: 0,
        y: 0,
        cx: 1_000_000,
        cy: 1_000_000,
        prst: 'cube',
        fill: `<a:gradFill>${STOPS}<a:lin ang="0" scaled="0"/></a:gradFill>`,
      })
    );
    // cube fills with norm, darkenLess, and lightenLess: one server for each mode.
    expect(cube.querySelectorAll('linearGradient').length).toBe(3);
    const group = paint(
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="5" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/><a:chOff x="0" y="0"/><a:chExt cx="2000000" cy="1000000"/></a:xfrm>' +
        `<a:gradFill>${STOPS}<a:path path="shape"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></p:grpSpPr>` +
        shape({
          id: 6,
          x: 0,
          y: 0,
          cx: 2_000_000,
          cy: 1_000_000,
          prst: 'ellipse',
          fill: '<a:grpFill/>',
        }) +
        '</p:grpSp>'
    );
    expect(group.querySelectorAll('use').length).toBe(MAX_SHAPE_BANDS);
    // A child with two paths bands each outline under the group's shape gradient.
    const twoPaths = paint(
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="5" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/><a:chOff x="0" y="0"/><a:chExt cx="1000000" cy="1000000"/></a:xfrm>' +
        `<a:gradFill>${STOPS}<a:path path="shape"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></p:grpSpPr>` +
        '<p:sp><p:nvSpPr><p:cNvPr id="6" name="Two"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm>' +
        '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>' +
        '<a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="40" y="0"/></a:lnTo><a:lnTo><a:pt x="0" y="40"/></a:lnTo><a:close/></a:path>' +
        '<a:path w="100" h="100"><a:moveTo><a:pt x="100" y="100"/></a:moveTo><a:lnTo><a:pt x="60" y="100"/></a:lnTo><a:lnTo><a:pt x="100" y="60"/></a:lnTo><a:close/></a:path>' +
        '</a:pathLst></a:custGeom><a:grpFill/></p:spPr></p:sp></p:grpSp>'
    );
    expect(twoPaths.querySelectorAll('use').length).toBe(2 * MAX_SHAPE_BANDS);
    // A nested group with `a:grpFill` passes the outer group's fill on.
    const nested = paint(
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="5" name="Outer"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/><a:chOff x="0" y="0"/><a:chExt cx="1000000" cy="1000000"/></a:xfrm>' +
        '<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></p:grpSpPr>' +
        '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="6" name="Inner"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/><a:chOff x="0" y="0"/><a:chExt cx="1000000" cy="1000000"/></a:xfrm><a:grpFill/></p:grpSpPr>' +
        shape({ id: 7, x: 0, y: 0, cx: 1_000_000, cy: 1_000_000, fill: '<a:grpFill/>' }) +
        '</p:grpSp></p:grpSp>'
    );
    expect([...nested.querySelectorAll('path')].map((p) => p.getAttribute('fill'))).toContain(
      '#00ff00'
    );
  });

  test('a slide spends a fixed band budget; later shapes use the rectangle construction', () => {
    const outline = Array.from(
      { length: 1000 },
      (_, i) => `<a:lnTo><a:pt x="${i % 100}" y="${(i * 7) % 100}"/></a:lnTo>`
    ).join('');
    const shapes = Array.from(
      { length: 6 },
      (_, i) =>
        `<p:sp><p:nvSpPr><p:cNvPr id="${i + 2}" name="S"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm>` +
        `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo>${outline}<a:close/></a:path></a:pathLst></a:custGeom>` +
        `<a:gradFill>${STOPS}<a:path path="shape"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></p:spPr></p:sp>`
    ).join('');
    const uses = paint(shapes).querySelectorAll('use').length;
    expect(uses).toBeGreaterThan(0);
    expect(uses).toBeLessThan(6 * MAX_SHAPE_BANDS);
    expect(uses % MAX_SHAPE_BANDS).toBe(0);
  });

  test('a picture mask covers curves that leave the box', () => {
    const root = paint(
      '<p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
        '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
        '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm>' +
        '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="100" h="100">' +
        '<a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:cubicBezTo><a:pt x="150" y="0"/><a:pt x="150" y="-50"/><a:pt x="100" y="100"/></a:cubicBezTo><a:close/></a:path></a:pathLst></a:custGeom></p:spPr></p:pic>',
      [['rIdImg', IMAGE_REL, '../media/image1.png']]
    );
    const mask = root.querySelector('mask')!;
    expect(Number(mask.getAttribute('x')) + Number(mask.getAttribute('width'))).toBeGreaterThan(
      1_500_000
    );
    expect(Number(mask.getAttribute('y'))).toBeLessThan(-500_000);
  });

  test('paint and the approximation report drop overlays at the same path count', () => {
    expect(MAX_OVERLAY_PATHS).toBe(REPORTED_OVERLAY_PATHS);
  });

  test('a large path draws no bands, and a shape with many paths draws no overlay', () => {
    const big = Array.from(
      { length: 2100 },
      (_, i) => `<a:lnTo><a:pt x="${i % 100}" y="${(i * 7) % 100}"/></a:lnTo>`
    ).join('');
    const shapeXml = (paths: string, extra: string): string =>
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Big"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm>' +
      `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>${paths}</a:pathLst></a:custGeom>${extra}</p:spPr></p:sp>`;
    const large = paint(
      shapeXml(
        `<a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo>${big}<a:close/></a:path>`,
        `<a:gradFill>${STOPS}<a:path path="shape"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>`
      )
    );
    expect(large.querySelectorAll('use').length).toBe(0);
    // It falls back to the rectangle construction, not to no fill.
    expect(large.querySelectorAll('polygon').length).toBe(4);
    const many = Array.from(
      { length: 70 },
      (_, i) =>
        `<a:path w="100" h="100"><a:moveTo><a:pt x="${i}" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:close/></a:path>`
    ).join('');
    const overlaid = paint(
      shapeXml(
        many,
        '<a:solidFill><a:srgbClr val="808080"/></a:solidFill><a:effectLst><a:fillOverlay blend="mult"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:fillOverlay></a:effectLst>'
      )
    );
    expect(overlaid.querySelectorAll('[style="isolation: isolate"]').length).toBe(0);
  });

  test('a tiled picture keeps its duotone', () => {
    const root = paint(
      shape({
        id: 2,
        x: 0,
        y: 0,
        cx: 2_000_000,
        cy: 1_000_000,
        fill: '<a:blipFill><a:blip r:embed="rIdImg"><a:duotone><a:srgbClr val="000000"/><a:srgbClr val="FFFFFF"/></a:duotone></a:blip><a:tile sx="50000" sy="50000" algn="tl"/></a:blipFill>',
      }),
      [['rIdImg', IMAGE_REL, '../media/image1.png']]
    );
    const tile = root.querySelector('image')!;
    expect(tile.getAttribute('width')).toBe('609600');
    expect(tile.parentElement!.getAttribute('filter')).toMatch(/^url\(#t-blip/);
  });

  const tiled = (cx: number, cy: number, tile: string, prst = 'rect', alpha = ''): string =>
    shape({
      id: 2,
      x: 0,
      y: 0,
      cx,
      cy,
      prst,
      fill: `<a:blipFill><a:blip r:embed="rIdImg">${alpha}</a:blip><a:tile ${tile}/></a:blipFill>`,
      line: '<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>',
    });
  const IMG_REL: [string, string, string][] = [['rIdImg', IMAGE_REL, '../media/image1.png']];

  test('a tiled picture draws one image per tile inside its box, not a printed pattern', () => {
    // Tiles of 1_219_200 × 914_400 at 50 %, mirrored both ways, in a box of 4 × 3 tiles.
    const root = paint(
      tiled(
        2_438_400,
        1_371_600,
        'sx="50000" sy="50000" flip="xy" algn="tl"',
        'rect',
        '<a:alphaModFix amt="50000"/>'
      ),
      IMG_REL
    );
    expect([...root.querySelectorAll('pattern')].some((p) => p.id.includes('tile'))).toBe(false);
    const images = [...root.querySelectorAll('image')];
    expect(images.length).toBe(12);
    // The box clips the tiles as a nested viewport; the outline draws on top, unfilled.
    const box = images[0]!.parentElement!.parentElement!;
    expect(box.tagName.toLowerCase()).toBe('svg');
    expect([box.getAttribute('width'), box.getAttribute('height')]).toEqual(['2438400', '1371600']);
    expect(images[0]!.parentElement!.getAttribute('opacity')).toBe('0.5');
    expect(root.querySelector('path[stroke="#000000"]')!.getAttribute('fill')).toBe('none');
    // The second tile of a row turns about the far edge of its cell.
    expect(images[1]!.getAttribute('transform')).toBe('translate(1219200 0) scale(-1 1)');
    expect(images[2]!.getAttribute('transform')).toBe('translate(0 914400) scale(1 -1)');
    expect(images[3]!.getAttribute('transform')).toBe('translate(1219200 914400) scale(-1 -1)');
    expect([images[4]!.getAttribute('x'), images[4]!.getAttribute('y')]).toEqual(['1219200', '0']);
  });

  test('a tiled picture in another shape keeps inside it with a mask', () => {
    const root = paint(
      tiled(2_438_400, 1_371_600, 'sx="50000" sy="50000" algn="tl"', 'ellipse'),
      IMG_REL
    );
    const image = root.querySelector('image')!;
    expect(image.closest('g[mask]')).not.toBeNull();
  });

  test('a box of too many tiles fills with the pattern', () => {
    // Tiles of 9_525 EMU: 2_048 tiles at most, so this box draws the pattern.
    const root = paint(tiled(9_525 * 50, 9_525 * 50, 'sx="781" sy="1041" algn="tl"'), IMG_REL);
    expect(root.querySelectorAll('image').length).toBe(1);
    expect([...root.querySelectorAll('pattern')].some((p) => p.id.includes('tile'))).toBe(true);
  });

  test('a tiled background draws its tiles in the slide viewport', () => {
    const r = readPptx(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            rels: IMG_REL,
            background:
              '<p:bg><p:bgPr><a:blipFill><a:blip r:embed="rIdImg"/><a:tile sx="100000" sy="100000" algn="tl"/></a:blipFill><a:effectLst/></p:bgPr></p:bg>',
            shapes: '',
          },
        ],
      })
    );
    if (!r.ok) throw new Error('open failed');
    const layout = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer());
    const root = paintSlide(doc, layout, {
      images: () => 'blob:test',
      imageSize: () => ({ cx: 1_219_200, cy: 914_400 }),
      idPrefix: 't',
    });
    const bg = root.querySelector('[data-kind="background"]')!;
    expect(bg.tagName.toLowerCase()).toBe('svg');
    // 12_192_000 × 6_858_000 in tiles of 1_219_200 × 914_400: 10 × 8 tiles, the last row cut.
    expect(bg.querySelectorAll('image').length).toBe(80);
  });

  test('tile placements cover the box once and stop at the image budget', () => {
    const g = { x: -100, y: 0, width: 400, height: 300, flipX: true, flipY: false };
    const tiles = tilePlacements(g, 1000, 300)!;
    // Cells of 800 from -100: tiles at -100, 300, 700 meet the box; 1100 does not.
    expect(tiles.map((t) => [t.x + (t.mirrorX ? g.width : 0), t.mirrorX])).toEqual([
      [-100, false],
      [300, true],
      [700, false],
    ]);
    const tiny = { ...g, x: 0, width: 1, height: 1, flipX: false };
    expect(tilePlacements(tiny, MAX_TILE_IMAGES, 1)!.length).toBe(MAX_TILE_IMAGES);
    expect(tilePlacements(tiny, MAX_TILE_IMAGES + 1, 1)).toBeUndefined();
  });

  const picture = (crop: string): string =>
    '<p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
    `<p:blipFill><a:blip r:embed="rIdImg"/>${crop}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    '<p:spPr><a:xfrm><a:off x="-2286000" y="0"/><a:ext cx="12192000" cy="6858000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';

  test('a cropped picture clips to its geometry with a mask, not a clip path', () => {
    const root = paint(picture('<a:srcRect l="10000" r="10000"/>'), [
      ['rIdImg', IMAGE_REL, '../media/image1.png'],
    ]);
    expect(root.querySelectorAll('clipPath').length).toBe(0);
    const mask = root.querySelector('mask')!;
    expect(mask.getAttribute('maskUnits')).toBe('userSpaceOnUse');
    expect(Number(mask.getAttribute('width'))).toBeGreaterThanOrEqual(12_192_000);
  });

  test('a rectangle picture inside its box needs no clip', () => {
    const root = paint(picture(''), [['rIdImg', IMAGE_REL, '../media/image1.png']]);
    expect(root.querySelectorAll('clipPath, mask').length).toBe(0);
    expect(root.querySelector('image')!.getAttribute('width')).toBe('12192000');
  });
});
