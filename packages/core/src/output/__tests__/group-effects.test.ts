// Group effects (`p:grpSpPr/a:effectLst`, ECMA-376 Part 1, 19.3.1.22): the effects apply to the
// group's content as one picture. The outer shadow is one copy of all children together, placed
// by the group's box under its turn, flips, and child space, and it paints without the Pro
// drawing renderers.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import {
  buildPptx,
  group,
  para,
  run,
  shape,
  solid,
} from '../../../../../scripts/fixtures/builder.ts';
import { openPresentation } from '../../editor/document-session.ts';
import { apply, type LayoutItem, type Matrix, type SlideLayout } from '../../layout/index.ts';
import { effectMatrix } from '../effect-transforms.ts';
import { groupTextCovers, MAX_GROUP_TEXT_CLIPS } from '../group-text-covers.ts';
import { paintSlide, type PaintOptions } from '../paint.ts';
import { SlidePainter } from '../slide-painter.ts';

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const IN = 914_400;
const OPTS: PaintOptions = { images: () => undefined, idPrefix: 'g', hitTargets: false };
const DIST = 152_400;

const shadow = (attrs = '') =>
  `<a:effectLst><a:outerShdw blurRad="0" dist="${DIST}" dir="0"${attrs}>` +
  '<a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:outerShdw></a:effectLst>';

/** Two overlapping rectangles in a group scaled 2× by its child space, turned and flipped. */
const groupDeck = (
  effects: string,
  xfrm: { rot?: number; flipH?: boolean; flipV?: boolean } = {}
) =>
  group(
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
      effects,
      ...xfrm,
    },
    shape({ id: 11, x: 0, y: 0, cx: IN, cy: 0.5 * IN, fill: solid('C00000') }) +
      shape({ id: 12, x: 0.75 * IN, y: 0.25 * IN, cx: IN, cy: 0.75 * IN, fill: solid('0070C0') })
  );

async function layoutOf(shapes: string): Promise<SlideLayout> {
  const opened = await openPresentation(buildPptx({ slides: [{ layout: 'blank', shapes }] }), {
    fixedMeasurer: true,
    substitutes: false,
  });
  if (!opened.ok) throw new Error('deck');
  return opened.session.layout(0)!;
}

type GroupItem = Extract<LayoutItem, { kind: 'group' }>;
const groupOf = (layout: SlideLayout): GroupItem => {
  const item = layout.items[0];
  if (item?.kind !== 'group') throw new Error('no group');
  return item;
};

/** The matrix of a `transform="matrix(…)"` attribute. */
function matrixOf(el: Element | null): Matrix {
  const m = /matrix\(([^)]+)\)/.exec(el?.getAttribute('transform') ?? '');
  if (!m) throw new Error('no matrix');
  return m[1]!.trim().split(/\s+/).map(Number) as unknown as Matrix;
}

/** Whether two points are within `tol` EMU. */
const near = (a: { x: number; y: number }, b: { x: number; y: number }, tol = 2000): void => {
  expect(Math.abs(a.x - b.x)).toBeLessThan(tol);
  expect(Math.abs(a.y - b.y)).toBeLessThan(tol);
};

describe('group effects', () => {
  test('the cascade resolves the effect list of the group', async () => {
    const layout = await layoutOf(groupDeck(shadow()));
    const fx = groupOf(layout).element.effects;
    expect(fx?.outerShadow?.dist).toBe(DIST);
    expect(fx?.outerShadow?.color.a).toBeCloseTo(0.5, 3);
    expect(layout.diagnostics).toContain('approximation: group-effects');
  });

  test("a group's inner shadow resolves to nothing; its other effects stay", async () => {
    const inner =
      '<a:effectLst><a:innerShdw blurRad="63500" dist="50800" dir="5400000">' +
      '<a:srgbClr val="000000"/></a:innerShdw></a:effectLst>';
    const alone = groupOf(await layoutOf(groupDeck(inner))).element.effects;
    expect(alone?.innerShadow).toBeUndefined();
    const both = inner.replace('</a:effectLst>', shadow().slice('<a:effectLst>'.length));
    const fx = groupOf(await layoutOf(groupDeck(both))).element.effects;
    expect(fx?.innerShadow).toBeUndefined();
    expect(fx?.outerShadow?.dist).toBe(DIST);
  });

  test('only a turned 3-D scene of a group is reported unsupported', async () => {
    const scene = (camera: string) =>
      `<a:scene3d><a:camera prst="${camera}"/>` + '<a:lightRig rig="threePt" dir="t"/></a:scene3d>';
    const flat = groupOf(await layoutOf(groupDeck(shadow() + scene('orthographicFront'))));
    expect(flat.element.effects?.unsupported).not.toContain('scene3d');
    const turned = groupOf(await layoutOf(groupDeck(shadow() + scene('isometricTopUp'))));
    expect(turned.element.effects?.unsupported).toContain('scene3d');
  });

  test('one shadow copies all children together; the children cast none', async () => {
    const layout = await layoutOf(groupDeck(shadow(), { rot: 1_800_000, flipH: true }));
    const svg = paintSlide(doc, layout, OPTS);
    const shadows = svg.querySelectorAll('[data-effect="shadow"]');
    expect(shadows.length).toBe(1);
    const node = svg.querySelector('[data-kind="group"]')!;
    expect(shadows[0]!.parentNode === node).toBe(true);
    const use = shadows[0]!.querySelector('use')!;
    const id = use.getAttribute('href')!.slice(1);
    const content = svg.querySelector(`[id="${id}"]`)!;
    expect(content.querySelectorAll('[data-kind="shape"]').length).toBe(2);
    // The content paints after its shadow, inside the group's node.
    expect(content.parentNode === node).toBe(true);
    expect([...node.children].indexOf(content)).toBeGreaterThan(
      [...node.children].indexOf(shadows[0]!)
    );
  });

  test('the shadow turns and flips with the group, by the group box', async () => {
    const layout = await layoutOf(groupDeck(shadow(), { rot: 1_800_000, flipH: true }));
    const item = groupOf(layout);
    const use = paintSlide(doc, layout, OPTS).querySelector('[data-effect="shadow"] use');
    const m = matrixOf(use);
    const { cx, cy } = item.element.xfrm;
    const expected = effectMatrix(item.element.effects!.outerShadow!, item.matrix, cx, cy);
    for (let i = 0; i < 4; i += 1) expect(m[i]!).toBeCloseTo(expected[i]!, 3);
    near({ x: m[4]!, y: m[5]! }, { x: expected[4]!, y: expected[5]! }, 2);
    // `dir` 0 along the group's own x axis: turned 30° and mirrored by the flip.
    const center = apply(item.matrix, cx / 2, cy / 2);
    const moved = apply(m, center.x, center.y);
    const turn = (30 * Math.PI) / 180;
    near(
      { x: moved.x - center.x, y: moved.y - center.y },
      { x: -Math.cos(turn) * DIST, y: -Math.sin(turn) * DIST }
    );
  });

  test('without rotWithShape, the shadow keeps its slide direction', async () => {
    const layout = await layoutOf(
      groupDeck(shadow(' rotWithShape="0"'), { rot: 1_800_000, flipV: true })
    );
    const item = groupOf(layout);
    const m = matrixOf(paintSlide(doc, layout, OPTS).querySelector('[data-effect="shadow"] use'));
    const p = apply(item.matrix, 0, 0);
    near(apply(m, p.x, p.y), { x: p.x + DIST, y: p.y }, 2);
  });

  test('the filter region holds the children placed by the child space', async () => {
    const layout = await layoutOf(groupDeck(shadow()));
    const item = groupOf(layout);
    // The second child reaches 1.75 in × 1 in in a child space scaled 2×: the group's corner.
    const child = item.children[1]!;
    const corner = apply(child.matrix, child.element.xfrm.cx, child.element.xfrm.cy);
    near(corner, { x: 2 * IN + 3.5 * IN, y: 1.5 * IN + 2 * IN });
    const svg = paintSlide(doc, layout, OPTS);
    const filter = svg.querySelector('filter')!;
    const x = Number(filter.getAttribute('x'));
    const y = Number(filter.getAttribute('y'));
    const r = x + Number(filter.getAttribute('width'));
    const b = y + Number(filter.getAttribute('height'));
    expect(x).toBeLessThanOrEqual(2 * IN + DIST);
    expect(y).toBeLessThanOrEqual(1.5 * IN);
    expect(r).toBeGreaterThanOrEqual(corner.x + DIST);
    expect(b).toBeGreaterThanOrEqual(corner.y);
  });

  test('a group without effects paints as before', async () => {
    const svg = paintSlide(doc, await layoutOf(groupDeck('')), OPTS);
    expect(svg.querySelectorAll('[data-effect]').length).toBe(0);
    const node = svg.querySelector('[data-kind="group"]')!;
    expect([...node.children].map((c) => c.getAttribute('data-kind'))).toEqual(['shape', 'shape']);
  });

  test('a fill overlay on a group is reported as not drawn', async () => {
    const overlay =
      '<a:effectLst><a:fillOverlay blend="mult"><a:solidFill><a:srgbClr val="FF0000"/>' +
      '</a:solidFill></a:fillOverlay></a:effectLst>';
    const layout = await layoutOf(groupDeck(overlay));
    expect(layout.diagnostics).toContain('effect not rendered: fillOverlay');
    expect(layout.diagnostics).toContain('approximation: effects-not-drawn');
  });

  test('a budgeted paint paints a group with effects whole', async () => {
    const layout = await layoutOf(groupDeck(shadow()));
    const painter = new SlidePainter(doc, OPTS);
    const svg = painter.paint(layout, 1);
    expect(painter.pending).toBe(false);
    const id = svg.querySelector('[data-effect="shadow"] use')!.getAttribute('href')!.slice(1);
    expect(svg.querySelector(`[id="${id}"]`)!.querySelectorAll('[data-kind="shape"]').length).toBe(
      2
    );
  });

  test('only the two outermost of nested shadowed groups draw their shadows', async () => {
    // Each shadow copies the content that holds the shadows inside it: copies would double per
    // level. A stack of 16 paints at most two shadows.
    let shapes = shape({ id: 300, x: 0, y: 0, cx: IN, cy: IN, fill: solid('C00000') });
    for (let level = 15; level >= 0; level -= 1)
      shapes = group(
        { id: 200 + level, x: 0, y: 0, cx: 2 * IN, cy: 2 * IN, effects: shadow() },
        shapes
      );
    const started = performance.now();
    const layout = await layoutOf(shapes);
    const svg = paintSlide(doc, layout, OPTS);
    expect(performance.now() - started).toBeLessThan(2_000);
    expect(svg.querySelectorAll('[data-effect="shadow"]').length).toBe(2);
    expect(layout.diagnostics).toContain('effect not rendered: nested group effects');
  });

  test('the effects of a group of more than 1,000 elements do not draw', async () => {
    const many = Array.from({ length: 1_001 }, (_, i) =>
      shape({ id: 1_000 + i, x: 0, y: 0, cx: IN, cy: IN, fill: solid('C00000') })
    ).join('');
    const layout = await layoutOf(
      group({ id: 10, x: 0, y: 0, cx: IN, cy: IN, effects: shadow() }, many)
    );
    expect(paintSlide(doc, layout, OPTS).querySelector('[data-effect="shadow"]')).toBeNull();
    expect(layout.diagnostics).toContain('effect not rendered: group effects of a large group');
  });
});

describe('text under a later body in a group', () => {
  const labelled = (id: number, x: number, y: number, fill?: string) =>
    shape({
      id,
      x,
      y,
      cx: 2 * IN,
      cy: IN,
      ...(fill ? { fill: solid(fill) } : {}),
      text: para(run(`Shape ${id}`)),
    });
  const coversOf = async (children: string) =>
    groupTextCovers(
      groupOf(await layoutOf(group({ id: 10, x: IN, y: IN, cx: 5 * IN, cy: 3 * IN }, children)))
    );
  const covered = async (children: string): Promise<string[]> =>
    [...(await coversOf(children)).clips.keys()].map((i) => i.element.drawingId!).sort();

  test('a later filled shape over a text covers it; an earlier one does not', async () => {
    // Shape 12 lies over the text of 11; the text of 12 lies over the body of 11, which is earlier.
    expect(
      await covered(labelled(11, IN, IN, 'C00000') + labelled(12, 2 * IN, IN, '0070C0'))
    ).toEqual(['11']);
  });

  test('a later text box with no fill or outline covers nothing', async () => {
    expect(await covered(labelled(11, IN, IN, 'C00000') + labelled(12, 2 * IN, IN))).toEqual([]);
  });

  test('shapes apart cover nothing, and a nested group adds its shapes in paint order', async () => {
    expect(
      await covered(labelled(11, IN, IN, 'C00000') + labelled(12, IN, 3 * IN, '0070C0'))
    ).toEqual([]);
    const nested = group(
      { id: 20, x: 2 * IN, y: IN, cx: 2 * IN, cy: IN },
      shape({ id: 21, x: 2 * IN, y: IN, cx: 2 * IN, cy: IN, fill: solid('00B050') })
    );
    expect(await covered(labelled(11, IN, IN, 'C00000') + nested)).toEqual(['11']);
  });

  test('each clip is a rectangle around the text and the outline of one later body', async () => {
    const { clips, complete } = await coversOf(
      labelled(11, IN, IN, 'C00000') + labelled(12, 2 * IN, IN, '0070C0')
    );
    expect(complete).toBe(true);
    const [paths] = [...clips.values()];
    expect(paths!.length).toBe(1);
    // The rectangle (5 commands), then the rectangle of shape 12 in slide space.
    const path = paths![0]!;
    expect(path.slice(0, 5).map((c) => c.op)).toEqual(['M', 'L', 'L', 'L', 'Z']);
    const xs = path.slice(5).flatMap((c) => ('x' in c ? [c.x] : []));
    expect([Math.min(...xs), Math.max(...xs)]).toEqual([2 * IN, 4 * IN]);
  });

  test('past the clip budget, the text keeps no clip and the result is incomplete', async () => {
    const many = Array.from({ length: MAX_GROUP_TEXT_CLIPS + 1 }, (_, i) =>
      shape({ id: 100 + i, x: IN, y: IN, cx: IN, cy: IN, fill: solid('0070C0') })
    ).join('');
    const { clips, complete } = await coversOf(labelled(11, IN, IN, 'C00000') + many);
    expect(clips.size).toBe(0);
    expect(complete).toBe(false);
  });
});
