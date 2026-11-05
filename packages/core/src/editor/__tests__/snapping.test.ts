// Smart guides and snapping: each guide kind, the tolerance, nearest-wins, per-axis snapping,
// resize edges, the off switch (Alt or View > Smart guides), and the group scope.
import { describe, expect, test } from 'bun:test';
import { buildPptx, shape, solid } from '../../../../../scripts/fixtures/builder.ts';
import { createPptxEditor, toolbarCommandState } from '../index.ts';
import { snapBox, type SnapBounds, type SnapGuide, type SnapRequest } from '../snapping.ts';
import { resizeLines, snapContext, snapMove, snapResize } from '../surface-guides.ts';

const SLIDE = { width: 12_000, height: 6_000 };
const box = (x: number, y: number, w: number, h: number): SnapBounds => ({
  x,
  y,
  r: x + w,
  b: y + h,
});
const snap = (moving: SnapBounds, targets: SnapBounds[], extra: Partial<SnapRequest> = {}) =>
  snapBox({ moving, targets, slide: SLIDE, tolerance: 50, ...extra });
const aligns = (guides: readonly SnapGuide[], axis: 'x' | 'y') =>
  guides.filter((g) => g.kind === 'align' && g.axis === axis);
const gaps = (guides: readonly SnapGuide[], axis: 'x' | 'y') =>
  guides.filter((g) => g.kind === 'gap' && g.axis === axis);

describe('alignment', () => {
  test('a left edge snaps to another left edge and shows a line across both boxes', () => {
    const target = box(1_000, 1_000, 2_000, 1_000);
    const r = snap(box(1_030, 3_000, 500, 500), [target]);
    expect(r.dx).toBe(-30);
    expect(r.dy).toBe(0);
    expect(aligns(r.guides, 'x')).toEqual([
      { kind: 'align', axis: 'x', at: 1_000, from: 1_000, to: 3_500, slide: false },
    ]);
  });

  test('a left edge snaps to a right edge (shapes side by side)', () => {
    const r = snap(box(3_020, 1_200, 500, 500), [box(1_000, 1_000, 2_000, 1_000)]);
    expect(r.dx).toBe(-20);
    expect(aligns(r.guides, 'x').map((g) => g.at)).toEqual([3_000]);
  });

  test('centers snap to centers; top, middle, and bottom on the other axis', () => {
    const target = box(1_000, 1_000, 2_000, 1_000);
    // Center x: 2000 vs 1990 + 10; middle y: 1500 vs 1525.
    const r = snap(box(4_000 - 2_010, 3_000, 0, 0), [target]);
    expect(r.dx).toBe(10);
    const m = snap(box(5_000, 1_275, 500, 500), [target]);
    expect(m.dy).toBe(-25);
    expect(aligns(m.guides, 'y').map((g) => g.at)).toEqual([1_500]);
  });

  test('the slide center and the slide edges', () => {
    const center = snap(box(5_740, 2_000, 500, 500), []);
    expect(center.dx).toBe(10);
    expect(aligns(center.guides, 'x')).toEqual([
      { kind: 'align', axis: 'x', at: 6_000, from: 0, to: SLIDE.height, slide: true },
    ]);
    const edge = snap(box(20, 5_470, 500, 500), []);
    expect([edge.dx, edge.dy]).toEqual([-20, 30]);
    expect(aligns(edge.guides, 'y').map((g) => g.at)).toEqual([6_000]);
  });
});

describe('tolerance and choice', () => {
  test('nothing snaps beyond the tolerance, and no guide shows', () => {
    const r = snap(box(1_060, 3_000, 500, 500), [box(1_000, 1_000, 2_000, 1_000)]);
    expect(r).toEqual({ dx: 0, dy: 0, guides: [] });
  });

  test('the nearest candidate wins', () => {
    const r = snap(box(1_000, 3_000, 500, 500), [box(970, 0, 100, 100), box(1_010, 200, 100, 100)]);
    expect(r.dx).toBe(10);
  });

  test('each axis snaps on its own', () => {
    const r = snap(box(1_040, 3_000, 500, 500), [box(1_000, 1_000, 2_000, 1_000)]);
    expect(r.dx).toBe(-40);
    expect(r.dy).toBe(0);
    expect(aligns(r.guides, 'y')).toEqual([]);
  });
});

describe('equal spacing', () => {
  test('between two shapes in a row, with equal gaps on both sides', () => {
    const left = box(0, 1_000, 1_000, 1_000);
    const right = box(3_000, 1_000, 1_000, 1_000);
    // Gaps 760 and 740: equal at x = 1750.
    const r = snap(box(1_760, 1_100, 500, 500), [left, right], { tolerance: 20 });
    expect(r.dx).toBe(-10);
    expect(gaps(r.guides, 'x').map((g) => [g.from, g.to])).toEqual([
      [1_000, 1_750],
      [2_250, 3_000],
    ]);
  });

  test('beside a pair, at the pair gap', () => {
    const a = box(0, 1_000, 1_000, 1_000);
    const b = box(1_400, 1_000, 1_000, 1_000);
    const r = snap(box(2_815, 1_000, 1_000, 1_000), [a, b], { tolerance: 20 });
    expect(r.dx).toBe(-15);
    expect(gaps(r.guides, 'x').map((g) => [g.from, g.to])).toEqual([
      [1_000, 1_400],
      [2_400, 2_800],
    ]);
  });

  test('shapes in another row do not count', () => {
    const r = snap(
      box(1_510, 4_000, 500, 500),
      [box(0, 0, 1_000, 1_000), box(3_000, 0, 1_000, 1_000)],
      {
        tolerance: 20,
      }
    );
    expect(gaps(r.guides, 'x')).toEqual([]);
  });
});

describe('moves and resizes', () => {
  const ctx = { targets: [box(1_000, 1_000, 2_000, 1_000)], slide: SLIDE };

  test('a resize snaps only the dragged edge', () => {
    // The right edge (end) at 2985 snaps to 3000; its left edge is never a candidate.
    const r = snapResize(ctx, [], box(2_000, 3_000, 900, 500), [1, 0.5], [85, 0], false, 50);
    expect(r.dx).toBe(100);
    expect(aligns(r.guides, 'x').map((g) => g.at)).toEqual([3_000]);
    const free = snapResize(ctx, [], box(1_020, 3_000, 900, 500), [1, 0.5], [0, 0], false, 50);
    expect(free.dx).toBe(0);
  });

  test('Shift at a corner keeps the aspect ratio and does not snap', () => {
    const r = snapResize(ctx, [], box(2_000, 3_000, 900, 500), [1, 1], [85, 0], true, 50);
    expect(r).toEqual({ dx: 85, dy: 0, guides: [] });
  });

  test('Alt or View > Smart guides off (no tolerance): no snap and no guides', () => {
    expect(snapMove(ctx, box(1_030, 3_000, 500, 500), 0, 0, false, null)).toEqual({
      dx: 0,
      dy: 0,
      guides: [],
    });
    expect(snapMove(ctx, box(1_030, 3_000, 500, 500), 0, 0, false, 50).dx).toBe(-30);
  });

  test('Shift keeps a move on one axis, and that axis does not snap', () => {
    // Unlocked, the top edge (2030) would snap to the target's bottom edge (2000).
    expect(snapMove(ctx, box(1_000, 2_010, 500, 500), 400, 20, false, 50).dy).toBe(-10);
    const r = snapMove(ctx, box(1_000, 2_010, 500, 500), 400, 20, true, 50);
    expect(r.dy).toBe(0);
    expect(r.dx).toBe(400);
  });
});

describe('scope', () => {
  const IN = 914_400;
  const GROUP =
    '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group 9"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="914400"/>' +
    '<a:chOff x="914400" y="914400"/><a:chExt cx="2743200" cy="914400"/></a:xfrm></p:grpSpPr>' +
    shape({ id: 11, x: IN, y: IN, cx: IN, cy: IN, fill: solid('4472C4') }) +
    shape({ id: 12, x: 3 * IN, y: IN, cx: IN, cy: IN, fill: solid('ED7D31') }) +
    '</p:grpSp>';
  const deck = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          GROUP +
          shape({ id: 2, x: 6 * IN, y: IN, cx: 2 * IN, cy: IN }).replace(
            '<a:xfrm>',
            '<a:xfrm rot="5400000">'
          ),
      },
    ],
  });

  test('a group child snaps to its siblings; a top-level shape to top-level shapes', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(deck)).ok).toBe(true);
    const layout = editor.session!.layout(0)!;
    const group = layout.items.find((i) => i.kind === 'group')!;
    const loose = layout.items.find((i) => i.kind !== 'group')!;
    if (group.kind !== 'group') throw new Error('no group');
    const [first, second] = group.children;
    const inGroup = snapContext(layout, [first.element.nodeId]);
    expect(inGroup.targets).toHaveLength(1);
    expect(inGroup.targets[0].x).toBeCloseTo(3 * IN, 0);
    const top = snapContext(layout, [loose.element.nodeId]);
    expect(top.targets).toHaveLength(1);
    expect(top.targets[0].r).toBeCloseTo(4 * IN, 0);
    expect(second.element.nodeId).not.toBe(first.element.nodeId);
    // The 90° turned shape: its bounding box is 1 in wide and 2 in tall.
    const turned = snapContext(layout, [group.element.nodeId]).targets[0];
    expect(turned.r - turned.x).toBeCloseTo(IN, 0);
    expect(turned.b - turned.y).toBeCloseTo(2 * IN, 0);
    // Its resize handles map to the bounds' edges; the right handle moves the bottom edge.
    expect(resizeLines(loose, 1, 0.5)).toEqual({ x: [], y: ['end'] });
    editor.destroy();
  });
});

describe('View > Smart guides', () => {
  test('is on by default, and the toggle is editor state, not an edit', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(buildPptx({ slides: [{ layout: 'blank', shapes: '' }] }))).ok).toBe(
      true
    );
    expect(editor.snapshot().smartGuides).toBe(true);
    expect(toolbarCommandState(editor, 'view.smartGuides').active).toBe(true);
    expect(editor.exec({ type: 'toggleSmartGuides' }).ok).toBe(true);
    expect(editor.snapshot().smartGuides).toBe(false);
    expect(toolbarCommandState(editor, 'view.smartGuides')).toMatchObject({
      enabled: true,
      active: false,
    });
    expect(editor.snapshot().canUndo).toBe(false);
    editor.destroy();
  });
});
