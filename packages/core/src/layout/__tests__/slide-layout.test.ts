import { describe, expect, test } from 'bun:test';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '../../store/index.ts';
import { resolveSlide } from '../../cascade/index.ts';
import {
  apply,
  createFixedMeasurer,
  layoutSlide,
  type LayoutItem,
  type Matrix,
  type SlideLayout,
} from '../index.ts';

const text = (id: number, x: number, flip = ''): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="S${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm${flip}><a:off x="${x}" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
  '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp>';

const group = (flip: string, rot = 0, child = text(11, 0)): string =>
  '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  `<p:grpSpPr><a:xfrm${flip}${rot ? ` rot="${rot}"` : ''}><a:off x="0" y="0"/><a:ext cx="3000000" cy="1000000"/>` +
  '<a:chOff x="0" y="0"/><a:chExt cx="3000000" cy="1000000"/></a:xfrm></p:grpSpPr>' +
  `${child}</p:grpSp>`;

function open(shapes: string): SlideLayout {
  const r = readPptx(buildPptx({ slides: [{ shapes, layout: 'blank' }] }));
  if (!r.ok) throw new Error('open failed');
  return layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer());
}

function firstShape(items: readonly LayoutItem[]): Extract<LayoutItem, { kind: 'shape' }> {
  for (const it of items) {
    if (it.kind === 'shape') return it;
    if (it.kind === 'group') return firstShape(it.children);
  }
  throw new Error('no shape');
}

const det = (m: Matrix): number => m[0] * m[3] - m[1] * m[2];
const round = (m: Matrix): number[] => m.slice(0, 4).map((v) => Math.round(v * 1000) / 1000 + 0);

describe('text in flipped groups (R2)', () => {
  test('a horizontally flipped group keeps its text upright and readable', () => {
    const shape = firstShape(open(group(' flipH="1"')).items);
    expect(det(shape.textMatrix!)).toBeGreaterThan(0);
    expect(round(shape.textMatrix!)).toEqual([1, 0, 0, 1]);
    // The text sits where the mirrored child is: the right third of the group.
    const c = apply(shape.textMatrix!, 500_000, 500_000);
    expect(c.x).toBeCloseTo(2_500_000, 0);
  });

  test('a vertically flipped group turns its text 180 degrees', () => {
    const shape = firstShape(open(group(' flipV="1"')).items);
    expect(det(shape.textMatrix!)).toBeGreaterThan(0);
    expect(round(shape.textMatrix!)).toEqual([-1, 0, 0, -1]);
  });

  test('flips cancel: a flipV child in a flipV group is upright', () => {
    const shape = firstShape(open(group(' flipV="1"', 0, text(11, 0, ' flipV="1"'))).items);
    expect(round(shape.textMatrix!)).toEqual([1, 0, 0, 1]);
  });

  test('a rotated, mirrored group mirrors the rotation, never the glyphs', () => {
    const shape = firstShape(open(group(' flipH="1"', 30 * 60_000)).items);
    const m = shape.textMatrix!;
    expect(det(m)).toBeGreaterThan(0);
    expect((Math.atan2(m[1], m[0]) * 180) / Math.PI).toBeCloseTo(30, 6);
  });

  test('text in a flipped group matches the same shape flipped on its own', () => {
    // Ungrouping a flipped group gives the child the group's flip: the text must not move.
    const grouped = firstShape(open(group(' flipH="1"', 0, text(11, 2_000_000))).items);
    const alone = firstShape(open(text(11, 0, ' flipH="1"')).items);
    expect(round(grouped.textMatrix!)).toEqual(round(alone.textMatrix!));
    expect(grouped.textMatrix![4]).toBeCloseTo(alone.textMatrix![4], 3);
  });

  test('shapes flipped on their own keep the existing rule', () => {
    expect(round(firstShape(open(text(2, 0, ' flipH="1"')).items).textMatrix!)).toEqual([
      1, 0, 0, 1,
    ]);
    expect(round(firstShape(open(text(2, 0, ' flipV="1"')).items).textMatrix!)).toEqual([
      -1, 0, 0, -1,
    ]);
  });
});

describe('an empty text body', () => {
  // A text box and a placeholder with no character still get one line: the caret of a press
  // in them needs its position and height. The line paints no glyph and no bullet.
  const empty =
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Box"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="400000"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>' +
    '<a:p><a:pPr><a:buChar char="&#8226;"/></a:pPr><a:endParaRPr lang="en-US" sz="2000"/></a:p>' +
    '</p:txBody></p:sp>';

  test('has one line with no glyph run and no bullet', () => {
    const shape = firstShape(open(empty).items);
    expect(shape.element.kind === 'shape' && shape.element.textBox).toBe(true);
    expect(shape.text?.lines.length).toBe(1);
    const line = shape.text!.lines[0];
    expect(line.end).toBe(line.start);
    expect(line.runs.filter((r) => r.kind !== 'bullet' && r.text.length > 0)).toEqual([]);
    expect(line.runs.some((r) => r.kind === 'bullet')).toBe(false);
    expect(line.fontHeight).toBeGreaterThan(0);
  });
});
