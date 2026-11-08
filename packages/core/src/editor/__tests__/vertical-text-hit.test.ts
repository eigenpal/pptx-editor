// Text in a turned frame whose lines do not run in place order: `mongolianVert` and
// `wordArtVert` place the first line at the left edge of the shape. A press picks the line under
// it, the caret of stacked text covers its letter cells, and `spAutoFit` sizes the columns of
// turned text with the insets of its frame.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '../../store/index.ts';
import { resolveSlide, type ResolvedShape } from '../../cascade/index.ts';
import { createFixedMeasurer, layoutSlide, type LayoutItem } from '../../layout/index.ts';
import { fitShape } from '../../layout/autofit.ts';
import { caretAt, pointAt, textFrameMatrix } from '../surface-geometry.ts';
import { moveCaret } from '../text-input.ts';

const IN = 914_400;
const TEXT = para(run('one')) + para(run('two')) + para(run('three'));

function slideOf(bodyPr: string, text = TEXT, cx = 6 * IN, cy = 4 * IN) {
  const bytes = buildPptx({
    slides: [{ layout: 'blank', shapes: shape({ id: 2, x: 0, y: 0, cx, cy, bodyPr, text }) }],
  });
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('open failed');
  const slide = resolveSlide(r.pkg, 0)!;
  return { slide, item: layoutSlide(slide, createFixedMeasurer()).items[0] };
}

type ShapeItem = Extract<LayoutItem, { kind: 'shape' }>;

describe('turned lines out of place order', () => {
  for (const vert of ['vert', 'wordArtVert', 'mongolianVert'])
    test(`a press on each ${vert} line places the caret in it`, () => {
      const item = slideOf(`<a:bodyPr vert="${vert}"/>`).item as ShapeItem;
      const m = textFrameMatrix(item)!;
      const hits = item.text!.lines.map((l) => {
        const fx = l.caretX[1];
        const fy = l.top + l.height / 2;
        return pointAt(item, m[0] * fx + m[2] * fy + m[4], m[1] * fx + m[3] * fy + m[5])?.paragraph;
      });
      expect(hits).toEqual([0, 1, 2]);
    });

  test('Down and Up step through the paragraphs in their order', () => {
    const item = slideOf('<a:bodyPr vert="mongolianVert"/>').item as ShapeItem;
    const down = moveCaret(item, { paragraph: 0, offset: 1 }, 'ArrowDown');
    expect(down.paragraph).toBe(1);
    expect(moveCaret(item, down, 'ArrowUp').paragraph).toBe(0);
  });

  test('the caret of stacked text is centered on its letter axis', () => {
    const item = slideOf('<a:bodyPr vert="wordArtVert"/>').item as ShapeItem;
    const line = item.text!.lines[0];
    const caret = caretAt(item, { paragraph: 0, offset: 1 })!;
    const m = textFrameMatrix(item)!;
    // The frame turns 90°: its y runs across the slide's x.
    const mid = (caret.x1 + caret.x2) / 2;
    const axis = m[0] * line.caretX[1] + m[2] * line.baseline + m[4];
    expect(mid).toBeCloseTo(axis, 0);
  });
});

describe('spAutoFit of turned columns', () => {
  const fit = (bodyPr: string, cx: number, cy: number): number => {
    const { slide } = slideOf(bodyPr, TEXT + TEXT, cx, cy);
    const s = slide.elements[0] as ResolvedShape;
    const f = fitShape(s, createFixedMeasurer());
    if (!f || f.kind !== 'size') throw new Error('no size');
    return s.text!.body.vert === 'horz' ? f.cy : f.cx;
  };
  const insets = 'lIns="10000" tIns="20000" rIns="30000" bIns="40000"';

  test('vert columns grow as horizontal columns in the turned box with turned insets', () => {
    const vert = fit(
      `<a:bodyPr vert="vert" numCol="2" ${insets}><a:spAutoFit/></a:bodyPr>`,
      IN,
      3 * IN
    );
    // The same frame unturned: the lines start at tIns and the text top is at rIns.
    const horz = fit(
      `<a:bodyPr numCol="2" lIns="20000" tIns="30000" rIns="40000" bIns="10000"><a:spAutoFit/></a:bodyPr>`,
      3 * IN,
      IN
    );
    expect(vert).toBe(horz);
  });

  test('mongolianVert columns grow as vert columns', () => {
    const body = (vert: string) =>
      `<a:bodyPr vert="${vert}" numCol="2" lIns="20000" rIns="20000"><a:spAutoFit/></a:bodyPr>`;
    expect(fit(body('mongolianVert'), IN, 3 * IN)).toBe(fit(body('vert'), IN, 3 * IN));
  });
});
