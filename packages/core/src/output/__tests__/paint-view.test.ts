// Paint for a view (`PaintOptions.view`) and paint a large group over several budgets. A text
// that overflows its slide by thousands of lines paints only the lines that can show; a group of
// thousands of shapes paints over several frames. What can show paints exactly as before.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import {
  buildPptx,
  para,
  run,
  shape,
  table,
  cell,
} from '../../../../../scripts/fixtures/builder.ts';
import { openPresentation } from '../../editor/document-session.ts';
import type { SlideLayout } from '../../layout/index.ts';
import { paintSlide, paintWeight, type PaintOptions, type ViewRect } from '../paint.ts';
import { SlidePainter } from '../slide-painter.ts';

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const IN = 914_400;
const NOFIT = '<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>';
const OPTS: PaintOptions = { images: () => undefined, idPrefix: 'x', hitTargets: false };

async function layoutOf(shapes: string): Promise<SlideLayout> {
  const opened = await openPresentation(buildPptx({ slides: [{ shapes }] }), {
    fixedMeasurer: true,
    substitutes: false,
  });
  if (!opened.ok) throw new Error('deck');
  return opened.session.layout(0)!;
}

const slideView = (lay: SlideLayout): ViewRect => ({
  x: 0,
  y: 0,
  width: lay.slide.width,
  height: lay.slide.height,
});

/** A text box at the top of the slide with `n` paragraphs, far more than the slide holds. */
const longText = (n: number, effects?: string) =>
  shape({
    id: 2,
    x: IN,
    y: IN,
    cx: 4 * IN,
    cy: IN,
    bodyPr: NOFIT,
    ...(effects ? { effects } : {}),
    text: Array.from({ length: n }, (_, i) => para(run(`line ${i}`))).join(''),
  });

const texts = (svg: Element) => [...svg.querySelectorAll('text')].map((t) => t.textContent);

describe('a view', () => {
  test('leaves out the lines of an overflowing text that cannot show', async () => {
    const lay = await layoutOf(longText(2000));
    const all = texts(paintSlide(doc, lay, OPTS));
    const shown = texts(paintSlide(doc, lay, { ...OPTS, view: slideView(lay) }));
    expect(all.length).toBe(2000);
    // The lines in view, and a margin of a line or two: the same lines, in the same order.
    expect(shown.length).toBeGreaterThan(10);
    expect(shown.length).toBeLessThan(60);
    expect(shown).toEqual(all.slice(0, shown.length));
  });

  test('paints a slide whose text fits exactly as a paint without a view', async () => {
    const lay = await layoutOf(longText(5));
    const a = paintSlide(doc, lay, OPTS).outerHTML;
    const b = paintSlide(doc, lay, { ...OPTS, view: slideView(lay) }).outerHTML;
    expect(b).toBe(a);
  });

  test('paints every line of a shape whose shadow copies its text', async () => {
    const shadow =
      '<a:effectLst><a:outerShdw dist="38100" dir="2700000"><a:srgbClr val="000000"/>' +
      '</a:outerShdw></a:effectLst>';
    const lay = await layoutOf(longText(300, shadow));
    expect(texts(paintSlide(doc, lay, { ...OPTS, view: slideView(lay) })).length).toBe(300);
  });

  test('leaves out the table cells that cannot show', async () => {
    const rows = Array.from(
      { length: 300 },
      (_, r) =>
        `<a:tr h="${IN / 4}">${Array.from({ length: 4 }, (_, c) => cell(`r${r}c${c}`)).join('')}</a:tr>`
    ).join('');
    const lay = await layoutOf(
      table({
        id: 2,
        x: IN,
        y: IN,
        cx: 6 * IN,
        cy: 75 * IN,
        columns: [1.5 * IN, 1.5 * IN, 1.5 * IN, 1.5 * IN].map(Math.round),
        tblPr: '<a:tblPr/>',
        rows,
      })
    );
    const all = texts(paintSlide(doc, lay, OPTS));
    const shown = texts(paintSlide(doc, lay, { ...OPTS, view: slideView(lay) }));
    expect(all.length).toBe(1200);
    expect(shown.length).toBeLessThan(200);
    expect(shown).toEqual(all.slice(0, shown.length));
  });

  test('a painter whose view grows paints the lines that came into view', async () => {
    const lay = await layoutOf(longText(2000) + shape({ id: 3, x: 0, y: 0, cx: IN, cy: IN }));
    const painter = new SlidePainter(doc, { ...OPTS, view: slideView(lay) });
    const svg = painter.paint(lay);
    const plain = svg.querySelector('[data-kind="content"]')!.children[1];
    const before = texts(svg).length;
    const taller = { ...slideView(lay), height: lay.slide.height * 4 };
    expect(painter.view(taller)).toBe(true);
    expect(painter.view(taller)).toBe(false);
    painter.paint(lay);
    expect(texts(svg).length).toBeGreaterThan(before * 3);
    // A shape without text keeps its node: only text and tables depend on the view.
    expect(svg.querySelector('[data-kind="content"]')!.children[1]).toBe(plain!);
  });
});

/** A group of `n` small shapes. */
const group = (n: number) =>
  '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="1000" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${8 * IN}" cy="${6 * IN}"/>` +
  `<a:chOff x="0" y="0"/><a:chExt cx="${8 * IN}" cy="${6 * IN}"/></a:xfrm></p:grpSpPr>` +
  Array.from({ length: n }, (_, i) =>
    shape({
      id: i + 2,
      x: (i % 10) * 0.8 * IN,
      y: Math.floor(i / 10) * 0.5 * IN,
      cx: IN / 2,
      cy: IN / 4,
    })
  ).join('') +
  '</p:grpSp>';

describe('a paint budget', () => {
  test('counts the shapes of a group, and paints a large group over several paints', async () => {
    const lay = await layoutOf(group(100));
    expect(paintWeight(lay.items[0]!)).toBe(101);
    const painter = new SlidePainter(doc, OPTS);
    const shapes = (svg: Element) => svg.querySelectorAll('[data-kind="shape"]').length;
    let svg = painter.paint(lay, 30);
    expect(painter.pending).toBe(true);
    expect(shapes(svg)).toBe(30);
    const counts: number[] = [];
    while (painter.pending) {
      svg = painter.paint(lay, 30);
      counts.push(shapes(svg));
    }
    expect(counts).toEqual([60, 90, 100]);
    expect(svg.outerHTML).toBe(paintSlide(doc, lay, OPTS).outerHTML);
  });

  test('a group that paints again keeps its earlier node until the new one is whole', async () => {
    const lay = await layoutOf(group(100));
    const painter = new SlidePainter(doc, OPTS);
    const svg = painter.paint(lay);
    const old = svg.querySelector('[data-kind="group"]');
    const whole = svg.outerHTML;
    painter.invalidate('new faces');
    painter.paint(lay, 30);
    expect(painter.pending).toBe(true);
    // Still the earlier group, whole.
    expect(svg.querySelector('[data-kind="group"]')).toBe(old);
    expect(svg.outerHTML).toBe(whole);
    while (painter.pending) painter.paint(lay, 30);
    expect(svg.querySelector('[data-kind="group"]')).not.toBe(old);
    expect(svg.querySelectorAll('[data-kind="group"]').length).toBe(1);
    expect(svg.outerHTML).toBe(whole);
  });
});
