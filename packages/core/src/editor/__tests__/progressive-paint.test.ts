// A slide with more items than one frame paints goes on screen over several frames: the first
// frame shows the bottom items, and the rest follow until the slide equals a full paint. The
// editor stays usable in between.

import { afterAll, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { createPptxEditor } from '../index.ts';
import { surfaceView } from '../surface-paint.ts';

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const IN = 914_400;
const SHAPES = 4_500;

const deck = () =>
  buildPptx({
    slides: [
      {
        shapes: Array.from({ length: SHAPES }, (_, i) =>
          shape({
            id: i + 2,
            x: (i % 90) * IN * 0.1,
            y: Math.floor(i / 90) * IN * 0.1,
            cx: IN / 20,
            cy: IN / 20,
          })
        ).join(''),
      },
    ],
  });

const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));

test('a slide of thousands of shapes paints over several frames, and then in full', async () => {
  const g = globalThis as { ResizeObserver?: unknown };
  g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false, zoom: 1 });
  expect((await editor.load(deck())).ok).toBe(true);
  const host = doc.createElement('div') as unknown as HTMLElement;
  doc.body.append(host as unknown as Node);
  editor.attach(host);
  const painted = () => host.querySelectorAll('.pptx-surface__slide [data-node-id]').length;
  const ids = editor.session!.layout(0)!.items.map((i) => i.element.nodeId);
  expect(ids.length).toBe(SHAPES);
  // The first frame paints part of the slide, from the bottom of the paint order.
  const first = painted();
  expect(first).toBeGreaterThan(0);
  expect(first).toBeLessThan(SHAPES);
  // A selection in between keeps the painted shapes and the frames that follow.
  editor.setSelection({ kind: 'shapes', nodeIds: [ids[0]!] });
  for (let i = 0; i < 50 && painted() < SHAPES; i += 1) await wait(20);
  expect(painted()).toBe(SHAPES);
  const order = [...host.querySelectorAll('.pptx-surface__slide [data-node-id]')].map((n) =>
    n.getAttribute('data-node-id')
  );
  expect(order).toEqual(ids);
  editor.destroy();
});

/** A slide whose only item is a group of `SHAPES` shapes. */
const groupDeck = () =>
  buildPptx({
    slides: [
      {
        shapes:
          '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="100000" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
          `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${9 * IN}" cy="${5 * IN}"/>` +
          `<a:chOff x="0" y="0"/><a:chExt cx="${9 * IN}" cy="${5 * IN}"/></a:xfrm></p:grpSpPr>` +
          Array.from({ length: SHAPES }, (_, i) =>
            shape({
              id: i + 2,
              x: (i % 90) * IN * 0.1,
              y: Math.floor(i / 90) * IN * 0.1,
              cx: IN / 20,
              cy: IN / 20,
            })
          ).join('') +
          '</p:grpSp>',
      },
    ],
  });

test('a group of thousands of shapes paints over several frames, and its thumbnail when idle', async () => {
  const g = globalThis as { ResizeObserver?: unknown };
  g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false, zoom: 1 });
  expect((await editor.load(groupDeck())).ok).toBe(true);
  const host = doc.createElement('div') as unknown as HTMLElement;
  doc.body.append(host as unknown as Node);
  editor.attach(host);
  const shapes = (root: ParentNode) => root.querySelectorAll('[data-kind="shape"]').length;
  const slide = () => host.querySelector('.pptx-surface__slide')!;
  // One item, but the first frame paints only part of its shapes.
  expect(editor.session!.layout(0)!.items.length).toBe(1);
  const first = shapes(slide());
  expect(first).toBeGreaterThan(0);
  expect(first).toBeLessThan(SHAPES);
  // The current slide's thumbnail is large: it waits for idle time, then fills in.
  const thumb = doc.createElement('div') as unknown as HTMLElement;
  editor.renderThumbnail(0, thumb, { defer: true });
  expect(thumb.firstChild).toBeNull();
  for (let i = 0; i < 100 && (shapes(slide()) < SHAPES || shapes(thumb) < SHAPES); i += 1)
    await wait(20);
  expect(shapes(slide())).toBe(SHAPES);
  expect(shapes(thumb)).toBe(SHAPES);
  editor.destroy();
});

test('the surface paints the lines of an overflowing text that its window can show', async () => {
  const g = globalThis as { ResizeObserver?: unknown };
  g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
  const lines = 3000;
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false, zoom: 1 });
  const text = Array.from({ length: lines }, (_, i) => para(run(`line ${i}`))).join('');
  const bytes = buildPptx({
    slides: [
      {
        shapes: shape({
          id: 2,
          x: IN,
          y: IN,
          cx: 4 * IN,
          cy: IN,
          bodyPr: '<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>',
          text,
        }),
      },
    ],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  const host = doc.createElement('div') as unknown as HTMLElement;
  doc.body.append(host as unknown as Node);
  editor.attach(host);
  const painted = () => host.querySelectorAll('.pptx-surface__slide text').length;
  const near = painted();
  expect(near).toBeGreaterThan(20);
  expect(near).toBeLessThan(lines / 4);
  // Zoomed far out, the window reaches further past the slide: more lines paint.
  editor.setZoom(0.1);
  for (let i = 0; i < 20 && painted() === near; i += 1) await wait(20);
  expect(painted()).toBeGreaterThan(near);
  // The thumbnail clips to the slide: fewer lines than the surface.
  const thumb = doc.createElement('div') as unknown as HTMLElement;
  editor.renderThumbnail(0, thumb);
  expect(thumb.querySelectorAll('text').length).toBeLessThan(near);
  expect(thumb.querySelectorAll('text').length).toBeGreaterThan(5);
  editor.destroy();
});

test('the surface view reaches as far as the window at the zoom, by powers of two', () => {
  const slide = { width: 12_192_000, height: 6_858_000 };
  const px = 9_525;
  // A 1440 px window at 100% reaches 13.7 M EMU: two slide widths past each edge.
  const fit = surfaceView(slide, 1440, 1, px);
  expect(fit).toEqual({
    x: -2 * slide.width,
    y: -2 * slide.width,
    width: 5 * slide.width,
    height: slide.height + 4 * slide.width,
  });
  // A zoom step changes nothing until the reach passes the next power of two.
  expect(surfaceView(slide, 1440, 0.9, px)).toEqual(fit);
  expect(surfaceView(slide, 1440, 0.5, px).x).toBe(-4 * slide.width);
  expect(surfaceView(slide, 1440, 0.1, px).x).toBe(-16 * slide.width);
  // No window (a server, a test) reaches one slide width.
  expect(surfaceView(slide, 0, 1, px).x).toBe(-slide.width);
});
