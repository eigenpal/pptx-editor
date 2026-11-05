// Text that overflows its shape: a press on an overflowing line selects the shape and places the
// caret at the pressed character, a drag selects across the overflowing lines, a double-click
// takes the word, and the pointer shows the text cursor there. Bottom, top, and center anchors,
// a rotated box, and a line that runs past the right edge when wrap is off.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { attr, elements, find, xmlPartOf } from '../../store/index.ts';
import type { LayoutItem } from '../../layout/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { caretAt } from '../surface-geometry.ts';
import { hitItem } from '../surface-hit.ts';

const IN = 914_400;
const body = (anchor: string, wrap = 'square') =>
  `<a:bodyPr wrap="${wrap}" anchor="${anchor}"><a:noAutofit/></a:bodyPr>`;
const LINES = para(run('Alpha beta', 'sz="2400"')) + para(run('Gamma delta', 'sz="2400"'));

/**
 * A full-slide shape under four text boxes whose text overflows: bottom-anchored (up), top-
 * anchored (down), centered (both ways), rotated 180 degrees and bottom-anchored (down on the
 * slide), and one long unwrapped line (right). Each frame is a fifth of an inch tall.
 */
function deck(): Uint8Array {
  const box = (id: number, name: string, x: number, y: number, bodyPr: string, extra = {}) =>
    shape({ id, name, x, y, cx: 3 * IN, cy: IN / 5, bodyPr, text: LINES, ...extra });
  return buildPptx({
    slides: [
      {
        shapes:
          shape({ id: 2, name: 'Under', x: 0, y: 0, cx: 10 * IN, cy: 7 * IN, fill: 'FF0000' }) +
          box(3, 'Bottom', IN / 2, 2 * IN, body('b')) +
          box(4, 'Top', 4 * IN, IN, body('t')) +
          box(5, 'Center', IN / 2, 4 * IN, body('ctr')) +
          box(6, 'Turned', 4 * IN, 4 * IN, body('b'), { rot: 180 * 60_000 }) +
          shape({
            id: 7,
            name: 'Wide',
            x: IN / 2,
            y: 6 * IN,
            cx: IN / 2,
            cy: IN / 2,
            bodyPr: body('t', 'none'),
            text: para(run('Unwrapped words run on', 'sz="2400"')),
          }),
      },
    ],
  });
}

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const caret = (paragraph: number, offset: number) => ({ paragraph, offset });

async function attached() {
  const g = globalThis as { ResizeObserver?: unknown };
  g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false, zoom: 1 });
  expect((await e.load(deck())).ok).toBe(true);
  const host = doc.createElement('div') as unknown as HTMLElement;
  doc.body.append(host as unknown as Node);
  e.attach(host);
  const stage = host.querySelector('.pptx-surface__stage') as HTMLElement;
  stage.getBoundingClientRect = () =>
    ({ left: 0, top: 0, right: 1280, bottom: 720, width: 1280, height: 720 }) as DOMRect;
  const width = e.session!.layout(0)!.slide.width;
  const px = (emu: number) => (emu / width) * 1280;
  const mouse = (type: string, x: number, y: number) =>
    new win.MouseEvent(type, {
      clientX: px(x),
      clientY: px(y),
      button: 0,
      bubbles: true,
      cancelable: true,
    }) as unknown as Event;
  const press = (x: number, y: number) => stage.dispatchEvent(mouse('pointerdown', x, y));
  const release = () =>
    doc.dispatchEvent(new win.MouseEvent('pointerup', { bubbles: true }) as unknown as Event);
  const move = (x: number, y: number) => doc.dispatchEvent(mouse('pointermove', x, y));
  const hover = (x: number, y: number) => stage.dispatchEvent(mouse('pointermove', x, y));
  return { e, stage, press, release, move, hover };
}

const idOf = (e: PptxEditorInstance, name: string): string => {
  const root = xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!.root;
  return elements(find(root, 'p:spTree')!).find((c) => attr(find(c, 'p:cNvPr'), 'name') === name)!
    .id;
};
const itemOf = (e: PptxEditorInstance, name: string): LayoutItem =>
  e.session!.layout(0)!.items.find((i) => i.element.nodeId === idOf(e, name))!;

/** A slide point a quarter of the way into the character after `offset` of `paragraph`. */
function pointIn(item: LayoutItem, paragraph: number, offset: number) {
  const a = caretAt(item, caret(paragraph, offset))!;
  const b = caretAt(item, caret(paragraph, offset + 1))!;
  return {
    x: (a.x1 * 3 + b.x1 + a.x2 * 3 + b.x2) / 8,
    y: (a.y1 * 3 + b.y1 + a.y2 * 3 + b.y2) / 8,
  };
}

/** Whether a slide point is outside an item's frame. */
function outside(item: LayoutItem, p: { x: number; y: number }): boolean {
  const [a, b, c, d, e, f] = item.matrix;
  const det = a * d - b * c;
  const lx = (d * (p.x - e) - c * (p.y - f)) / det;
  const ly = (-b * (p.x - e) + a * (p.y - f)) / det;
  const { cx, cy } = item.element.xfrm;
  return lx < 0 || lx > cx || ly < 0 || ly > cy;
}

describe('overflowing text', () => {
  for (const [name, paragraph] of [
    ['Bottom', 0],
    ['Top', 1],
    ['Center', 0],
    ['Center', 1],
    ['Turned', 0],
  ] as const)
    test(`a press on an overflowing line of "${name}" places the caret at that character`, async () => {
      const { e, press, release } = await attached();
      const item = itemOf(e, name);
      const p = pointIn(item, paragraph, 2);
      expect(outside(item, p)).toBe(true);
      expect(hitItem(e.session!.layout(0), p, 1)).toBe(idOf(e, name));
      press(p.x, p.y);
      release();
      expect(e.snapshot().selection).toEqual({
        kind: 'text',
        nodeId: idOf(e, name),
        anchor: caret(paragraph, 2),
        focus: caret(paragraph, 2),
      });
      e.destroy();
    });

  test('a line that runs past the right edge with wrap off takes the press', async () => {
    const { e, press, release } = await attached();
    const item = itemOf(e, 'Wide');
    const p = pointIn(item, 0, 15);
    expect(outside(item, p)).toBe(true);
    press(p.x, p.y);
    release();
    expect(e.snapshot().selection).toEqual({
      kind: 'text',
      nodeId: idOf(e, 'Wide'),
      anchor: caret(0, 15),
      focus: caret(0, 15),
    });
    e.destroy();
  });

  test('a drag selects across the overflowing lines', async () => {
    const { e, press, release, move } = await attached();
    const item = itemOf(e, 'Bottom');
    const from = pointIn(item, 0, 1);
    const to = pointIn(item, 1, 4);
    press(from.x, from.y);
    move(to.x, to.y);
    release();
    expect(e.snapshot().selection).toEqual({
      kind: 'text',
      nodeId: idOf(e, 'Bottom'),
      anchor: caret(0, 1),
      focus: caret(1, 4),
    });
    e.destroy();
  });

  test('a double-click on an overflowing line selects its word', async () => {
    const { e, press, release } = await attached();
    const item = itemOf(e, 'Bottom');
    const p = pointIn(item, 0, 7);
    press(p.x, p.y);
    release();
    press(p.x, p.y);
    release();
    expect(e.snapshot().selection).toMatchObject({
      nodeId: idOf(e, 'Bottom'),
      anchor: caret(0, 6),
      focus: caret(0, 10),
    });
    e.destroy();
  });

  test('the pointer shows the text cursor over an overflowing line', async () => {
    // The hover reads `instanceof Node`; the engine test run registers no global DOM.
    const g = globalThis as { Node?: unknown };
    const had = 'Node' in g;
    g.Node ??= (win as unknown as { Node: unknown }).Node;
    try {
      const { e, stage, hover } = await attached();
      const p = pointIn(itemOf(e, 'Bottom'), 0, 2);
      hover(p.x, p.y);
      expect(stage.style.cursor).toBe('text');
      e.destroy();
    } finally {
      if (!had) delete g.Node;
    }
  });

  test('away from the lines, the shape under the overflow takes the press', async () => {
    const { e } = await attached();
    const item = itemOf(e, 'Bottom');
    const p = pointIn(item, 0, 2);
    // Right of the first line's end, above the frame: no text there.
    const right = { x: item.matrix[4] + 2.9 * IN, y: p.y };
    expect(hitItem(e.session!.layout(0), right, 1)).toBe(idOf(e, 'Under'));
    e.destroy();
  });
});
