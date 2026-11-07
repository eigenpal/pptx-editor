// An edit paints the shapes it changed: the cascade, the layout, and the painter keep the objects
// of unchanged shapes, and the painted slide equals a fresh paint of the same layout.

import { afterAll, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { createPptxEditor } from '../../editor/index.ts';
import type { LayoutItem } from '../../layout/index.ts';
import { paintSlide, SlidePainter, type PaintOptions } from '../paint.ts';

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const IN = 914_400;
const GRADIENT =
  '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>' +
  '<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill>';

const deck = () =>
  buildPptx({
    slides: [
      {
        shapes: Array.from({ length: 12 }, (_, i) =>
          shape({
            id: i + 2,
            x: (i % 4) * 2 * IN,
            y: Math.floor(i / 4) * 2 * IN,
            cx: 1.8 * IN,
            cy: 1.5 * IN,
            ...(i % 3 === 0 ? { fill: GRADIENT } : {}),
            text: para(run(`Shape ${i}`)),
          })
        ).join(''),
      },
    ],
  });

/**
 * The painted slide as text, with its generated ids numbered in order of use and its
 * definitions sorted: an incremental paint adds a changed shape's definitions last.
 */
function normalized(svg: Element): string {
  const defs = svg.querySelector('defs')!;
  const used = svg.outerHTML.replace(defs.outerHTML, '');
  const ids = new Map<string, string>();
  const rename = (text: string) =>
    text.replace(/x-[a-z]+\d+/g, (id) => {
      if (!ids.has(id)) ids.set(id, `id${ids.size}`);
      return ids.get(id)!;
    });
  const body = rename(used);
  const definitions = [...defs.children].map((d) => rename(d.outerHTML)).sort();
  return `${body}\n${definitions.join('\n')}`;
}

test('typing repaints one shape, and the slide equals a fresh paint', async () => {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(deck())).ok).toBe(true);
  const opts: PaintOptions = { images: () => undefined, idPrefix: 'x' };
  const painter = new SlidePainter(doc, opts);
  const first = editor.session!.layout(0)!;
  const svg = painter.paint(first);
  const nodeOf = (id: string) => svg.querySelector(`[data-node-id="${id}"]`);
  const ids = first.items.map((i) => i.element.nodeId);
  const before = new Map(ids.map((id) => [id, nodeOf(id)]));
  const edited = ids[5]!;
  editor.setSelection({
    kind: 'text',
    nodeId: edited,
    anchor: { paragraph: 0, offset: 0 },
    focus: { paragraph: 0, offset: 0 },
  });
  for (const key of 'abc') editor.input('text', key);
  const next = editor.session!.layout(0)!;
  // The layout keeps the items of the shapes that did not change.
  const reused = next.items.filter((item, i) => item === first.items[i]).length;
  expect(reused).toBe(first.items.length - 1);
  expect(painter.paint(next)).toBe(svg);
  for (const id of ids) {
    if (id === edited) expect(nodeOf(id)).not.toBe(before.get(id) ?? null);
    else expect(nodeOf(id)).toBe(before.get(id) ?? null);
  }
  expect(svg.querySelector(`[data-node-id="${edited}"]`)!.textContent).toContain('abcShape 5');
  expect(normalized(svg)).toBe(normalized(paintSlide(doc, next, opts)));
  editor.destroy();
});

test('a removed shape leaves no node and no definitions behind', async () => {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(deck())).ok).toBe(true);
  const opts: PaintOptions = { images: () => undefined, idPrefix: 'x' };
  const painter = new SlidePainter(doc, opts);
  const first = editor.session!.layout(0)!;
  const svg = painter.paint(first);
  const defs = svg.querySelector('defs')!.children.length;
  const gone = first.items[0]!.element.nodeId;
  editor.setSelection({ kind: 'shapes', nodeIds: [gone] });
  expect(editor.exec({ type: 'delete' }).ok).toBe(true);
  const next = editor.session!.layout(0)!;
  painter.paint(next);
  expect(svg.querySelector(`[data-node-id="${gone}"]`)).toBeNull();
  expect(svg.querySelector('defs')!.children.length).toBeLessThan(defs);
  expect(normalized(svg)).toBe(normalized(paintSlide(doc, next, opts)));
  editor.destroy();
});

test('a budget paints a slide over several calls, in paint order, as one paint would', async () => {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(deck())).ok).toBe(true);
  const opts: PaintOptions = { images: () => undefined, idPrefix: 'x' };
  const painter = new SlidePainter(doc, opts);
  const layout = editor.session!.layout(0)!;
  const ids = layout.items.map((i) => i.element.nodeId);
  const shown = (svg: Element) =>
    [...svg.querySelectorAll('[data-kind="content"] > [data-node-id]')].map((n) =>
      n.getAttribute('data-node-id')
    );
  const svg = painter.paint(layout, 5);
  expect(painter.pending).toBe(true);
  expect(shown(svg)).toEqual(ids.slice(0, 5));
  // An edit while items wait: the painted items stay, and new ones still wait for the budget.
  editor.setSelection({
    kind: 'text',
    nodeId: ids[8]!,
    anchor: { paragraph: 0, offset: 0 },
    focus: { paragraph: 0, offset: 0 },
  });
  editor.input('text', 'z');
  const next = editor.session!.layout(0)!;
  const kept = new Map(
    ids.slice(0, 5).map((id) => [id, svg.querySelector(`[data-node-id="${id}"]`)])
  );
  expect(painter.paint(next, 5)).toBe(svg);
  expect(shown(svg)).toEqual(ids.slice(0, 10));
  for (const [id, node] of kept) expect(svg.querySelector(`[data-node-id="${id}"]`)).toBe(node);
  painter.paint(next, 5);
  expect(painter.pending).toBe(false);
  expect(shown(svg)).toEqual(ids);
  expect(normalized(svg)).toBe(normalized(paintSlide(doc, next, opts)));
  editor.destroy();
});

test('while a budget repaints new items, each earlier node stays in its place', async () => {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(deck())).ok).toBe(true);
  const opts: PaintOptions = { images: () => undefined, idPrefix: 'x' };
  const painter = new SlidePainter(doc, opts);
  const layout = editor.session!.layout(0)!;
  const ids = layout.items.map((i) => i.element.nodeId);
  const shown = (svg: Element) =>
    [...svg.querySelectorAll('[data-kind="content"] > [data-node-id]')].map((n) =>
      n.getAttribute('data-node-id')
    );
  const svg = painter.paint(layout);
  const before = new Map(ids.map((id) => [id, svg.querySelector(`[data-node-id="${id}"]`)]));
  // Every item new, as after a font load: the same shapes, new objects.
  const fresh = { ...layout, items: layout.items.map((i) => ({ ...i })) };
  painter.paint(fresh, 5);
  expect(painter.pending).toBe(true);
  expect(shown(svg)).toEqual(ids);
  ids.forEach((id, i) => {
    const node = svg.querySelector(`[data-node-id="${id}"]`);
    if (i < 5) expect(node).not.toBe(before.get(id)!);
    else expect(node).toBe(before.get(id)!);
  });
  painter.paint(fresh, 5);
  painter.paint(fresh, 5);
  expect(painter.pending).toBe(false);
  expect(shown(svg)).toEqual(ids);
  for (const id of ids)
    expect(svg.querySelector(`[data-node-id="${id}"]`)).not.toBe(before.get(id)!);
  expect(normalized(svg)).toBe(normalized(paintSlide(doc, fresh, opts)));
  editor.destroy();
});

test('a stand-in is the earlier node of the same part and layer, not of a shared node id', async () => {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(deck())).ok).toBe(true);
  const layout = editor.session!.layout(0)!;
  const first = layout.items[0]!;
  // Node ids are numbered per part: a master shape can share the id of a slide shape.
  const master = {
    ...first,
    element: { ...first.element, partName: '/ppt/slideMasters/slideMaster1.xml' },
  } as LayoutItem;
  const items: LayoutItem[] = [master, ...layout.items];
  const painter = new SlidePainter(doc, { images: () => undefined, idPrefix: 'x' });
  const svg = painter.paint({ ...layout, items });
  const content = () => [...svg.querySelectorAll('[data-kind="content"] > *')];
  const before = content();
  expect(before.length).toBe(items.length);
  painter.paint({ ...layout, items: items.map((i) => ({ ...i })) }, 0);
  expect(content()).toEqual(before);
  editor.destroy();
});
