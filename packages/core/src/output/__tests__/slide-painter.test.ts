// An edit paints the shapes it changed: the cascade, the layout, and the painter keep the objects
// of unchanged shapes, and the painted slide equals a fresh paint of the same layout.

import { afterAll, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { createPptxEditor } from '../../editor/index.ts';
import { resolveSlide } from '../../cascade/index.ts';
import { createFixedMeasurer, layoutSlide, type LayoutItem } from '../../layout/index.ts';
import { readPptx } from '../../store/index.ts';
import { paintSlide, type PaintOptions } from '../paint.ts';
import { MAX_SLIDE_TILE_IMAGES } from '../paint-fill.ts';
import { SlidePainter } from '../slide-painter.ts';

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

test('a new key repaints the same items, each earlier node standing in until it does', async () => {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(deck())).ok).toBe(true);
  const opts: PaintOptions = { images: () => undefined, idPrefix: 'x' };
  const painter = new SlidePainter(doc, opts);
  const layout = editor.session!.layout(0)!;
  const ids = layout.items.map((i) => i.element.nodeId);
  const node = (id: string) => svg.querySelector(`[data-node-id="${id}"]`);
  expect(painter.invalidate('a')).toBe(true);
  const svg = painter.paint(layout);
  const before = new Map(ids.map((id) => [id, node(id)]));
  // The same key: nothing paints again.
  expect(painter.invalidate('a')).toBe(false);
  painter.paint(layout);
  for (const id of ids) expect(node(id)).toBe(before.get(id)!);
  // A new key (a converted picture): the same layout items paint again over several calls.
  expect(painter.invalidate('b')).toBe(true);
  painter.paint(layout, 5);
  expect(painter.pending).toBe(true);
  ids.forEach((id, i) => {
    if (i < 5) expect(node(id)).not.toBe(before.get(id)!);
    else expect(node(id)).toBe(before.get(id)!);
  });
  while (painter.pending) painter.paint(layout, 5);
  for (const id of ids) expect(node(id)).not.toBe(before.get(id)!);
  // Done: the next paint keeps every node.
  const after = new Map(ids.map((id) => [id, node(id)]));
  painter.paint(layout, 5);
  expect(painter.pending).toBe(false);
  for (const id of ids) expect(node(id)).toBe(after.get(id)!);
  expect(normalized(svg)).toBe(normalized(paintSlide(doc, layout, opts)));
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

const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const TILE = (sx: number, sy: number) =>
  `<a:blipFill><a:blip r:embed="rIdImg"/><a:tile sx="${sx}" sy="${sy}" algn="tl"/></a:blipFill>`;

/** The layout of one slide with `shapes` and an optional `background`, with a tile image. */
function tiledLayout(shapes: string, background?: string) {
  const r = readPptx(
    buildPptx({
      slides: [
        {
          layout: 'blank',
          rels: [['rIdImg', IMAGE_REL, '../media/image1.png']],
          ...(background ? { background } : {}),
          shapes,
        },
      ],
    })
  );
  if (!r.ok) throw new Error('open failed');
  return layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer());
}

const TILE_OPTS: PaintOptions = {
  images: () => 'blob:test',
  imageSize: () => ({ cx: 1_219_200, cy: 914_400 }),
  idPrefix: 'x',
};

test('each paint of a slide has the whole tile budget: a tiled background stays tiles', () => {
  const layout = tiledLayout(
    '',
    `<p:bg><p:bgPr>${TILE(100_000, 100_000)}<a:effectLst/></p:bgPr></p:bg>`
  );
  const painter = new SlidePainter(doc, TILE_OPTS);
  const tiles = () => svg.querySelectorAll('[data-kind="background"] image').length;
  const svg = painter.paint(layout);
  // 10 × 8 tiles of 1_219_200 × 914_400 cover the slide, on each paint.
  expect(tiles()).toBe(80);
  const paints = Math.ceil(MAX_SLIDE_TILE_IMAGES / 80) + 10;
  for (let i = 0; i < paints; i += 1) painter.paint(layout);
  expect(tiles()).toBe(80);
  expect([...svg.querySelectorAll('pattern')].some((p) => p.id.includes('tile'))).toBe(false);
});

test('shapes painted over several frames share one tile budget, as a fresh paint', () => {
  // Three boxes of 40 × 40 tiles: two fit the slide budget, the third draws the pattern.
  const box = (id: number): string =>
    shape({ id, x: id * 400_000, y: 0, cx: 9_525 * 40, cy: 9_525 * 40, fill: TILE(781, 1041) });
  const layout = tiledLayout(box(2) + box(3) + box(4));
  const painter = new SlidePainter(doc, TILE_OPTS);
  let svg = painter.paint(layout, 1);
  while (painter.pending) svg = painter.paint(layout, 1);
  for (let i = 0; i < 3; i += 1) painter.paint(layout);
  expect(svg.querySelectorAll('image').length).toBe(2 * 1_600 + 1);
  expect(normalized(svg)).toBe(normalized(paintSlide(doc, layout, TILE_OPTS)));
});
