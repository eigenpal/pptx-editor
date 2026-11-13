// An edit paints the shapes it changed: the cascade, the layout, and the painter keep the objects
// of unchanged shapes, and the painted slide equals a fresh paint of the same layout.

import { afterAll, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, group, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { createPptxEditor } from '../../editor/index.ts';
import { resolveSlide } from '../../cascade/index.ts';
import {
  createFixedMeasurer,
  layoutSlide,
  type LayoutItem,
  type SlideLayout,
} from '../../layout/index.ts';
import { readPptx } from '../../store/index.ts';
import { paintSlide, paintWeight, type PaintOptions } from '../paint.ts';
import { MAX_SLIDE_TILE_IMAGES } from '../paint-budgets.ts';
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

test('a tiled shape added in front leaves the kept shapes as a fresh paint draws them', () => {
  const box = (id: number): string =>
    shape({ id, x: id * 400_000, y: 0, cx: 9_525 * 40, cy: 9_525 * 40, fill: TILE(781, 1041) });
  const layout = tiledLayout(box(2) + box(3) + box(4));
  const painter = new SlidePainter(doc, TILE_OPTS);
  painter.paint(layout);
  // The kept boxes paint after the new one: the second of them no longer fits the budget.
  const added = { ...layout, items: [tiledLayout(box(5)).items[0]!, ...layout.items] };
  const svg = painter.paint(added);
  const fresh = paintSlide(doc, added, TILE_OPTS);
  expect(fresh.querySelectorAll('image').length).toBe(2 * 1_600 + 2);
  expect(svg.querySelectorAll('image').length).toBe(2 * 1_600 + 2);
  expect(normalized(svg)).toBe(normalized(fresh));
  // Taking the new shape out gives the kept boxes their tiles back.
  painter.paint(layout);
  expect(normalized(svg)).toBe(normalized(paintSlide(doc, layout, TILE_OPTS)));
});

const SHAPE_GRADIENT =
  '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>' +
  '<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst>' +
  '<a:path path="shape"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>' +
  '</a:gradFill>';

test('each paint of a slide has the whole band budget: a shape gradient keeps its bands', () => {
  const layout = tiledLayout(
    shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: 2 * IN, prst: 'star5', fill: SHAPE_GRADIENT })
  );
  const painter = new SlidePainter(doc, TILE_OPTS);
  let svg = painter.paint(layout);
  // Each new key paints the shape again: 400 paints spend the band budget of each paint, not
  // of the painter.
  for (let i = 0; i < 400; i += 1) {
    painter.invalidate(i);
    svg = painter.paint(layout);
  }
  expect(normalized(svg)).toBe(normalized(paintSlide(doc, layout, TILE_OPTS)));
});

/**
 * The painted slide as `normalized` text with its ids left out: the rectangle construction of a
 * shape gradient refers to its gradients from its definitions, whose order an incremental paint
 * changes.
 */
const unnumbered = (svg: Element): string =>
  normalized(svg).replace(/id\d+/g, 'id').split('\n').sort().join('\n');

/** A shape with a shape gradient on a custom outline of `count` commands, at `id` × 300,000. */
function customBanded(id: number, count: number): string {
  const points = Array.from({ length: count - 1 }, (_, i) => {
    const t = (i / (count - 1)) * 2 * Math.PI;
    const x = Math.round(50_000 + 50_000 * Math.cos(t));
    return `<a:lnTo><a:pt x="${x}" y="${Math.round(50_000 + 50_000 * Math.sin(t))}"/></a:lnTo>`;
  }).join('');
  const custom =
    '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>' +
    '<a:pathLst><a:path w="100000" h="100000"><a:moveTo><a:pt x="100000" y="50000"/></a:moveTo>' +
    `${points}</a:path></a:pathLst></a:custGeom>`;
  return shape({ id, x: id * 300_000, y: 0, cx: IN, cy: IN, fill: SHAPE_GRADIENT }).replace(
    /<a:prstGeom[\s\S]*?<\/a:prstGeom>/,
    custom
  );
}

test('a banded shape added in front leaves the kept shapes as a fresh paint draws them', () => {
  // Each 1,500-command outline costs 192,000 of the 512,000 band work: two fit, the rest not.
  const banded = (id: number): string => customBanded(id, 1_500);
  const layout = tiledLayout([2, 3, 4, 5, 6].map(banded).join(''));
  const painter = new SlidePainter(doc, TILE_OPTS);
  painter.paint(layout);
  const added = { ...layout, items: [tiledLayout(banded(7)).items[0]!, ...layout.items] };
  const svg = painter.paint(added);
  expect(unnumbered(svg)).toBe(unnumbered(paintSlide(doc, added, TILE_OPTS)));
  painter.paint(layout);
  expect(unnumbered(svg)).toBe(unnumbered(paintSlide(doc, layout, TILE_OPTS)));
});

/** A group at `x` of `count` plain shapes, with `first` as its first child and `last` after. */
const heavyGroup = (id: number, x: number, count: number, first: string, last = ''): string =>
  group(
    { id, x, y: 2 * IN, cx: IN, cy: IN },
    first +
      Array.from({ length: count - 1 }, (_, i) =>
        shape({ id: id + 1 + i, x, y: 2 * IN, cx: 9_525, cy: 9_525 })
      ).join('') +
      last
  );

/** Paint `layout` a frame at a time with `budget`: the frames it takes, at most `max`. */
function frames(painter: SlidePainter, layout: SlideLayout, budget: number, max = 40): number {
  for (let n = 1; n <= max; n += 1) {
    painter.paint(layout, budget);
    if (!painter.pending) return n;
  }
  return Infinity;
}

test('a heavy group that paints again for its band budget finishes over frames', () => {
  const star = (id: number, x: number) =>
    shape({ id, x, y: 0, cx: IN, cy: IN, prst: 'star5', fill: SHAPE_GRADIENT });
  const layout = tiledLayout(heavyGroup(10, 2 * IN, 2_100, star(5_000, 2 * IN)));
  const painter = new SlidePainter(doc, TILE_OPTS);
  const svg = painter.paint(layout);
  // A shape gradient added behind the group leaves the group's star less band work.
  const added = { ...layout, items: [tiledLayout(star(2, 0)).items[0]!, ...layout.items] };
  expect(frames(painter, added, 2_000)).toBeLessThan(10);
  expect(unnumbered(svg)).toBe(unnumbered(paintSlide(doc, added, TILE_OPTS)));
});

test('heavy groups that paint again for their tile budget finish as a fresh paint', () => {
  const box = (id: number, x: number) =>
    shape({ id, x, y: 0, cx: 9_525 * 40, cy: 9_525 * 40, fill: TILE(781, 1041) });
  const layout = tiledLayout(
    heavyGroup(10, 0, 2_100, box(5_000, 0)) + heavyGroup(6_000, 2 * IN, 2_100, box(9_000, 2 * IN))
  );
  const painter = new SlidePainter(doc, TILE_OPTS);
  const svg = painter.paint(layout);
  expect(svg.querySelectorAll('image').length).toBe(2 * 1_600);
  const added = { ...layout, items: [tiledLayout(box(2, 4 * IN)).items[0]!, ...layout.items] };
  expect(frames(painter, added, 2_000)).toBeLessThan(10);
  expect(svg.querySelectorAll('image').length).toBe(2 * 1_600 + 1);
  expect(unnumbered(svg)).toBe(unnumbered(paintSlide(doc, added, TILE_OPTS)));
});

test('a kept item paints again only for the budgets it read', () => {
  const box = shape({ id: 2, x: 0, y: 0, cx: 9_525 * 40, cy: 9_525 * 40, fill: TILE(781, 1041) });
  const star = (prst: string) =>
    shape({ id: 3, x: 2 * IN, y: 0, cx: IN, cy: IN, prst, fill: SHAPE_GRADIENT });
  for (const prst of ['star5', 'star32']) {
    const layout = tiledLayout(star(prst));
    const painter = new SlidePainter(doc, TILE_OPTS);
    const svg = painter.paint(layout);
    const node = () => svg.querySelector('[data-node-id]:last-child');
    const before = node();
    // A tiled shape in front spends tiles, not band work: the star keeps its node.
    painter.paint({ ...layout, items: [tiledLayout(box).items[0]!, ...layout.items] });
    expect(node()).toBe(before);
  }
  // A banded shape in front spends band work: a star past the limits of one shape (more than
  // 2,000 commands) draws the rectangle construction whatever is left, and keeps its node.
  const custom = tiledLayout(customBanded(3, 2_400)).items[0]!;
  const layout = { ...tiledLayout(''), items: [custom] };
  const painter = new SlidePainter(doc, TILE_OPTS);
  const svg = painter.paint(layout);
  const before = svg.querySelector('[data-node-id]')!;
  painter.paint({ ...layout, items: [tiledLayout(star('star5')).items[0]!, custom] });
  expect(svg.querySelectorAll('[data-node-id]')[1]).toBe(before);
});

/** A tiled box of 40 × 40 tiles at `x`. */
const tiledBox = (id: number, x: number) =>
  shape({ id, x, y: 0, cx: 9_525 * 40, cy: 9_525 * 40, fill: TILE(781, 1041) });
/** A star with a shape gradient at `x`. */
const bandedStar = (id: number, x: number) =>
  shape({ id, x, y: 0, cx: IN, cy: IN, prst: 'star5', fill: SHAPE_GRADIENT });

/** Paint `layout` a frame at a time until nothing is pending, and expect a fresh paint. */
function settles(painter: SlidePainter, layout: SlideLayout, budget: number, max: number) {
  let svg = painter.paint(layout, budget);
  for (let n = 1; painter.pending && n < max; n += 1) svg = painter.paint(layout, budget);
  expect(painter.pending).toBe(false);
  expect(unnumbered(svg)).toBe(unnumbered(paintSlide(doc, layout, TILE_OPTS)));
  return svg;
}

test('a group painted in part paints again when the items before it change', () => {
  const all = tiledLayout(
    tiledBox(2, 0) + tiledBox(3, IN) + heavyGroup(10, 0, 2_100, tiledBox(5_000, 0))
  );
  const [box1, box2, g] = all.items;
  const painter = new SlidePainter(doc, TILE_OPTS);
  // The group's box finds 896 tiles left after the two boxes: it draws the pattern.
  painter.paint({ ...all, items: [box1!, box2!, g!] }, 2_000);
  expect(painter.pending).toBe(true);
  // Alone, it has every tile: its earlier children paint again with them.
  const svg = settles(painter, { ...all, items: [g!] }, 2_000, 10);
  expect(svg.querySelectorAll('image').length).toBe(1_600);
});

test('a group painted in part keeps what was left of a budget its later children read', () => {
  // The group's first frame reads only tiles; its last child, a banded shape, reads bands in a
  // later frame, after two banded shapes in front leave it too little band work.
  const all = tiledLayout(
    customBanded(2, 1_500) +
      customBanded(3, 1_500) +
      heavyGroup(10, 3 * IN, 2_100, tiledBox(5_000, 3 * IN), customBanded(7_000, 1_500))
  );
  const [b1, b2, g] = all.items;
  const alone = { ...all, items: [g!] };
  const painter = new SlidePainter(doc, TILE_OPTS);
  painter.paint(alone, 2_000);
  settles(painter, { ...all, items: [b1!, b2!, g!] }, 2_000, 40);
  // Alone again, the last child has its full band work: a fresh paint draws its bands.
  settles(painter, alone, Infinity, 40);
});

test('an item twice in one layout lets the items after it paint and settle', () => {
  const layout = tiledLayout(bandedStar(2, 0) + bandedStar(3, IN));
  const [x, y] = layout.items;
  const twice = { ...layout, items: [x!, x!, y!] };
  settles(new SlidePainter(doc, TILE_OPTS), twice, 1, 10);
});

test('a group that paints again in part goes back to its node when the reason goes', () => {
  const all = tiledLayout(
    tiledBox(2, 0) + tiledBox(3, IN) + heavyGroup(10, 0, 2_100, tiledBox(5_000, 0))
  );
  const [box1, box2, g] = all.items;
  const painter = new SlidePainter(doc, TILE_OPTS);
  const svg = painter.paint({ ...all, items: [g!] });
  const node = svg.querySelector('[data-kind="group"]');
  painter.paint({ ...all, items: [box1!, box2!, g!] }, 2_000);
  expect(painter.pending).toBe(true);
  settles(painter, { ...all, items: [g!] }, 2_000, 10);
  // Its first paint fits again: it stays, and later paints keep it.
  expect(svg.querySelector('[data-kind="group"]')).toBe(node);
  painter.paint({ ...all, items: [g!] }, 2_000);
  expect(painter.pending).toBe(false);
  expect(svg.querySelector('[data-kind="group"]')).toBe(node);
});

test('a group that painted again for its budgets is kept by later paints', () => {
  const all = tiledLayout(bandedStar(2, 0) + heavyGroup(10, 0, 2_100, bandedStar(5_000, 0)));
  const [star, g] = all.items;
  const painter = new SlidePainter(doc, TILE_OPTS);
  painter.paint({ ...all, items: [star!, g!] }, 2_000);
  const svg = settles(painter, { ...all, items: [g!] }, 2_000, 10);
  const node = svg.querySelector('[data-kind="group"]');
  for (let i = 0; i < 3; i += 1) {
    painter.paint({ ...all, items: [g!] }, 2_000);
    expect(painter.pending).toBe(false);
    expect(svg.querySelector('[data-kind="group"]')).toBe(node);
  }
});

/** A seeded linear congruential generator in [0, 1). */
function seeded(seed: number): () => number {
  let s = seed >>> 0;
  return () => {
    s = (s * 1_664_525 + 1_013_904_223) >>> 0;
    return s / 2 ** 32;
  };
}

test('edits painted over frames of any budget settle to a fresh paint, and stay', () => {
  const pool = tiledLayout(
    tiledBox(2, 0) +
      tiledBox(3, IN) +
      tiledBox(4, 2 * IN) +
      customBanded(20, 1_500) +
      customBanded(21, 1_500) +
      bandedStar(23, 4 * IN) +
      heavyGroup(100, 0, 60, tiledBox(99, 0)) +
      heavyGroup(300, 2 * IN, 60, bandedStar(299, 0)) +
      heavyGroup(600, 3 * IN, 2_100, tiledBox(599, 3 * IN))
  ).items;
  const base = tiledLayout('');
  const budgets = [1, 3, 2_000, Infinity];
  for (let trial = 0; trial < 6; trial += 1) {
    const random = seeded(trial * 7_919 + 1);
    const pick = <T>(list: readonly T[]): T => list[Math.floor(random() * list.length)]!;
    const painter = new SlidePainter(doc, TILE_OPTS);
    let items = pool.filter(() => random() < 0.4);
    painter.paint({ ...base, items }, pick(budgets));
    for (let step = 0; step < 6; step += 1) {
      const op = random();
      const out = pool.filter((i) => !items.includes(i));
      if (op < 0.4 && out.length > 0) {
        const at = Math.floor(random() * (items.length + 1));
        items = [...items.slice(0, at), pick(out), ...items.slice(at)];
      } else if (op < 0.7 && items.length > 0) {
        const at = Math.floor(random() * items.length);
        items = items.filter((_, i) => i !== at);
      } else items = [...items].sort(() => random() - 0.5);
      const count = 1 + Math.floor(random() * 2);
      for (let k = 0; k < count; k += 1) painter.paint({ ...base, items }, pick(budgets));
    }
    const layout = { ...base, items };
    const budget = pick(budgets);
    const weight = items.reduce((n, i) => n + paintWeight(i), 0);
    // Each item paints at most twice: once for the edits, once more for the budgets it read.
    const max = 2 * Math.ceil(weight / budget) + 2 * items.length + 4;
    const svg = settles(painter, layout, budget, max);
    const before = [...svg.querySelectorAll('[data-kind="content"] > *')];
    for (const b of budgets) {
      painter.paint(layout, b);
      expect(painter.pending).toBe(false);
      expect([...svg.querySelectorAll('[data-kind="content"] > *')]).toEqual(before);
    }
  }
});
