import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape, solid } from '../../../../../scripts/fixtures/builder.ts';
import {
  attr,
  find,
  readPptx,
  relationshipsOf,
  xmlPartOf,
  type Element,
} from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { cornersOf, findItem } from '../../automation/commands/placement.ts';
import type { ClipboardPayload } from '../../contracts/index.ts';

const IN = 914_400;
const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0x0d, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 1,
  0, 0, 0, 1, 8, 6, 0, 0, 0, 0x1f, 0x15, 0xc4, 0x89, 0, 0, 0, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78,
  0x9c, 0x63, 0, 1, 0, 0, 5, 0, 1, 0x0d, 0x0a, 0x2d, 0xb4, 0, 0, 0, 0, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

/** A rotated, flipped, scaled group with two children, and a loose shape. */
const GROUP =
  '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group 9"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm rot="1800000" flipH="1"><a:off x="914400" y="914400"/><a:ext cx="4572000" cy="2743200"/>' +
  '<a:chOff x="0" y="0"/><a:chExt cx="2286000" cy="1371600"/></a:xfrm></p:grpSpPr>' +
  shape({ id: 11, x: 0, y: 0, cx: 1143000, cy: 685800, fill: solid('4472C4') }).replace(
    '<a:xfrm>',
    '<a:xfrm rot="600000">'
  ) +
  shape({ id: 12, x: 1143000, y: 685800, cx: 1143000, cy: 685800, fill: solid('ED7D31') }) +
  '</p:grpSp>';

const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          GROUP +
          shape({ id: 2, x: 6 * IN, y: IN, cx: 2 * IN, cy: IN, text: para(run('Loose')) }) +
          shape({ id: 3, x: 6 * IN, y: 3 * IN, cx: 2 * IN, cy: IN, prst: 'ellipse' }),
      },
      { layout: 'blank', shapes: '' },
    ],
  });

async function open(bytes = deck()): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

const part = (e: PptxEditorInstance, i = 0) => e.session!.info.slides[i].partName;
const root = (e: PptxEditorInstance, i = 0) =>
  xmlPartOf(e.session!.store.package, part(e, i))!.root;

/** The node id of the shape whose `p:cNvPr/@name` is `name`. */
function byName(e: PptxEditorInstance, name: string, i = 0): string {
  const walk = (el: Element): string | undefined => {
    for (const c of el.children) {
      if (c.type !== 'element') continue;
      if (['sp', 'grpSp', 'pic'].includes(c.kind)) {
        const nv = c.children.find((x) => x.type === 'element') as Element | undefined;
        const cNvPr = nv?.children.find((x) => x.type === 'element' && x.local === 'cNvPr') as
          Element | undefined;
        if (cNvPr && attr(cNvPr, 'name') === name) return c.id;
      }
      const hit = walk(c);
      if (hit) return hit;
    }
    return undefined;
  };
  const id = walk(root(e, i));
  if (!id) throw new Error(`no shape ${name}`);
  return id;
}

/** Slide-space corners of a shape, sorted, for position comparisons. */
function corners(e: PptxEditorInstance, id: string, slide = 0): number[][] {
  const item = findItem(e.session!.layout(slide)!.items, id)!;
  return cornersOf(item)
    .map((p) => [Math.round(p.x), Math.round(p.y)])
    .sort((a, b) => a[0] - b[0] || a[1] - b[1]);
}

function expectNear(a: number[][], b: number[][], tolerance = 3): void {
  expect(a.length).toBe(b.length);
  a.forEach((p, i) => {
    expect(Math.abs(p[0] - b[i][0])).toBeLessThanOrEqual(tolerance);
    expect(Math.abs(p[1] - b[i][1])).toBeLessThanOrEqual(tolerance);
  });
}

describe('group and ungroup', () => {
  test('ungrouping a rotated, flipped, scaled group keeps every child in place', async () => {
    const e = await open();
    const group = byName(e, 'Group 9');
    const a = byName(e, 'Shape 11');
    const b = byName(e, 'Shape 12');
    const before = [corners(e, a), corners(e, b)];
    e.setSelection({ kind: 'shapes', nodeIds: [group] });
    expect(e.exec({ type: 'ungroup' }).ok).toBe(true);
    expect(e.snapshot().selection).toEqual({ kind: 'shapes', nodeIds: [a, b] });
    expect(find(root(e), 'p:grpSp')).toBeUndefined();
    expectNear(corners(e, a), before[0]);
    expectNear(corners(e, b), before[1]);
    e.exec({ type: 'undo' });
    expect(find(root(e), 'p:grpSp')).toBeDefined();
  });

  test('grouping keeps positions and z-order; regrouping round-trips through save', async () => {
    const e = await open();
    const loose = byName(e, 'Shape 2');
    const ellipse = byName(e, 'Shape 3');
    const before = [corners(e, loose), corners(e, ellipse)];
    e.setSelection({ kind: 'shapes', nodeIds: [ellipse, loose] });
    expect(e.exec({ type: 'group' }).ok).toBe(true);
    const sel = e.snapshot().selection;
    expect(sel.kind === 'shapes' && sel.nodeIds.length).toBe(1);
    const groupId = sel.kind === 'shapes' ? sel.nodeIds[0] : '';
    const group = findItem(e.session!.layout(0)!.items, groupId)!;
    expect(group.kind).toBe('group');
    if (group.kind === 'group')
      expect(group.children.map((c) => c.element.nodeId)).toEqual([loose, ellipse]);
    expectNear(corners(e, loose), before[0]);
    expectNear(corners(e, ellipse), before[1]);
    const reopened = await open(e.save()!);
    expect(reopened.session!.layout(0)!.items.filter((i) => i.kind === 'group').length).toBe(2);
  });

  test('group needs two shapes, refuses placeholders, ungroup needs a group', async () => {
    const e = await open();
    e.setSelection({ kind: 'shapes', nodeIds: [byName(e, 'Shape 2')] });
    expect(e.can({ type: 'group' })).toEqual({ ok: false, reason: 'disabled.groupNeedsTwo' });
    expect(e.can({ type: 'ungroup' })).toEqual({ ok: false, reason: 'disabled.notAGroup' });
    e.setSelection({
      kind: 'shapes',
      nodeIds: [byName(e, 'Shape 2'), byName(e, 'Shape 11')],
    });
    expect(e.can({ type: 'group' })).toEqual({ ok: false, reason: 'disabled.groupScope' });
  });
});

describe('editing inside groups', () => {
  test('moving a child refits the group and leaves its sibling in place', async () => {
    const e = await open();
    const group = byName(e, 'Group 9');
    const a = byName(e, 'Shape 11');
    const b = byName(e, 'Shape 12');
    const sibling = corners(e, b);
    // Move child A far left in its group's child space.
    e.setSelection({ kind: 'shapes', nodeIds: [a], group });
    expect(
      e.exec({
        type: 'setTransforms',
        items: [{ nodeId: a, x: -1_000_000, y: 0, cx: 1143000, cy: 685800, rot: 10 }],
      }).ok
    ).toBe(true);
    expectNear(corners(e, b), sibling);
    const chOff = find(find(find(root(e), 'p:grpSp'), 'p:grpSpPr'), 'a:chOff')!;
    expect(Number(attr(chOff, 'x'))).toBeLessThan(0);
    // The group's box still encloses both children.
    const g = findItem(e.session!.layout(0)!.items, group)!;
    expect(g.element.xfrm.cx).toBeGreaterThan(4572000);
    expect(e.snapshot().canUndo).toBe(true);
    e.exec({ type: 'undo' });
    expect(attr(find(find(find(root(e), 'p:grpSp'), 'p:grpSpPr'), 'a:chOff')!, 'x')).toBe('0');
  });

  test('nudging several shapes is one undo step', async () => {
    const e = await open();
    const ids = [byName(e, 'Shape 2'), byName(e, 'Shape 3')];
    const before = ids.map((id) => corners(e, id));
    e.setSelection({ kind: 'shapes', nodeIds: ids });
    e.exec({ type: 'nudge', dx: 76_200, dy: 0 });
    ids.forEach((id, i) =>
      expectNear(
        corners(e, id),
        before[i].map(([x, y]) => [x + 76_200, y])
      )
    );
    e.exec({ type: 'undo' });
    ids.forEach((id, i) => expectNear(corners(e, id), before[i]));
  });

  test('deleting every child of a group removes the group', async () => {
    const e = await open();
    const group = byName(e, 'Group 9');
    e.setSelection({
      kind: 'shapes',
      nodeIds: [byName(e, 'Shape 11'), byName(e, 'Shape 12')],
      group,
    });
    expect(e.exec({ type: 'delete' }).ok).toBe(true);
    expect(find(root(e), 'p:grpSp')).toBeUndefined();
  });

  test('select all works in the current scope', async () => {
    const e = await open();
    e.selectAll();
    expect(e.snapshot().selection).toMatchObject({ kind: 'shapes' });
    const top = e.snapshot().selection;
    expect(top.kind === 'shapes' && top.nodeIds.length).toBe(3);
    const group = byName(e, 'Group 9');
    e.setSelection({ kind: 'shapes', nodeIds: [byName(e, 'Shape 11')], group });
    e.selectAll();
    const inner = e.snapshot().selection;
    expect(inner.kind === 'shapes' && inner.group).toBe(group);
    expect(inner.kind === 'shapes' && inner.nodeIds.length).toBe(2);
  });
});

describe('clipboard', () => {
  async function withPicture(): Promise<PptxEditorInstance> {
    const e = await open();
    e.exec({ type: 'insertImage', bytes: PNG, mime: 'image/png', width: 10, height: 10 });
    return e;
  }

  test('copy and paste on the same slide offsets the copy and reuses the image part', async () => {
    const e = await withPicture();
    const pic = e.snapshot().selection;
    const parts = e.session!.store.package.parts.size;
    const payload = e.copy()!;
    expect(payload.relationships.length).toBe(1);
    expect(e.exec({ type: 'paste', payload }).ok).toBe(true);
    const pasted = e.snapshot().selection;
    expect(pasted.kind === 'shapes' && pasted.nodeIds.length).toBe(1);
    expect(e.session!.store.package.parts.size).toBe(parts);
    const a = corners(e, pic.kind === 'shapes' ? pic.nodeIds[0] : '');
    const b = corners(e, pasted.kind === 'shapes' ? pasted.nodeIds[0] : '');
    expectNear(
      b,
      a.map(([x, y]) => [x + 228_600, y + 228_600])
    );
  });

  test('paste into another deck adds the media part and a new relationship', async () => {
    const src = await withPicture();
    const payload = JSON.parse(JSON.stringify(src.copy())) as ClipboardPayload;
    const dst = await open();
    dst.goToSlide(1);
    expect(dst.exec({ type: 'paste', payload }).ok).toBe(true);
    const rels = relationshipsOf(dst.session!.store.package, part(dst, 1));
    const image = rels.find((r) => r.type.endsWith('/image'));
    expect(image?.target).toMatch(/^\/ppt\/media\/image\d+\.png$/);
    const blip = find(root(dst, 1), 'a:blip')!;
    expect(blip.attrs.find((a) => a.local === 'embed')?.value).toBe(image!.id);
    expect(readPptx(dst.save()!).ok).toBe(true);
  });

  test('pasted shapes get fresh ids', async () => {
    const e = await open();
    e.setSelection({ kind: 'shapes', nodeIds: [byName(e, 'Group 9')] });
    e.exec({ type: 'paste', payload: e.copy()! });
    const ids: string[] = [];
    const walk = (el: Element): void => {
      for (const c of el.children)
        if (c.type === 'element') {
          if (c.local === 'cNvPr') ids.push(attr(c, 'id')!);
          walk(c);
        }
    };
    walk(root(e));
    expect(new Set(ids).size).toBe(ids.length);
  });

  test('a hostile payload is refused, external links are dropped', async () => {
    const e = await open();
    for (const bad of [
      null,
      { version: 2 },
      { version: 1, document: 'x', slidePart: 'x', shapes: [{ ns: 1 }], relationships: [] },
      {
        version: 1,
        document: 'x',
        slidePart: 'x',
        shapes: [{ ns: 'urn:x', local: 'script' }],
        relationships: [],
      },
    ])
      expect(e.exec({ type: 'paste', payload: bad as unknown as ClipboardPayload }).ok).toBe(false);
    const linked = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            x: 0,
            y: 0,
            cx: IN,
            cy: IN,
            text: para(run('link', '', '<a:hlinkClick r:id="rId99"/>')),
          }),
          // The builder writes targets verbatim; this one also closes the attribute to add TargetMode.
          rels: [
            [
              'rId99',
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
              'https://example.com/" TargetMode="External',
            ],
          ],
        },
      ],
    });
    const src = await open(linked);
    src.selectAll();
    const payload = src.copy()!;
    expect(payload.relationships).toEqual([]);
    expect(e.exec({ type: 'paste', payload }).ok).toBe(true);
    expect(JSON.stringify(root(e))).not.toContain('hlinkClick');
  });
});
