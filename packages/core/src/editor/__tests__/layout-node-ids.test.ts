// Node ids are unique in one part only: a layout shape and a slide shape at the same position
// in their parts have the same node id, whatever their `cNvPr` ids are. Commands and the
// editor's paint act on slide shapes only. A removed slide shape leaves the selection, a
// layout connector does not follow a slide shape, and a layout group is not the parent of a
// slide shape. A layout shape takes no placeholder prompt, a placeholder fill reads the slide
// placeholder's box, and a copy keeps slide order. A shape picture and a copy as a picture draw
// the slide shape, and a layout object frame does not turn off Copy.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { findItem, parentGroupOf } from '../../automation/commands/placement.ts';
import { emptyPlaceholderOf } from '../../automation/commands/placeholder-fill.ts';
import { OpBatch } from '../../automation/commands/op-batch.ts';
import type { LayoutItem } from '../../layout/index.ts';
import type { ClipboardNode } from '../../contracts/index.ts';
import { serializeXmlPart, xmlPartOf } from '../../store/index.ts';
import { adapterActionState } from '../chrome-adapter-state.ts';
import { copySelection } from '../clipboard.ts';
import { selectedItems } from '../clipboard-picture.ts';
import { renderImageWith } from '../render-image.ts';
import { withPrompts } from '../placeholder-prompts.ts';
import { createPptxEditor } from '../index.ts';

const IN = 914_400;

async function open(layoutShapes: string, slideShapes: string) {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const deck = buildPptx({
    layouts: [{ name: 'Branded', shapes: layoutShapes }],
    slides: [{ layoutName: 'Branded', shapes: slideShapes }],
  });
  expect((await e.load(deck)).ok).toBe(true);
  const items = e.session!.layout(0)!.items;
  const named = (name: string, layer = 'slide'): LayoutItem =>
    items.find((i) => i.element.name === name && i.element.layer === layer)!;
  return { e, items, named };
}

/** A straight connector glued from shape `st` to shape `en`. */
const connector = (id: number, st: number, en: number): string =>
  `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="Link"/><p:cNvCxnSpPr>` +
  `<a:stCxn id="${st}" idx="3"/><a:endCxn id="${en}" idx="1"/></p:cNvCxnSpPr><p:nvPr/>` +
  `</p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="${3 * IN}" y="${IN / 4}"/>` +
  `<a:ext cx="${3 * IN}" cy="0"/></a:xfrm><a:prstGeom prst="straightConnector1"><a:avLst/>` +
  `</a:prstGeom></p:spPr></p:cxnSp>`;

/** A layout group with one child. */
const group = (child: string): string =>
  `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="20" name="Frame"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
  `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${IN}" cy="${IN}"/><a:chOff x="0" y="0"/>` +
  `<a:chExt cx="${IN}" cy="${IN}"/></a:xfrm></p:grpSpPr>${child}</p:grpSp>`;

const box = (id: number, name: string, x: number, y: number): string =>
  shape({ id, name, x, y, cx: IN, cy: IN });

describe('a layout shape with the node id of a slide shape', () => {
  test('a slide shape that another edit removes leaves the selection', async () => {
    const { e, items, named } = await open(box(9, 'Logo', 0, 0), box(2, 'A', 3 * IN, 3 * IN));
    const id = named('A').element.nodeId;
    expect(items.find((i) => i.element.name === 'Logo')!.element.nodeId).toBe(id);
    e.setSelection({ kind: 'shapes', nodeIds: [id] });
    // A host or a collaborator removes the shape through the store, not through a command.
    const part = e.session!.info.slides[0]!.partName;
    expect(e.session!.store.transact([{ op: 'removeNode', part, node: id }]).ok).toBe(true);
    // The selection does not move to the layout logo that has the same node id.
    expect(e.snapshot().selection.kind).toBe('none');
    expect(findItem(e.session!.layout(0)!.items, id)).toBeUndefined();
  });

  test('a layout connector does not follow a moved slide shape', async () => {
    const { e, items, named } = await open(
      box(5, 'Left', IN, 0) + box(6, 'Right', 6 * IN, 0) + connector(7, 5, 6),
      box(5, 'A', IN, 3 * IN) +
        box(6, 'B', 4 * IN, 3 * IN) +
        box(7, 'C', 6 * IN, 3 * IN) +
        box(8, 'D', 6 * IN, 5 * IN)
    );
    // The layout connector has the node id of slide shape C.
    expect(items.find((i) => i.element.name === 'Link')!.element.nodeId).toBe(
      named('C').element.nodeId
    );
    const part = () =>
      serializeXmlPart(xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!);
    const c = (xml: string) => xml.slice(xml.indexOf('name="C"'), xml.indexOf('name="D"'));
    const before = c(part());
    e.setSelection({ kind: 'shapes', nodeIds: [named('A').element.nodeId] });
    expect(e.exec({ type: 'nudge', dx: IN, dy: 0 }).ok).toBe(true);
    expect(c(part())).toBe(before);
  });

  test('a layout group is not the parent of a slide shape', async () => {
    // One slide shape before A puts A at the node id of the layout group's child.
    for (const pad of [0, 1, 2]) {
      const pads = Array.from({ length: pad }, (_, k) => box(30 + k, `Pad ${k}`, 0, 0)).join('');
      const { items, named } = await open(
        group(box(21, 'Inner', 0, 0)),
        pads + box(5, 'A', 3 * IN, 3 * IN)
      );
      expect(parentGroupOf(items, named('A').element.nodeId)).toBeUndefined();
    }
  });

  const title = (id: number, name: string, text: string) =>
    shape({
      id,
      name,
      x: IN,
      y: IN,
      cx: 6 * IN,
      cy: IN,
      ph: '<p:ph type="title"/>',
      text: para(text ? run(text) : ''),
    });
  const body = shape({
    id: 3,
    name: 'Body',
    x: IN,
    y: 3 * IN,
    cx: 6 * IN,
    cy: 2 * IN,
    ph: '<p:ph idx="1"/>',
    text: para(''),
  });
  const logo = shape({
    id: 9,
    name: 'Logo',
    x: 8 * IN,
    y: 5 * IN,
    cx: IN,
    cy: IN / 2,
    rot: 90 * 60_000,
    text: para(run('Brand')),
  });

  test('a layout shape keeps its text when a slide placeholder shows its prompt', async () => {
    const { e, named } = await open(logo + title(2, 'Title Layout', ''), title(2, 'Title', ''));
    expect(named('Logo', 'layout').element.nodeId).toBe(named('Title').element.nodeId);
    const s = e.session!;
    const painted = withPrompts(s.store.package, s.measurer, s.layout(0)!, (k) => k, null)!;
    const shown = painted.items.find(
      (i) => i.element.name === 'Logo' && i.element.layer === 'layout'
    )!;
    expect(shown).toBe(named('Logo', 'layout'));
    // The slide title still shows its prompt.
    const prompted = painted.items.find((i) => i.element.name === 'Title')!;
    expect(prompted).not.toBe(named('Title'));
  });

  test('a placeholder fill reads the box of the slide placeholder', async () => {
    const { e, named } = await open(
      logo + title(2, 'Title Layout', ''),
      title(2, 'Title', '') + body
    );
    const s = e.session!;
    const ph = emptyPlaceholderOf(
      new OpBatch(s.store.package),
      s.info.slides[0]!.partName,
      0,
      new Set(['title', 'body', 'obj']),
      undefined
    );
    // The rotated layout logo has the title's node id: the title is still the first candidate.
    expect(ph?.nodeId).toBe(named('Title').element.nodeId);
  });

  test('a copy keeps the slide order of the selected shapes', async () => {
    const { e, items, named } = await open('', box(2, 'A', IN, IN) + box(3, 'B', 3 * IN, IN));
    const a = named('A');
    const b = named('B');
    // Layout shapes with the slide shapes' node ids, in the opposite order, paint first.
    const asLayout = (i: LayoutItem) =>
      ({ ...i, element: { ...i.element, layer: 'layout' } }) as LayoutItem;
    const scope = [asLayout(b), asLayout(a), ...items];
    const s = e.session!;
    const payload = copySelection(
      s.store.package,
      s.info.slides[0]!.partName,
      scope,
      { nodeIds: [b.element.nodeId, a.element.nodeId] },
      'doc'
    )!;
    expect(payload.shapes.map(nameOf)).toEqual(['A', 'B']);
  });

  const labeled = (id: number, name: string, text: string) =>
    shape({ id, name, x: IN, y: IN, cx: 3 * IN, cy: IN, text: para(run(text)) });

  test('a shape picture draws the slide shape, not the layout shape', async () => {
    const { e, named } = await open(labeled(9, 'Header', 'Brand'), labeled(2, 'Title', 'Heading'));
    let svg = '';
    await renderImageWith(e.session!, 0, { shapeNodeId: named('Title').element.nodeId }, (r) => {
      svg = r.svg;
      throw new Error('no raster');
    }).catch(() => undefined);
    expect(svg).toContain('>Heading<');
    expect(svg).not.toContain('>Brand<');
  });

  test('a copy as a picture draws the selected slide shapes only', async () => {
    const { e, named } = await open(labeled(9, 'Header', 'Brand'), labeled(2, 'Title', 'Heading'));
    const id = named('Title').element.nodeId;
    const drawn = selectedItems(e.session!.layout(0)!, { kind: 'shapes', nodeIds: [id] });
    expect(drawn).toEqual([named('Title')]);
  });

  test('a layout object frame does not turn off Copy for a slide shape', async () => {
    const frame =
      `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="Object"/>` +
      `<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/>` +
      `<a:ext cx="${IN}" cy="${IN}"/></p:xfrm><a:graphic><a:graphicData ` +
      `uri="http://schemas.openxmlformats.org/presentationml/2006/ole"/></a:graphic>` +
      `</p:graphicFrame>`;
    const { e, items, named } = await open(frame, box(2, 'A', IN, IN));
    const layoutFrame = items.find((i) => i.element.name === 'Object')!;
    expect(layoutFrame.element.nodeId).toBe(named('A').element.nodeId);
    e.setSelection({ kind: 'shapes', nodeIds: [named('A').element.nodeId] });
    expect(adapterActionState(e, 'edit.copy').enabled).toBe(true);
  });
});

/** The `cNvPr` name of a copied shape. */
function nameOf(node: ClipboardNode): string | undefined {
  if (node.local === 'cNvPr') return node.attrs?.find((a) => a.local === 'name')?.value;
  for (const child of node.children ?? []) {
    if ('text' in child) continue;
    const name = nameOf(child);
    if (name !== undefined) return name;
  }
  return undefined;
}
