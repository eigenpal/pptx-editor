// Commands on many shapes at once do their work in one pass: z-order moves, list commands with
// one cascade resolve for all targets, and one re-fit pass. These tests check that the one-pass
// result is the result of the command on each shape alone, across groups and table cells.

import { expect, test } from 'bun:test';
import { buildPptx, para, run, shape, table } from '../../../../../scripts/fixtures/builder.ts';
import {
  attr,
  elements,
  find,
  serializeXmlPart,
  xmlPartOf,
  type Element,
} from '../../store/index.ts';
import type { EditorCommand } from '../../contracts/index.ts';
import { reorder } from '../../automation/commands/shape-commands.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

const IN = 914_400;
const ext = '<p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"/></p:extLst>';
const group = (id: number, name: string, inner: string) =>
  `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvGrpSpPr/><p:nvPr/>` +
  '</p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="914400" y="914400"/>' +
  '<a:ext cx="2743200" cy="914400"/><a:chOff x="914400" y="914400"/>' +
  `<a:chExt cx="2743200" cy="914400"/></a:xfrm></p:grpSpPr>${inner}${ext}</p:grpSp>`;

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const rootOf = (e: PptxEditorInstance): Element =>
  xmlPartOf(e.session!.store.package, e.session!.info.slides[0]!.partName)!.root;
const nameOf = (el: Element) => attr(find(el, 'p:cNvPr'), 'name');

test('z-order moves across parents and in groups give the order of `reorder`, and undo', async () => {
  let id = 2;
  const sh = (n: string) => shape({ id: id++, name: n, x: IN, y: IN, cx: IN, cy: IN });
  const names = (p: string, n: number) => Array.from({ length: n }, (_, i) => `${p}${i}`);
  const top = names('T', 12);
  const bytes = buildPptx({
    slides: [
      {
        shapes:
          top.slice(0, 5).map(sh).join('') +
          group(500, 'GG', names('G', 15).map(sh).join('')) +
          top.slice(5).map(sh).join('') +
          group(501, 'HH', names('H', 9).map(sh).join('')) +
          ext,
      },
    ],
  });
  let seed = 11;
  const r = (n: number) => (seed = (seed * 1_103_515_245 + 12_345) % 2 ** 31) % n;
  for (let round = 0; round < 40; round += 1) {
    for (const move of ['front', 'back', 'forward', 'backward'] as const) {
      const e = await open(bytes);
      const tree = () => find(rootOf(e), 'p:spTree')!;
      const parents = () => [
        tree(),
        ...elements(tree()).filter((c) => nameOf(c) === 'GG' || nameOf(c) === 'HH'),
      ];
      // Each parent's shapes and its trailing `p:extLst`; the group properties stay first.
      const lists = () =>
        parents().map((p) =>
          elements(p)
            .filter((c) => c.local !== 'nvGrpSpPr' && c.local !== 'grpSpPr')
            .map((c) => (c.local === 'extLst' ? 'ext' : nameOf(c)!))
        );
      const start = lists();
      const byName = new Map(parents().flatMap((p) => elements(p).map((c) => [nameOf(c), c.id])));
      const share = [1, 3, 6, 9][round % 4]!;
      const picked: string[] = [];
      const want = start.map((l) => {
        // Shapes only: a group and its own child in one selection is another case.
        const chosen = new Set(
          l.slice(0, -1).filter((n) => n !== 'GG' && n !== 'HH' && r(10) < share)
        );
        picked.push(...chosen);
        return [...reorder(l.slice(0, -1), chosen, move), 'ext'];
      });
      if (picked.length > 0) {
        e.setSelection({ kind: 'shapes', nodeIds: picked.map((n) => byName.get(n)!) });
        const result = e.exec({ type: 'order', move });
        expect(lists()).toEqual(want);
        // A move that changes nothing is refused.
        expect(result.ok).toBe(JSON.stringify(want) !== JSON.stringify(start));
        if (result.ok) {
          e.exec({ type: 'undo' });
          expect(lists()).toEqual(start);
        }
      }
      e.destroy();
    }
  }
});

const bullet = '<a:pPr marL="457200" indent="-228600"><a:buChar char="•"/></a:pPr>';

/**
 * Text shapes, two in a group, and a 2×2 table, each with two paragraphs: all with bullets
 * (`listed`) or none. One state for all, so a toggle does the same on each shape alone.
 */
function listDeck(listed: boolean): Uint8Array {
  const p = (t: string) => para(run(t), listed ? bullet : undefined);
  const body = p('a') + p('b');
  const cell = `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${body}</a:txBody><a:tcPr/></a:tc>`;
  const box = (id: number) => shape({ id, x: IN, y: IN, cx: IN, cy: IN, text: body });
  const rows = `<a:tr h="370840">${cell}${cell}</a:tr>`.repeat(2);
  return buildPptx({
    slides: [
      {
        shapes:
          box(2) +
          box(3) +
          group(50, 'G', box(4) + box(5)) +
          table({
            id: 9,
            x: IN,
            y: 3 * IN,
            cx: 2 * IN,
            cy: IN,
            columns: [IN, IN],
            tblPr: '<a:tblPr/>',
            rows,
          }),
      },
    ],
  });
}

/** The ids of the slide's text shapes and table frame, group children in place of groups. */
const targets = (e: PptxEditorInstance): string[] =>
  e
    .session!.layout(0)!
    .items.filter((i) => i.element.layer === 'slide')
    .flatMap((i) =>
      i.kind === 'group' ? i.children.map((c) => c.element.nodeId) : [i.element.nodeId]
    );

const xml = (e: PptxEditorInstance) =>
  serializeXmlPart({ partName: 'slide', root: rootOf(e), nextId: 0 });

for (const [command, listed] of [
  [{ type: 'removeBullet' }, true],
  [{ type: 'toggleBullets' }, true],
  [{ type: 'toggleBullets' }, false],
  [{ type: 'toggleNumbering' }, false],
  [{ type: 'setAutofit', mode: 'shrink' }, false],
  [{ type: 'setBulletFormat', format: { kind: 'bullet', char: '-' } }, false],
  [{ type: 'setBulletFormat', format: { kind: 'number' } }, false],
  [{ type: 'setBulletFormat', format: { kind: 'bullet', sizePercent: 120 } }, true],
] as const satisfies readonly (readonly [EditorCommand, boolean])[]) {
  test(`${command.type} on many shapes gives the result of each shape alone`, async () => {
    const together = await open(listDeck(listed));
    together.setSelection({ kind: 'shapes', nodeIds: targets(together) });
    expect(together.exec(command).ok).toBe(true);
    const alone = await open(listDeck(listed));
    for (const id of targets(alone)) {
      alone.setSelection({ kind: 'shapes', nodeIds: [id] });
      alone.exec(command);
    }
    expect(xml(together)).toBe(xml(alone));
    together.destroy();
    alone.destroy();
  });
}

test('a shape selected twice gets one bullet', async () => {
  const e = await open(listDeck(false));
  const [id] = targets(e);
  e.setSelection({ kind: 'shapes', nodeIds: [id!, id!] });
  expect(e.exec({ type: 'toggleBullets' }).ok).toBe(true);
  const sp = elements(find(rootOf(e), 'p:spTree')!).find((c) => c.id === id)!;
  for (const pPr of elements(find(sp, 'p:txBody')!).flatMap((p) => elements(p)))
    if (pPr.local === 'pPr')
      expect(elements(pPr).filter((c) => c.local === 'buChar')).toHaveLength(1);
  e.destroy();
});

test('a slide shape with the node id of a layout shape edits and re-fits itself', async () => {
  // The same tree in the layout and the slide: node ids are numbered per part, so they match.
  const text = para(run('a'), bullet) + para(run('b'), bullet);
  const box = shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN, text });
  const e = await open(
    buildPptx({
      layouts: [{ name: 'Deco', shapes: box }],
      slides: [{ layoutName: 'Deco', shapes: box }],
    })
  );
  const items = e.session!.layout(0)!.items;
  const mine = items.find((i) => i.element.layer === 'slide')!;
  expect(
    items.some((i) => i.element.layer !== 'slide' && i.element.nodeId === mine.element.nodeId)
  ).toBe(true);
  e.setSelection({ kind: 'shapes', nodeIds: [mine.element.nodeId] });
  expect(e.exec({ type: 'setBulletFormat', format: { kind: 'bullet', sizePercent: 120 } }).ok).toBe(
    true
  );
  expect(xml(e)).toContain('<a:buSzPct val="120000"/>');
  // Shape fit re-fits the slide shape, not the layout one.
  expect(e.exec({ type: 'setAutofit', mode: 'shape' }).ok).toBe(true);
  const fitted = e.session!.layout(0)!.items.find((i) => i.element.layer === 'slide')!;
  expect(fitted.element.xfrm.cy).not.toBe(IN);
  e.destroy();
});

test('Bullet Format None re-fits only the shapes whose bullets it changes', async () => {
  // Shape B has no list and a fit scale of its own, which a re-fit would recompute.
  const plain =
    '<p:sp><p:nvSpPr><p:cNvPr id="3" name="B"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>' +
    '<a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr>' +
    '<a:normAutofit fontScale="40000" lnSpcReduction="20000"/></a:bodyPr><a:lstStyle/>' +
    '<a:p><a:pPr marL="0" indent="0"/><a:r><a:rPr lang="en-US"/><a:t>b</a:t></a:r></a:p>' +
    '</p:txBody></p:sp>';
  const listed = shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN, text: para(run('a'), bullet) });
  const e = await open(buildPptx({ slides: [{ shapes: listed + plain }] }));
  e.setSelection({ kind: 'shapes', nodeIds: targets(e) });
  expect(e.exec({ type: 'setBulletFormat', format: null }).ok).toBe(true);
  expect(xml(e)).toContain('fontScale="40000"');
  expect(xml(e)).not.toContain('<a:buChar');
  e.destroy();
});

test('a slide table with the node id of a layout table re-fits itself', async () => {
  const cell = `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${para(run('x'))}</a:txBody><a:tcPr/></a:tc>`;
  const frame = table({
    id: 2,
    x: IN,
    y: IN,
    cx: IN,
    cy: 370_840,
    columns: [IN],
    tblPr: '<a:tblPr/>',
    rows: `<a:tr h="370840">${cell}</a:tr>`,
  });
  const e = await open(
    buildPptx({
      layouts: [{ name: 'Deco', shapes: frame }],
      slides: [{ layoutName: 'Deco', shapes: frame }],
    })
  );
  const items = e.session!.layout(0)!.items;
  const mine = items.find((i) => i.element.layer === 'slide')!;
  const id = mine.element.nodeId;
  expect(items.some((i) => i.element.layer !== 'slide' && i.element.nodeId === id)).toBe(true);
  e.setSelection({ kind: 'shapes', nodeIds: [id] });
  expect(e.exec({ type: 'setFontSize', points: 60 }).ok).toBe(true);
  const grown = e.session!.layout(0)!.items.find((i) => i.element.layer === 'slide')!;
  expect(grown.element.xfrm.cy).toBeGreaterThan(370_840);
  e.destroy();
});
