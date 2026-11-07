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

/** The node ids of the deck's text shapes and table frame, group children in place of groups. */
const targets = (e: PptxEditorInstance): string[] =>
  e
    .session!.layout(0)!
    .items.flatMap((i) =>
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
