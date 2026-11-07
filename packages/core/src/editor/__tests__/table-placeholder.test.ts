// Insert > Table on a slide with an empty content placeholder: the table takes the placeholder's
// place, box, and index, so the slide's layout positions it (ECMA-376 Part 1, 19.3.1.36 ph).

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run } from '../../../../../scripts/fixtures/builder.ts';
import { timingWriter } from '../../../../../scripts/fixtures/timing.ts';
import { attr, elements, find, path, xmlPartOf, type Element } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

/** The content placeholder of the master: its box when the slide and layout give none. */
const BOX = { x: 838_200, y: 1_825_625, cx: 10_515_600 };

const placeholder = (id: number, idx: string, text = '', xfrm = '') =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Content Placeholder ${id}"/><p:cNvSpPr>` +
  `<a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="${idx}"/></p:nvPr></p:nvSpPr>` +
  `<p:spPr>${xfrm}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>` +
  `${text ? para(run(text)) : '<a:p><a:endParaRPr lang="en-US"/></a:p>'}</p:txBody></p:sp>`;

const title =
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
  '<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>' +
  `${para(run('Results'))}</p:txBody></p:sp>`;

async function open(shapes: string): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(buildPptx({ slides: [{ layout: 'title', shapes }] }))).ok).toBe(true);
  return e;
}

const tree = (e: PptxEditorInstance): Element[] => {
  const part = e.session!.info.slides[0]!.partName;
  return elements(find(xmlPartOf(e.session!.store.package, part)!.root, 'p:spTree')!).filter(
    (c) => c.local === 'sp' || c.local === 'graphicFrame'
  );
};
const nameOf = (el: Element) =>
  attr(find(el, 'p:cNvPr'), 'name') + (el.local === 'graphicFrame' ? ' (table)' : '');
const offOf = (el: Element) => {
  const off = path(el, 'p:xfrm', 'a:off');
  const ext = path(el, 'p:xfrm', 'a:ext');
  return { x: Number(attr(off, 'x')), y: Number(attr(off, 'y')), cx: Number(attr(ext, 'cx')) };
};

describe('a new table fills an empty content placeholder', () => {
  test('it takes the place, box, and index of the placeholder; undo brings it back', async () => {
    const e = await open(title + placeholder(3, '1'));
    expect(e.exec({ type: 'insertTable', rows: 2, cols: 3 }).ok).toBe(true);
    const els = tree(e);
    expect(els.map(nameOf)).toEqual(['Title 1', 'Content Placeholder 3 (table)']);
    const frame = els[1]!;
    expect(attr(path(frame, 'p:nvGraphicFramePr', 'p:nvPr', 'p:ph'), 'idx')).toBe('1');
    const colW = Math.floor(BOX.cx / 3);
    expect(offOf(frame)).toEqual({ x: BOX.x, y: BOX.y, cx: colW * 3 });
    // The caret goes to the first cell.
    expect(e.snapshot().selection).toMatchObject({ kind: 'text', cell: { row: 0, column: 0 } });
    // Save and reopen keep the table as the placeholder.
    const again = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await again.load(e.save()!)).ok).toBe(true);
    const saved = tree(again)[1]!;
    expect(attr(path(saved, 'p:nvGraphicFramePr', 'p:nvPr', 'p:ph'), 'idx')).toBe('1');
    // One undo step brings the empty placeholder back.
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(tree(e).map(nameOf)).toEqual(['Title 1', 'Content Placeholder 3']);
  });

  test('the selected empty placeholder wins over the first one', async () => {
    const own =
      '<a:xfrm><a:off x="6096000" y="1825625"/><a:ext cx="5181600" cy="4351338"/></a:xfrm>';
    const e = await open(title + placeholder(3, '1') + placeholder(4, '2', '', own));
    const second = tree(e)[2]!.id;
    e.setSelection({ kind: 'shapes', nodeIds: [second] });
    expect(e.exec({ type: 'insertTable', rows: 1, cols: 2 }).ok).toBe(true);
    expect(tree(e).map(nameOf)).toEqual([
      'Title 1',
      'Content Placeholder 3',
      'Content Placeholder 4 (table)',
    ]);
    expect(offOf(tree(e)[2]!)).toEqual({ x: 6_096_000, y: 1_825_625, cx: 5_181_600 });
  });

  test('a placeholder with text stays, and the table goes in the middle of the slide', async () => {
    const e = await open(title + placeholder(3, '1', 'Keep me'));
    expect(e.exec({ type: 'insertTable', rows: 2, cols: 2 }).ok).toBe(true);
    const els = tree(e);
    expect(els.map(nameOf)).toEqual(['Title 1', 'Content Placeholder 3', 'Table 3 (table)']);
    expect(path(els[2]!, 'p:nvGraphicFramePr', 'p:nvPr', 'p:ph')).toBeUndefined();
    const { width } = e.session!.info;
    const { x, cx } = offOf(els[2]!);
    expect(Math.abs(x - (width - cx) / 2)).toBeLessThanOrEqual(1);
  });

  test('the placeholder that holds the caret wins; a turned one is passed over', async () => {
    const own =
      '<a:xfrm><a:off x="6096000" y="1825625"/><a:ext cx="5181600" cy="4351338"/></a:xfrm>';
    const turned =
      '<a:xfrm rot="5400000"><a:off x="838200" y="1825625"/><a:ext cx="5181600" cy="4351338"/></a:xfrm>';
    const e = await open(title + placeholder(3, '1', '', turned) + placeholder(4, '2', '', own));
    // Without a selection, the turned first placeholder is passed over for the second.
    expect(e.exec({ type: 'insertTable', rows: 1, cols: 1 }).ok).toBe(true);
    expect(tree(e).map(nameOf)).toEqual([
      'Title 1',
      'Content Placeholder 3',
      'Content Placeholder 4 (table)',
    ]);
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    // With the caret in the second placeholder, it is the one filled.
    const plain = await open(title + placeholder(3, '1') + placeholder(4, '2', '', own));
    const second = tree(plain)[2]!.id;
    plain.setSelection({
      kind: 'text',
      nodeId: second,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 0 },
    });
    expect(plain.exec({ type: 'insertTable', rows: 1, cols: 2 }).ok).toBe(true);
    expect(tree(plain).map(nameOf)).toEqual([
      'Title 1',
      'Content Placeholder 3',
      'Content Placeholder 4 (table)',
    ]);
  });

  test('a table placeholder is filled; body, hidden, and formatted placeholders are not', async () => {
    const typed = (id: number, ph: string, extra = '', spPr = '') =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="P ${id}"${extra}/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr>` +
      `</p:nvSpPr><p:spPr>${spPr}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`;
    const red = '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>';
    const skipped =
      typed(3, '<p:ph type="body" idx="1"/>') +
      typed(4, '<p:ph idx="2"/>', ' hidden="1"') +
      typed(5, '<p:ph idx="3"/>', '', red);
    const none = await open(title + skipped);
    expect(none.exec({ type: 'insertTable', rows: 1, cols: 1 }).ok).toBe(true);
    expect(tree(none).at(-1)!.local).toBe('graphicFrame');
    expect(path(tree(none).at(-1)!, 'p:nvGraphicFramePr', 'p:nvPr', 'p:ph')).toBeUndefined();
    const e = await open(title + skipped + typed(6, '<p:ph type="tbl" idx="4"/>'));
    expect(e.exec({ type: 'insertTable', rows: 1, cols: 1 }).ok).toBe(true);
    const ph = path(tree(e).at(-1)!, 'p:nvGraphicFramePr', 'p:nvPr', 'p:ph');
    expect([attr(ph, 'type'), attr(ph, 'idx')]).toEqual(['tbl', '4']);
  });

  test('the table keeps the shape id, so an animation of the placeholder stays', async () => {
    const { effect, click, mainSeq, timing } = timingWriter();
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const deck = buildPptx({
      slides: [
        {
          layout: 'title',
          shapes: title + placeholder(3, '1'),
          ext: timing(mainSeq(click(effect(3))), [3]),
        },
      ],
    });
    expect((await e.load(deck)).ok).toBe(true);
    expect(e.exec({ type: 'insertTable', rows: 1, cols: 1 }).ok).toBe(true);
    const frame = tree(e)[1]!;
    expect(attr(find(frame, 'p:cNvPr'), 'id')).toBe('3');
    const part = e.session!.info.slides[0]!.partName;
    const root = xmlPartOf(e.session!.store.package, part)!.root;
    expect(attr(find(root, 'p:spTgt'), 'spid')).toBe('3');
  });
});
