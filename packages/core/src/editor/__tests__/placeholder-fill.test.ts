// Empty placeholders: a new picture takes the place of an empty content or picture placeholder,
// and the surface chrome draws a dotted outline and insert buttons on empty placeholders
// (`automation/commands/placeholder-fill.ts`, `surface-placeholder-chrome.ts`).

import { describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, para, run } from '../../../../../scripts/fixtures/builder.ts';
import { attr, elements, find, path, xmlPartOf, type Element } from '../../store/index.ts';
import { emptyPlaceholderType } from '../../automation/commands/placeholder-fill.ts';
import type { EditorCommand } from '../../contracts/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { PlaceholderChrome, type PlaceholderChromeHost } from '../surface-placeholder-chrome.ts';

/** The content placeholder of the master: its box when the slide and layout give none. */
const BOX = { x: 838_200, y: 1_825_625, cx: 10_515_600, cy: 4_351_338 };
const PX = 9_525;
/** A 1 × 1 PNG. */
const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
  ),
  (c) => c.charCodeAt(0)
);

const placeholder = (id: number, ph: string, text = '', xfrm = '') =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Placeholder ${id}"/><p:cNvSpPr>` +
  `<a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${ph}</p:nvPr></p:nvSpPr>` +
  `<p:spPr>${xfrm}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>` +
  `${text ? para(run(text)) : '<a:p><a:endParaRPr lang="en-US"/></a:p>'}</p:txBody></p:sp>`;
const title = (text = 'Results') => placeholder(2, '<p:ph type="title"/>', text);
const content = placeholder(3, '<p:ph idx="1"/>');

async function open(shapes: string): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(buildPptx({ slides: [{ layout: 'title', shapes }] }))).ok).toBe(true);
  return e;
}

const tree = (e: PptxEditorInstance): Element[] => {
  const part = e.session!.info.slides[0]!.partName;
  return elements(find(xmlPartOf(e.session!.store.package, part)!.root, 'p:spTree')!).filter((c) =>
    ['sp', 'pic', 'graphicFrame'].includes(c.local)
  );
};
const boxOf = (el: Element) => {
  const off = find(el, 'a:off');
  const ext = find(el, 'a:ext');
  return {
    x: Number(attr(off, 'x')),
    y: Number(attr(off, 'y')),
    cx: Number(attr(ext, 'cx')),
    cy: Number(attr(ext, 'cy')),
  };
};
const image = (width: number, height: number): EditorCommand => ({
  type: 'insertImage',
  bytes: PNG,
  mime: 'image/png',
  width,
  height,
});

describe('a new picture fills an empty placeholder', () => {
  test('a content placeholder holds the picture whole, centered, at its own size at most', async () => {
    const e = await open(title() + content);
    // 400 × 200 px: 3,810,000 × 1,905,000 EMU, which fits inside the box as it is.
    expect(e.exec(image(400, 200)).ok).toBe(true);
    const els = tree(e);
    expect(els.map((el) => el.local)).toEqual(['sp', 'pic']);
    const pic = els[1]!;
    expect(attr(path(pic, 'p:nvPicPr', 'p:nvPr', 'p:ph'), 'idx')).toBe('1');
    expect(boxOf(pic)).toEqual({
      x: BOX.x + (BOX.cx - 3_810_000) / 2,
      y: Math.round(BOX.y + (BOX.cy - 1_905_000) / 2),
      cx: 3_810_000,
      cy: 1_905_000,
    });
    expect(find(pic, 'a:srcRect')).toBeUndefined();
    // One undo step brings the empty placeholder back.
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(tree(e).map((el) => el.local)).toEqual(['sp', 'sp']);
  });

  test('a picture placeholder is filled, and the picture is cropped to its shape', async () => {
    const own = '<a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="3657600"/></a:xfrm>';
    const e = await open(title() + placeholder(3, '<p:ph type="pic" idx="1"/>', '', own));
    // A 2:1 picture in a square box: half its width goes, a quarter from each side.
    expect(e.exec(image(800, 400)).ok).toBe(true);
    const pic = tree(e)[1]!;
    expect(attr(path(pic, 'p:nvPicPr', 'p:nvPr', 'p:ph'), 'type')).toBe('pic');
    expect(boxOf(pic)).toEqual({ x: 914_400, y: 914_400, cx: 3_657_600, cy: 3_657_600 });
    const crop = find(pic, 'a:srcRect')!;
    expect([attr(crop, 'l'), attr(crop, 'r'), attr(crop, 't')]).toEqual([
      '25000',
      '25000',
      undefined,
    ]);
  });

  test('a slide without an empty placeholder centers the picture as before', async () => {
    const e = await open(title() + placeholder(3, '<p:ph idx="1"/>', 'Text'));
    expect(e.exec(image(100, 100)).ok).toBe(true);
    const pic = tree(e)[2]!;
    expect(path(pic, 'p:nvPicPr', 'p:nvPr', 'p:ph')).toBeUndefined();
    expect(attr(find(pic, 'a:picLocks'), 'noGrp')).toBeUndefined();
  });
});

const doc = new GlobalWindow().document as unknown as Document;

function chrome(e: PptxEditorInstance, readOnly = false) {
  const commands: EditorCommand[] = [];
  const picked: string[] = [];
  const host: PlaceholderChromeHost = {
    selection: () => e.snapshot().selection,
    readOnly: () => readOnly,
    label: (key) => key,
    select: (sel) => e.setSelection(sel),
    run: (command) => {
      commands.push(command);
      expect(e.exec(command).ok).toBe(true);
    },
    pickImage: () => picked.push(JSON.stringify(e.snapshot().selection)),
    fillableType: (id) => {
      const el = tree(e).find((c) => c.id === id);
      return el ? emptyPlaceholderType(el) : undefined;
    },
  };
  return { c: new PlaceholderChrome(host), commands, picked };
}

const layoutOf = (e: PptxEditorInstance) => e.session!.layout(0)!;
const buttons = (parts: globalThis.Element[]) =>
  parts
    .map((p) => p.getAttribute('data-placeholder-insert'))
    .filter((v): v is string => !!v)
    .map((v) => v.split('|')[1]);

describe('the placeholder chrome', () => {
  test('empty placeholders show a dotted outline; a content one shows its insert buttons', async () => {
    const e = await open(title('') + content);
    const { c } = chrome(e);
    const parts = c.elements(doc, layoutOf(e), PX);
    const outlines = parts.filter(
      (p) => p.getAttribute('class') === 'pptx-surface__placeholder-outline'
    );
    expect(outlines.length).toBe(2);
    expect(outlines[1]!.getAttribute('stroke-dasharray')).toBe(`${PX} ${2 * PX}`);
    expect(buttons(parts)).toEqual(['table', 'chart', 'picture']);
    // A placeholder with text shows nothing, and view mode shows nothing.
    const filled = await open(title() + placeholder(3, '<p:ph idx="1"/>', 'Text'));
    expect(chrome(filled).c.elements(doc, layoutOf(filled), PX)).toEqual([]);
    expect(chrome(e, true).c.elements(doc, layoutOf(e), PX)).toEqual([]);
  });

  test('the placeholder being typed in shows no chrome', async () => {
    const e = await open(title() + content);
    const id = tree(e)[1]!.id;
    e.setSelection({
      kind: 'text',
      nodeId: id,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 0 },
    });
    expect(chrome(e).c.elements(doc, layoutOf(e), PX)).toEqual([]);
  });

  test('insert buttons stay out of the way while another object is selected or edited', async () => {
    const e = await open(title() + content);
    const titleId = tree(e)[0]!.id;
    const contentId = tree(e)[1]!.id;
    const { c } = chrome(e);
    for (const selection of [
      { kind: 'shapes', nodeIds: [titleId] },
      { kind: 'shapes', nodeIds: [titleId, contentId] },
      {
        kind: 'text',
        nodeId: titleId,
        anchor: { paragraph: 0, offset: 0 },
        focus: { paragraph: 0, offset: 0 },
      },
    ] as const) {
      e.setSelection(selection);
      const parts = c.elements(doc, layoutOf(e), PX);
      expect(buttons(parts)).toEqual([]);
      expect(
        parts.some((p) => p.getAttribute('class') === 'pptx-surface__placeholder-outline')
      ).toBe(true);
    }
    e.setSelection({ kind: 'none' });
    expect(buttons(c.elements(doc, layoutOf(e), PX))).toEqual(['table', 'chart', 'picture']);
    e.setSelection({ kind: 'shapes', nodeIds: [contentId] });
    expect(buttons(c.elements(doc, layoutOf(e), PX))).toEqual(['table', 'chart', 'picture']);
  });

  test('the table button fills the placeholder; the picture button asks for a file', async () => {
    const e = await open(title() + content);
    const id = tree(e)[1]!.id;
    const { c, commands, picked } = chrome(e);
    const press = (insert: string) => {
      const parts = c.elements(doc, layoutOf(e), PX);
      const button = parts.find(
        (p) => p.getAttribute('data-placeholder-insert') === `${id}|${insert}`
      )!;
      return c.press(layoutOf(e), button.firstElementChild!);
    };
    const centerOf = (insert: string) => {
      const parts = c.elements(doc, layoutOf(e), PX);
      const g = parts.find((p) => p.getAttribute('data-placeholder-insert') === `${id}|${insert}`)!;
      const [x, y] = g
        .getAttribute('transform')!
        .match(/[-\d.]+/g)!
        .map(Number);
      return { x: x! + 14 * PX, y: y! + 14 * PX };
    };
    // The Picture press selects the placeholder; a release over the button asks for the file.
    const center = centerOf('picture');
    expect(press('picture')).toBe(true);
    expect(picked).toEqual([]);
    expect(c.release(center)).toBe(true);
    expect(picked).toEqual([JSON.stringify({ kind: 'shapes', nodeIds: [id] })]);
    // A release with no press before it asks for nothing; so does a drag off the button.
    expect(c.release(center)).toBe(false);
    expect(press('picture')).toBe(true);
    expect(c.release({ x: 0, y: 0 })).toBe(false);
    expect(picked).toHaveLength(1);
    e.setSelection({ kind: 'none' });
    expect(press('table')).toBe(true);
    expect(commands).toEqual([{ type: 'insertTable', rows: 3, cols: 3 }]);
    expect(tree(e).map((el) => el.local)).toEqual(['sp', 'graphicFrame']);
    // A press elsewhere is not the chrome's.
    expect(c.press(layoutOf(e), doc.createElementNS('http://www.w3.org/2000/svg', 'rect'))).toBe(
      false
    );
  });

  test('the chart button fills the placeholder with a column chart', async () => {
    const e = await open(title() + content);
    const id = tree(e)[1]!.id;
    const { c, commands } = chrome(e);
    const button = c
      .elements(doc, layoutOf(e), PX)
      .find((p) => p.getAttribute('data-placeholder-insert') === `${id}|chart`)!;
    expect(button.querySelector('title')!.textContent).toBe('canvas.placeholderChart');
    expect(c.press(layoutOf(e), button.firstElementChild!)).toBe(true);
    expect(commands).toEqual([{ type: 'insertChart', chartType: 'columnClustered' }]);
    expect(tree(e).map((el) => el.local)).toEqual(['sp', 'graphicFrame']);
    const format = e.chartFormat();
    expect(format.ok && format.format.type).toBe('columnClustered');
  });

  test('a placeholder with its own fill shows its outline but no buttons', async () => {
    const red =
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Red"/><p:cNvSpPr/><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>' +
      '<p:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>' +
      '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>';
    const e = await open(title() + red);
    const parts = chrome(e).c.elements(doc, layoutOf(e), PX);
    expect(
      parts.filter((p) => p.getAttribute('class') === 'pptx-surface__placeholder-outline')
    ).toHaveLength(1);
    expect(buttons(parts)).toEqual([]);
  });
});

describe('a filled picture placeholder keeps its shape and its links', () => {
  test('a round picture placeholder gives a round picture', async () => {
    const own =
      '<a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="3657600"/></a:xfrm>' +
      '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>';
    const e = await open(title() + placeholder(3, '<p:ph type="pic" idx="1"/>', '', own));
    expect(e.exec(image(400, 400)).ok).toBe(true);
    const pic = tree(e)[1]!;
    expect(attr(find(pic, 'a:prstGeom'), 'prst')).toBe('ellipse');
    // The picture has an id of its own: two people who fill one placeholder make two ids.
    expect(attr(find(pic, 'p:cNvPr'), 'id')).not.toBe('3');
    expect(attr(find(pic, 'p:cNvPr'), 'name')).toBe('Placeholder 3');
  });
});
