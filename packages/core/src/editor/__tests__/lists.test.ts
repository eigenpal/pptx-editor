// Lists and the Enter, Tab, and Backspace keys while editing text. Each test failed before the
// fix: Enter made a line break inside one paragraph (a textarea reports Enter as
// `insertLineBreak`), Tab had no handler, indent did not move a list with its own margin, and
// bullets on overwrote a placeholder's inherited bullet.

import { describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { attr, child, elements, find, xmlPartOf, type Element } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { listKeyAction, type ListKeyAction } from '../list-keys.ts';
import { ghostBullet } from '../surface-ghost.ts';
import { textFrameMatrix } from '../surface.ts';

const IN = 914_400;
const BODY = '<p:ph type="body" idx="1"/>';

async function open(text: string, ph?: string): Promise<PptxEditorInstance> {
  const bytes = buildPptx({
    slides: [
      { shapes: shape({ id: 2, name: 'T', x: IN, y: IN, cx: 6 * IN, cy: 3 * IN, text, ph }) },
    ],
  });
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const root = (e: PptxEditorInstance): Element =>
  xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!.root;
const shapeId = (e: PptxEditorInstance): string =>
  elements(find(root(e), 'p:spTree')!).find((c) => attr(find(c, 'p:cNvPr'), 'name') === 'T')!.id;
const paragraphs = (e: PptxEditorInstance): Element[] =>
  elements(find(root(e), 'p:txBody')!).filter((c) => c.local === 'p');
const pPr = (e: PptxEditorInstance, i: number): Element | undefined =>
  child(paragraphs(e)[i], 'a:pPr');
const resolved = (e: PptxEditorInstance, i: number) => {
  const item = e.session!.layout(0)!.items.find((x) => x.kind === 'shape' && x.element.text);
  return item?.kind === 'shape' ? item.element.text!.paragraphs[i] : undefined;
};
const caret = (e: PptxEditorInstance, paragraph: number, offset: number) =>
  e.setSelection({
    kind: 'text',
    nodeId: shapeId(e),
    anchor: { paragraph, offset },
    focus: { paragraph, offset },
  });

/** Press a key the way the surface does: the list rules first, then the plain input. */
function press(e: PptxEditorInstance, key: string, shift = false): ListKeyAction | null {
  const sel = e.snapshot().selection;
  if (sel.kind !== 'text') throw new Error('not editing text');
  const item = e.session!.layout(0)!.items.find((x) => x.kind === 'shape' && x.element.text);
  const action = listKeyAction(item, sel, key, { shift, other: false });
  if (action?.kind === 'input') e.input(action.input, action.data);
  else if (action) e.exec(action.command);
  else if (key === 'Backspace') e.input('deleteBackward');
  return action;
}

describe('Enter and Shift+Enter', () => {
  test('Enter starts a paragraph; Shift+Enter inserts a line break', async () => {
    const e = await open(para(run('one')));
    caret(e, 0, 3);
    press(e, 'Enter');
    e.input('text', 'two');
    expect(paragraphs(e)).toHaveLength(2);
    press(e, 'Enter', true);
    e.input('text', 'more');
    expect(paragraphs(e)).toHaveLength(2);
    expect(find(paragraphs(e)[1], 'a:br')).toBeDefined();
  });

  test('bullets apply to the caret paragraph only, and Enter continues the list', async () => {
    const e = await open(para(run('one')) + para(run('two')) + para(run('three')));
    caret(e, 1, 1);
    e.exec({ type: 'toggleBullets' });
    expect(resolved(e, 0)?.bullet).toBeNull();
    expect(resolved(e, 1)?.bullet?.type).toBe('char');
    expect(resolved(e, 2)?.bullet).toBeNull();
    caret(e, 1, 3);
    press(e, 'Enter');
    e.input('text', 'next');
    expect(resolved(e, 2)?.bullet?.type).toBe('char');
    expect(resolved(e, 3)?.bullet).toBeNull();
  });

  test('Enter on an empty list item leaves the list', async () => {
    const e = await open(para(run('one')));
    caret(e, 0, 3);
    e.exec({ type: 'toggleBullets' });
    press(e, 'Enter');
    expect(resolved(e, 1)?.bullet?.type).toBe('char');
    press(e, 'Enter');
    expect(paragraphs(e)).toHaveLength(2);
    expect(resolved(e, 1)?.bullet).toBeNull();
    expect(resolved(e, 1)?.marL).toBe(0);
  });
});

describe('Tab, Shift+Tab, and Backspace in lists', () => {
  test('Tab moves a list item a level in and its text with it; Shift+Tab moves it back', async () => {
    const e = await open(para(run('one')) + para(run('two')));
    caret(e, 1, 2);
    e.exec({ type: 'toggleBullets' });
    const before = resolved(e, 1)!.marL;
    press(e, 'Tab');
    expect(resolved(e, 1)!.level).toBe(1);
    expect(resolved(e, 1)!.marL).toBe(before + 457_200);
    press(e, 'Tab', true);
    expect(resolved(e, 1)!.level).toBe(0);
    expect(resolved(e, 1)!.marL).toBe(before);
  });

  test('Tab in plain text inserts a tab, not a level', async () => {
    const e = await open(para(run('ab')));
    caret(e, 0, 1);
    expect(press(e, 'Tab')).toEqual({ kind: 'input', input: 'text', data: '\t' });
    expect(resolved(e, 0)!.level).toBe(0);
  });

  test('Backspace at the start of a list item removes the bullet before it merges', async () => {
    const e = await open(para(run('one')) + para(run('two')));
    caret(e, 1, 0);
    e.exec({ type: 'toggleBullets' });
    press(e, 'Backspace');
    expect(paragraphs(e)).toHaveLength(2);
    expect(resolved(e, 1)?.bullet).toBeNull();
    press(e, 'Backspace');
    expect(paragraphs(e)).toHaveLength(1);
  });
});

describe('inherited bullets', () => {
  test('bullets back on in a body placeholder restore the master bullet, with no override', async () => {
    const e = await open(para(run('item')), BODY);
    expect(resolved(e, 0)?.bullet?.type).toBe('char');
    const inheritedMarL = resolved(e, 0)!.marL;
    caret(e, 0, 1);
    e.exec({ type: 'toggleBullets' });
    expect(resolved(e, 0)?.bullet).toBeNull();
    expect(child(pPr(e, 0)!, 'a:buNone')).toBeDefined();
    e.exec({ type: 'toggleBullets' });
    expect(resolved(e, 0)?.bullet?.type).toBe('char');
    expect(resolved(e, 0)!.marL).toBe(inheritedMarL);
    const own = pPr(e, 0);
    for (const name of ['a:buNone', 'a:buChar', 'a:buFont'])
      expect(own && child(own, name)).toBeFalsy();
    expect(own && attr(own, 'marL')).toBeFalsy();
  });

  test('bullets off in a text box remove the list margins with the bullet', async () => {
    const e = await open(para(run('item')));
    caret(e, 0, 1);
    e.exec({ type: 'toggleBullets' });
    e.exec({ type: 'toggleBullets' });
    expect(resolved(e, 0)?.bullet).toBeNull();
    expect(resolved(e, 0)!.marL).toBe(0);
    expect(resolved(e, 0)!.indent).toBe(0);
  });
});

describe('ghost bullets', () => {
  test('an empty list item paints no bullet but lays out a pending one and the caret after it', async () => {
    const e = await open(para(run('one')));
    caret(e, 0, 3);
    e.exec({ type: 'toggleBullets' });
    press(e, 'Enter');
    const item = e.session!.layout(0)!.items.find((x) => x.kind === 'shape' && x.element.text);
    if (item?.kind !== 'shape' || !item.text) throw new Error('no text');
    const [first, empty] = item.text.lines;
    const painted = first.runs.find((r) => r.kind === 'bullet')!;
    expect(empty.runs.some((r) => r.kind === 'bullet')).toBe(false);
    expect(empty.pendingBullet?.text).toBe('•');
    expect(empty.pendingBullet?.x[0]).toBe(painted.x[0]);
    expect(empty.caretX[0]).toBe(first.caretX[0]);
  });

  test('an empty numbered item shows the next number without taking it', async () => {
    const e = await open(para(run('one')));
    caret(e, 0, 3);
    e.exec({ type: 'toggleNumbering' });
    press(e, 'Enter');
    let item = e.session!.layout(0)!.items.find((x) => x.kind === 'shape' && x.element.text);
    if (item?.kind !== 'shape' || !item.text) throw new Error('no text');
    expect(item.text.lines[1].pendingBullet?.text).toBe('2.');
    e.input('text', 'two');
    item = e.session!.layout(0)!.items.find((x) => x.kind === 'shape' && x.element.text);
    if (item?.kind !== 'shape' || !item.text) throw new Error('no text');
    expect(item.text.lines[1].runs[0].text).toBe('2.');
  });

  test('the surface draws the ghost only for a caret in an empty list item', async () => {
    const e = await open(para(run('one')));
    caret(e, 0, 3);
    e.exec({ type: 'toggleBullets' });
    press(e, 'Enter');
    const doc = new GlobalWindow().document as unknown as Document;
    const item = e.session!.layout(0)!.items.find((x) => x.kind === 'shape' && x.element.text)!;
    if (item.kind !== 'shape') throw new Error('no shape');
    const at = (paragraph: number, offset: number) =>
      ({
        kind: 'text',
        nodeId: shapeId(e),
        anchor: { paragraph, offset },
        focus: { paragraph, offset },
      }) as const;
    const ghost = ghostBullet(doc, item, at(1, 0), textFrameMatrix(item));
    expect(ghost?.querySelector('text')?.textContent).toBe('•');
    expect(ghost?.getAttribute('class')).toBe('pptx-surface__ghost-bullet');
    expect(ghostBullet(doc, item, at(0, 3), textFrameMatrix(item))).toBeNull();
  });
});
