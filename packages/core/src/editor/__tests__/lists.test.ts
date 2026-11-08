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

  test('Enter on an empty list item starts another item at the same level', async () => {
    const e = await open(para(run('one')));
    caret(e, 0, 3);
    e.exec({ type: 'toggleBullets' });
    caret(e, 0, 0);
    press(e, 'Tab');
    caret(e, 0, 3);
    press(e, 'Enter');
    press(e, 'Enter');
    expect(paragraphs(e)).toHaveLength(3);
    for (const i of [1, 2]) {
      expect(resolved(e, i)?.bullet?.type).toBe('char');
      expect(resolved(e, i)?.level).toBe(1);
    }
  });
});

describe('Tab, Shift+Tab, and Backspace in lists', () => {
  test('Tab moves a list item a level in and its text with it; Shift+Tab moves it back', async () => {
    const e = await open(para(run('one')) + para(run('two')));
    caret(e, 1, 2);
    e.exec({ type: 'toggleBullets' });
    const before = resolved(e, 1)!.marL;
    caret(e, 1, 0);
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

  test('Tab inside the text of a list item types a tab, as does Shift+Tab', async () => {
    const e = await open(para(run('ab')));
    caret(e, 0, 0);
    e.exec({ type: 'toggleBullets' });
    caret(e, 0, 1);
    expect(press(e, 'Tab')).toEqual({ kind: 'input', input: 'text', data: '\t' });
    expect(press(e, 'Tab', true)).toEqual({ kind: 'input', input: 'text', data: '\t' });
    expect(resolved(e, 0)!.level).toBe(0);
  });

  test('Tab at the start of a paragraph with no bullet types a tab to the next tab stop', async () => {
    const e = await open(para(run('ab')));
    caret(e, 0, 0);
    expect(press(e, 'Tab')).toEqual({ kind: 'input', input: 'text', data: '\t' });
    const text = resolved(e, 0)!
      .runs.map((r) => (r.kind === 'text' ? r.text : ''))
      .join('');
    expect(text).toBe('\tab');
    // The text after the tab starts at the first default tab stop (1 inch).
    const item = e.session!.layout(0)!.items.find((x) => x.kind === 'shape' && x.element.text);
    const line = item?.kind === 'shape' ? item.text!.lines[0] : undefined;
    expect(line!.caretX[1]! - line!.caretX[0]!).toBeGreaterThan(IN / 2);
  });

  test('Tab over a selection that starts a list item, or across paragraphs, changes the level', async () => {
    const e = await open(para(run('one')) + para(run('two')));
    const select = (from: [number, number], to: [number, number]) =>
      e.setSelection({
        kind: 'text',
        nodeId: shapeId(e),
        anchor: { paragraph: from[0], offset: from[1] },
        focus: { paragraph: to[0], offset: to[1] },
      });
    select([0, 0], [1, 3]);
    e.exec({ type: 'toggleBullets' });
    select([1, 3], [1, 0]);
    press(e, 'Tab');
    expect(resolved(e, 1)!.level).toBe(1);
    select([0, 2], [1, 1]);
    press(e, 'Tab');
    expect([resolved(e, 0)!.level, resolved(e, 1)!.level]).toEqual([1, 2]);
    select([0, 1], [0, 2]);
    expect(press(e, 'Tab')).toEqual({ kind: 'input', input: 'text', data: '\t' });
  });

  test('Backspace at the start of a list item removes the bullet, keeps the level, then merges', async () => {
    const e = await open(para(run('one')) + para(run('two')));
    caret(e, 1, 0);
    e.exec({ type: 'toggleBullets' });
    press(e, 'Tab');
    const before = resolved(e, 1)!;
    press(e, 'Backspace');
    expect(paragraphs(e)).toHaveLength(2);
    const after = resolved(e, 1)!;
    expect(after.bullet).toBeNull();
    expect(after.level).toBe(1);
    // The text moves to where the bullet was.
    expect(after.marL).toBe(before.marL + before.indent);
    expect(after.indent).toBe(0);
    expect(press(e, 'Backspace')).toBeNull();
    expect(paragraphs(e)).toHaveLength(1);
  });

  test('Backspace in a body placeholder keeps the level and the bullet position of the master', async () => {
    const e = await open(para(run('one')) + para(run('two'), '<a:pPr lvl="1"/>'), BODY);
    const before = resolved(e, 1)!;
    expect(before.bullet).not.toBeNull();
    caret(e, 1, 0);
    press(e, 'Backspace');
    const after = resolved(e, 1)!;
    expect(after.bullet).toBeNull();
    expect(after.level).toBe(1);
    expect(after.marL).toBe(before.marL + before.indent);
    expect(child(pPr(e, 1)!, 'a:buNone')).toBeDefined();
  });
});

describe('Backspace removes a bullet in place', () => {
  test('bullets back on restore the placeholder hanging indent', async () => {
    const e = await open(para(run('one')) + para(run('two'), '<a:pPr lvl="1"/>'), BODY);
    const before = resolved(e, 1)!;
    caret(e, 1, 0);
    press(e, 'Backspace');
    e.exec({ type: 'toggleBullets' });
    const after = resolved(e, 1)!;
    expect(after.bullet?.type).toBe(before.bullet?.type);
    expect([after.level, after.marL, after.indent]).toEqual([1, before.marL, before.indent]);
    expect(attr(pPr(e, 1)!, 'marL')).toBeUndefined();
  });

  test('bullets back on keep a margin the author wrote with no bullet', async () => {
    const e = await open(
      para(run('one'), '<a:pPr marL="914400" indent="0"><a:buNone/></a:pPr>'),
      BODY
    );
    caret(e, 0, 0);
    e.exec({ type: 'toggleBullets' });
    expect(resolved(e, 0)?.bullet).not.toBeNull();
    expect(resolved(e, 0)?.marL).toBe(914_400);
  });

  test('a numbered item and an empty item lose the number or bullet first', async () => {
    const e = await open(para(run('one')) + para(''));
    caret(e, 0, 0);
    e.exec({ type: 'toggleNumbering' });
    press(e, 'Backspace');
    expect(resolved(e, 0)?.bullet).toBeNull();
    caret(e, 1, 0);
    e.exec({ type: 'toggleBullets' });
    press(e, 'Backspace');
    expect(paragraphs(e)).toHaveLength(2);
    expect(resolved(e, 1)?.bullet).toBeNull();
  });

  test('one undo brings the bullet back, and no bullet means nothing to do', async () => {
    const e = await open(para(run('one')));
    caret(e, 0, 0);
    expect(e.can({ type: 'removeBullet' }).ok).toBe(false);
    e.exec({ type: 'toggleBullets' });
    expect(e.can({ type: 'removeBullet' }).ok).toBe(true);
    press(e, 'Backspace');
    expect(resolved(e, 0)?.bullet).toBeNull();
    e.exec({ type: 'undo' });
    expect(resolved(e, 0)?.bullet?.type).toBe('char');
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

  test('a right-to-left ghost bullet stands in its cells and is not reordered again', () => {
    const doc = new GlobalWindow().document as unknown as Document;
    const pending = {
      kind: 'bullet',
      text: '(\u05d0',
      x: [100, 110, 120],
      rtl: true,
      face: { id: 'f' },
      sizePt: 18,
      rise: 0,
      props: { fill: { type: 'none' } },
    };
    const item = {
      kind: 'shape',
      text: { lines: [{ paragraph: 0, baseline: 0, pendingBullet: pending }] },
    } as unknown as Parameters<typeof ghostBullet>[1];
    const sel = {
      kind: 'text',
      nodeId: 'n',
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 0 },
    } as const;
    const text = ghostBullet(doc, item, sel, [1, 0, 0, 1, 0, 0])!.querySelector('text')!;
    // Source order, each character in its cell from the right, the bracket already mirrored.
    expect(text.textContent).toBe(')\u05d0');
    expect(text.getAttribute('x')).toBe(`${110 / 12_700} ${100 / 12_700}`);
    expect(text.getAttribute('style')).toContain('unicode-bidi:bidi-override');
  });
});
