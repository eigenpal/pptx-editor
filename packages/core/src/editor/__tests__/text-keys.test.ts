// Text editing keys and AutoCorrect while typing: paragraph alignment, font size, list level,
// and Change Case keys; Enter on a selected shape; the word deletion of Cmd+Backspace; and the
// corrections and lists that typing makes.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { attr, elements, find, xmlPartOf, type Element } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { corrections } from '../autocorrect.ts';
import { nextCase } from '../format-keys.ts';
import { dispatchBeforeInput } from '../surface-input.ts';
import type { TextInputKind } from '../text-input.ts';

const IN = 914_400;
const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

async function open(text: string, options: { autoCorrect?: boolean } = {}) {
  const g = globalThis as { ResizeObserver?: unknown };
  g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false, ...options });
  const bytes = buildPptx({
    slides: [{ shapes: shape({ id: 2, name: 'T', x: IN, y: IN, cx: 6 * IN, cy: 2 * IN, text }) }],
  });
  expect((await e.load(bytes)).ok).toBe(true);
  const host = doc.createElement('div') as unknown as HTMLElement;
  doc.body.append(host as unknown as Node);
  e.attach(host);
  const input = host.querySelector('textarea') as HTMLTextAreaElement;
  type Mods = { altKey?: boolean; ctrlKey?: boolean; metaKey?: boolean; shiftKey?: boolean };
  /** A key down on the editor's input; answers whether the editor took it. */
  const key = (k: string, mods: Mods = {}): boolean => {
    const ev = new win.KeyboardEvent('keydown', {
      key: k,
      bubbles: true,
      cancelable: true,
      ...mods,
    });
    input.dispatchEvent(ev as unknown as Event);
    return ev.defaultPrevented;
  };
  /** Type one character at a time, as a keyboard does. */
  const type = (text: string): void => {
    for (const ch of text) {
      if (ch === '\n') key('Enter');
      else e.input('text', ch);
    }
  };
  return { e, key, type };
}

const root = (e: PptxEditorInstance): Element =>
  xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!.root;
const shapeId = (e: PptxEditorInstance): string =>
  elements(find(root(e), 'p:spTree')!).find((c) => attr(find(c, 'p:cNvPr'), 'name') === 'T')!.id;
const resolved = (e: PptxEditorInstance) => {
  const item = e.session!.layout(0)!.items.find((x) => x.kind === 'shape' && x.element.text);
  return item?.kind === 'shape' ? item.element.text!.paragraphs : [];
};
const texts = (e: PptxEditorInstance): string[] =>
  resolved(e).map((p) => p.runs.map((r) => (r.kind === 'break' ? '\n' : r.text)).join(''));
const caret = (e: PptxEditorInstance, paragraph: number, offset: number, to = offset) =>
  e.setSelection({
    kind: 'text',
    nodeId: shapeId(e),
    anchor: { paragraph, offset },
    focus: { paragraph, offset: to },
  });

describe('paragraph and font keys', () => {
  test('Ctrl/Cmd+L, E, R, and J align the paragraph, in text and on a selected shape', async () => {
    const { e, key } = await open(para(run('One two')));
    caret(e, 0, 2);
    expect(key('e', { metaKey: true })).toBe(true);
    expect(resolved(e)[0].align).toBe('ctr');
    key('r', { ctrlKey: true });
    expect(resolved(e)[0].align).toBe('r');
    key('j', { metaKey: true });
    expect(resolved(e)[0].align).toBe('just');
    e.setSelection({ kind: 'shapes', nodeIds: [shapeId(e)] });
    key('l', { metaKey: true });
    expect(resolved(e)[0].align).toBe('l');
    // With nothing selected the keys keep their browser meaning (Cmd+R reloads).
    e.setSelection({ kind: 'none' });
    expect(key('r', { metaKey: true })).toBe(false);
    e.destroy();
  });

  test('Ctrl/Cmd+Shift+> and < step the font size; Ctrl/Cmd+] and [ step the list level', async () => {
    const { e, key } = await open(para(run('One two')));
    caret(e, 0, 0, 7);
    const size = () => resolved(e)[0].runs[0].kind === 'text' && resolved(e)[0].runs[0];
    const before = (size() as { props: { size: number } }).props.size;
    key('>', { metaKey: true, shiftKey: true });
    const bigger = (size() as { props: { size: number } }).props.size;
    expect(bigger).toBeGreaterThan(before);
    key(',', { ctrlKey: true, shiftKey: true });
    expect((size() as { props: { size: number } }).props.size).toBe(before);
    key(']', { metaKey: true });
    expect(resolved(e)[0].level).toBe(1);
    key('[', { metaKey: true });
    expect(resolved(e)[0].level).toBe(0);
    e.destroy();
  });

  test('Shift+F3 steps the case of the word at the caret: Capitalized, UPPER, lower', async () => {
    expect([nextCase('hello'), nextCase('Hello'), nextCase('HELLO')]).toEqual([
      'capitalize',
      'upper',
      'lower',
    ]);
    const { e, key } = await open(para(run('say hello now')));
    caret(e, 0, 6);
    key('F3', { shiftKey: true });
    expect(texts(e)[0]).toBe('say Hello now');
    key('F3', { shiftKey: true });
    expect(texts(e)[0]).toBe('say HELLO now');
    key('F3', { shiftKey: true });
    expect(texts(e)[0]).toBe('say hello now');
    e.destroy();
  });
});

describe('keys on a selected shape and word deletion', () => {
  test('Enter on a selected shape selects all its text; a typed key adds to its end', async () => {
    const { e, key } = await open(para(run('One')) + para(run('Two')));
    e.setSelection({ kind: 'shapes', nodeIds: [shapeId(e)] });
    key('Enter');
    expect(e.snapshot().selection).toMatchObject({
      kind: 'text',
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 1, offset: 3 },
    });
    key('Escape');
    key('Z');
    expect(texts(e)).toEqual(['One', 'TwoZ']);
    e.destroy();
  });

  test('Cmd+Backspace (a soft line deletion) deletes the word before the caret', () => {
    const kinds: TextInputKind[] = [];
    const host = { textInput: (kind: TextInputKind) => void kinds.push(kind) };
    dispatchBeforeInput(host as never, { inputType: 'deleteSoftLineBackward' } as InputEvent);
    expect(kinds).toEqual(['deleteWordBackward']);
  });
});

describe('AutoCorrect while typing', () => {
  test('the rules, one key at a time', () => {
    expect(corrections('x (c', ')')).toEqual({ edits: [{ from: 2, to: 4, text: '' }], typed: '©' });
    expect(corrections('x (R', ')').typed).toBe('®');
    expect(corrections('wait..', '.').typed).toBe('…');
    expect(corrections('a 1/2', ' ').edits).toEqual([{ from: 2, to: 5, text: '½' }]);
    expect(corrections('a 3/4', ',').edits).toEqual([{ from: 2, to: 5, text: '¾' }]);
    // Capitals, quotes, and dashes follow the system settings: no rule makes them.
    expect(corrections('a --', ' ').edits).toEqual([]);
    expect(corrections('hello', ' ').edits).toEqual([]);
    expect(corrections('Done. next', '.').edits).toEqual([]);
    expect(corrections('so i', ' ').edits).toEqual([]);
    expect(corrections('on monday', ' ').edits).toEqual([]);
    expect(corrections('say', '"')).toEqual({ edits: [], typed: '"' });
    expect(corrections('it', "'")).toEqual({ edits: [], typed: "'" });
  });

  test('typing corrects symbols, the ellipsis, and fractions, and nothing else', async () => {
    const { e, type } = await open(para(''));
    caret(e, 0, 0);
    type('hello world. this is "it" -- i think so on monday (c) (tm) wait... 1/2 ');
    expect(texts(e)[0]).toBe('hello world. this is "it" -- i think so on monday © ™ wait… ½ ');
    e.destroy();
  });

  test('a quote typed over a selection stays straight; Backspace after a later edit is a Backspace', async () => {
    const { e, type } = await open(para(run('Say word')));
    caret(e, 0, 4, 8);
    type('"');
    expect(texts(e)[0]).toBe('Say "');
    const fresh = await open(para(''));
    caret(fresh.e, 0, 0);
    fresh.type('1. ');
    expect(resolved(fresh.e)[0].bullet?.type).toBe('autonum');
    fresh.key('e', { metaKey: true });
    fresh.key('Backspace');
    // Backspace now removes the number; the alignment stays.
    expect(resolved(fresh.e)[0].align).toBe('ctr');
    expect(resolved(fresh.e)[0].bullet).toBeNull();
    fresh.e.destroy();
    e.destroy();
  });

  test('a typed list marker starts a list; Backspace right after gives the marker back', async () => {
    const { e, type, key } = await open(para('') + para(''));
    caret(e, 0, 0);
    type('1. first\nsecond ');
    expect(texts(e).slice(0, 2)).toEqual(['first', 'second ']);
    expect(resolved(e)[0].bullet?.type).toBe('autonum');
    expect(resolved(e)[1].bullet?.type).toBe('autonum');
    caret(e, 2, 0);
    type('* ');
    expect(resolved(e)[2].bullet?.type).toBe('char');
    key('Backspace');
    expect(resolved(e)[2].bullet).toBeNull();
    expect(texts(e)[2]).toBe('*');
    e.destroy();
  });

  test('1) numbers with a parenthesis, and a paragraph that has a bullet is left alone', async () => {
    const { e, type } = await open(para(''));
    caret(e, 0, 0);
    type('1) one');
    const pPr = find(root(e), 'a:pPr');
    expect(attr(find(pPr!, 'a:buAutoNum'), 'type')).toBe('arabicParenR');
    type('\n- two');
    expect(texts(e)[1]).toBe('- two');
    e.destroy();
  });

  test('autoCorrect: false types every character as it is; a paste is never corrected', async () => {
    const off = await open(para(''), { autoCorrect: false });
    caret(off.e, 0, 0);
    off.type('hello "x" (c) 1. ');
    expect(texts(off.e)[0]).toBe('hello "x" (c) 1. ');
    off.e.destroy();
    const { e } = await open(para(''));
    caret(e, 0, 0);
    e.input('text', 'pasted. text "here" ');
    expect(texts(e)[0]).toBe('pasted. text "here" ');
    e.destroy();
  });
});
