// System text changes on the slide surface: the input holds the text before the caret, and an
// `insertReplacementText` replaces the text the system targeted in the model: a word, the
// double-space period, a correction during a typing burst, and one undo step. A target in
// another paragraph is never matched, and a keyboard that recomposes a word changes it once.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import {
  buildPptx,
  cell,
  para,
  run,
  shape,
  table,
} from '../../../../../scripts/fixtures/builder.ts';
import { attr, child, elements, find, textOf, xmlPartOf } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import {
  compositionEdit,
  contextBefore,
  CONTEXT_LIMIT,
  paragraphKey,
  replacementRange,
} from '../replacement-text.ts';
import { dispatchBeforeInput, replacementTarget } from '../surface-input.ts';

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
  const fire = (event: Event): boolean => {
    input.dispatchEvent(event);
    return event.defaultPrevented;
  };
  const beforeInput = (inputType: string, data?: string) =>
    fire(
      new win.InputEvent('beforeinput', {
        inputType,
        data,
        bubbles: true,
        cancelable: true,
      }) as unknown as Event
    );
  /** Type through the input, one `beforeinput` for each character, as a keyboard does. */
  const type = (text: string): void => {
    for (const ch of text) beforeInput('insertText', ch);
  };
  /** The system replaces `[from, to)` of the input's text with `text`. */
  const replace = (from: number, to: number, text: string): boolean => {
    input.setSelectionRange(from, to);
    return beforeInput('insertReplacementText', text);
  };
  return { e, input, type, replace, beforeInput, fire };
}

const shapeId = (e: PptxEditorInstance): string => {
  const part = e.session!.info.slides[0].partName;
  const tree = find(xmlPartOf(e.session!.store.package, part)!.root, 'p:spTree')!;
  return elements(tree).find((c) => attr(find(c, 'p:cNvPr'), 'name') === 'T')!.id;
};
const texts = (e: PptxEditorInstance): string[] => {
  const item = e.session!.layout(0)!.items.find((x) => x.kind === 'shape' && x.element.text);
  const paragraphs = item?.kind === 'shape' ? item.element.text!.paragraphs : [];
  return paragraphs.map((p) => p.runs.map((r) => (r.kind === 'break' ? '\n' : r.text)).join(''));
};
const caret = (e: PptxEditorInstance, paragraph: number, offset: number) =>
  e.setSelection({
    kind: 'text',
    nodeId: shapeId(e),
    anchor: { paragraph, offset },
    focus: { paragraph, offset },
  });
const focusOf = (e: PptxEditorInstance) => {
  const sel = e.snapshot().selection;
  return sel.kind === 'text' ? sel.focus : null;
};

describe('the text the input holds for the system', () => {
  test('the caret paragraph before the caret, from its last line break, cut when long', () => {
    expect(contextBefore('Hello world', 5)).toBe('Hello');
    expect(contextBefore('one\ntwo three', 13)).toBe('two three');
    const long = `${'word '.repeat(100)}end`;
    const kept = contextBefore(long, long.length);
    expect(kept.length).toBeLessThanOrEqual(CONTEXT_LIMIT);
    expect(kept.length).toBeGreaterThanOrEqual(200);
    expect(kept.startsWith('word ')).toBe(true);
    expect(long.endsWith(kept)).toBe(true);
    // A long word keeps at least 200 characters; the cut never splits an emoji.
    const word = `${'x'.repeat(300)} word`;
    expect(contextBefore(word, word.length).length).toBeGreaterThanOrEqual(200);
    const emoji = '😀'.repeat(200);
    const cut = contextBefore(emoji, emoji.length);
    expect(cut.length).toBeGreaterThanOrEqual(200);
    expect(cut.length % 2).toBe(0);
    expect(cut.codePointAt(0)).toBe(0x1f600);
  });

  test('the input holds the text before the caret and follows the caret', async () => {
    const { e, input } = await open(para(run('Hello world')) + para(run('Second')));
    // No attribute turns off the system's own text changes.
    expect(input.hasAttribute('autocorrect')).toBe(false);
    expect(input.hasAttribute('autocapitalize')).toBe(false);
    expect(input.hasAttribute('spellcheck')).toBe(false);
    // A screen reader hears that the value is the text up to the caret.
    expect(input.getAttribute('aria-description')).toBe('canvas.inputDescription');
    caret(e, 0, 5);
    expect(input.value).toBe('Hello');
    expect([input.selectionStart, input.selectionEnd]).toEqual([5, 5]);
    caret(e, 1, 6);
    expect(input.value).toBe('Second');
    e.setSelection({ kind: 'none' });
    expect(input.value).toBe('');
    e.destroy();
  });
});

describe('a system replacement (insertReplacementText)', () => {
  test('a corrected word replaces the word in the model, once', async () => {
    const { e, input, replace } = await open(para(run('I like teh')));
    caret(e, 0, 10);
    expect(replace(7, 10, 'the')).toBe(true);
    expect(texts(e)).toEqual(['I like the']);
    expect(focusOf(e)).toEqual({ paragraph: 0, offset: 10 });
    expect(input.value).toBe('I like the');
    e.destroy();
  });

  test('a word corrected after the space that ends it keeps the caret after the space', async () => {
    const { e, replace } = await open(para(run('teh cat')));
    caret(e, 0, 4);
    replace(0, 3, 'the');
    expect(texts(e)).toEqual(['the cat']);
    expect(focusOf(e)).toEqual({ paragraph: 0, offset: 4 });
    e.destroy();
  });

  test('the double-space period replaces the space: "word. ", not "word . "', async () => {
    const { e, input, type, replace } = await open(para(''));
    caret(e, 0, 0);
    type('word ');
    expect(input.value).toBe('word ');
    replace(4, 5, '. ');
    expect(texts(e)).toEqual(['word. ']);
    expect(focusOf(e)).toEqual({ paragraph: 0, offset: 6 });
    e.destroy();
  });

  test('a capital or a curly quote from the system goes in as sent, with no AutoCorrect', async () => {
    const { e, type, replace } = await open(para(''));
    caret(e, 0, 0);
    type('hello (c]');
    replace(0, 1, 'H');
    // A typed `)` here would make ©; the system's `)` stays as sent.
    replace(8, 9, ')');
    expect(texts(e)).toEqual(['Hello (c)']);
    e.destroy();
  });

  test('a target in another paragraph is never matched', async () => {
    const { e } = await open(para(run('teh')) + para(run('teh')));
    const part = e.session!.info.slides[0].partName;
    const first = paragraphKey(part, shapeId(e), undefined, 0);
    const second = paragraphKey(part, shapeId(e), undefined, 1);
    const target = { text: 'teh', after: 0, paragraph: first };
    expect(replacementRange('teh', 3, first, target)).toEqual({ from: 0, to: 3 });
    // The same characters in the caret's paragraph are a coincidence, not the target.
    expect(replacementRange('teh', 3, second, target)).toBeNull();
    // A target the model text does not hold is never matched either.
    expect(replacementRange('tea', 3, first, target)).toBeNull();
    expect(replacementRange('eh', 2, first, target)).toBeNull();
    e.destroy();
  });

  test('a target the input does not hold goes in at the caret and edits nothing else', async () => {
    const { e, input, beforeInput } = await open(para(run('one two')));
    caret(e, 0, 7);
    // The input's caret is collapsed: the event names no target.
    expect(input.selectionStart).toBe(7);
    beforeInput('insertReplacementText', '!');
    expect(texts(e)).toEqual(['one two!']);
    e.destroy();
  });

  test('a replacement during a typing burst is one undo step after the typed text', async () => {
    const { e, type, replace } = await open(para(''));
    caret(e, 0, 0);
    type('I lik teh');
    replace(6, 9, 'the');
    type(' cat');
    expect(texts(e)).toEqual(['I lik the cat']);
    // The replacement and the typing after it are one step; the burst before it is another.
    e.exec({ type: 'undo' });
    expect(texts(e)).toEqual(['I lik teh']);
    e.exec({ type: 'undo' });
    expect(texts(e)).toEqual(['']);
    e.exec({ type: 'redo' });
    e.exec({ type: 'redo' });
    expect(texts(e)).toEqual(['I lik the cat']);
    e.destroy();
  });

  test('undo gives the replaced word back, with the caret where it was', async () => {
    const { e, input, replace } = await open(para(run('Hello wrold')));
    caret(e, 0, 11);
    replace(6, 11, 'world');
    expect(texts(e)).toEqual(['Hello world']);
    e.exec({ type: 'undo' });
    expect(texts(e)).toEqual(['Hello wrold']);
    expect(e.snapshot().selection).toMatchObject({
      anchor: { paragraph: 0, offset: 11 },
      focus: { paragraph: 0, offset: 11 },
    });
    expect(input.value).toBe('Hello wrold');
    e.destroy();
  });

  test('a target the model text does not hold is refused, so no word goes in twice', async () => {
    const { e, fire } = await open(para(run('say teh ')));
    caret(e, 0, 8);
    const revision = e.session!.store.revision;
    // A browser names a target range whose text is not before the caret in the model.
    const node = doc.createTextNode('teh');
    const range = { startContainer: node, startOffset: 0, endContainer: node, endOffset: 3 };
    const event = new win.InputEvent('beforeinput', {
      inputType: 'insertReplacementText',
      data: 'the',
      bubbles: true,
      cancelable: true,
    }) as unknown as Event;
    Object.defineProperty(event, 'getTargetRanges', { value: () => [range] });
    fire(event);
    expect(texts(e)).toEqual(['say teh ']);
    expect(e.snapshot().lastRejection?.reason).toBe('errors.inputUnsupported');
    expect(e.session!.store.revision).toBe(revision);
    e.destroy();
  });

  test('a replacement with a caret format armed replaces the word and keeps the format', async () => {
    const { e, type, replace } = await open(para(''));
    caret(e, 0, 0);
    type('say teh ');
    e.exec({ type: 'toggleBold' });
    replace(4, 7, 'the');
    expect(texts(e)).toEqual(['say the ']);
    type('x');
    expect(texts(e)).toEqual(['say the x']);
    const runs = e.session!.layout(0)!.items.find((x) => x.kind === 'shape' && x.element.text);
    const last = runs?.kind === 'shape' ? runs.element.text!.paragraphs[0].runs.at(-1) : null;
    expect(last && 'props' in last && last.text).toBe('x');
    expect((last as { props: { bold: boolean } }).props.bold).toBe(true);
    e.destroy();
  });

  test('a selection that a native key left in the input is never typed over', async () => {
    const { e, input, type, fire, beforeInput } = await open(para(''));
    caret(e, 0, 0);
    type('one two three');
    // A line-wise native key (Ctrl+Shift+P) selects in the input; the model caret stays.
    fire(
      new win.KeyboardEvent('keydown', {
        key: 'P',
        ctrlKey: true,
        shiftKey: true,
        bubbles: true,
      }) as unknown as Event
    );
    input.setSelectionRange(12, 13);
    fire(new win.KeyboardEvent('keydown', { key: 'X', bubbles: true }) as unknown as Event);
    expect([input.selectionStart, input.selectionEnd]).toEqual([13, 13]);
    beforeInput('insertText', 'X');
    expect(texts(e)).toEqual(['one two threeX']);
    e.destroy();
  });

  test('an input event with no beforeinput leaves no selection to type over', async () => {
    const { e, input, type, fire } = await open(para(''));
    caret(e, 0, 0);
    type('say teh ');
    input.setSelectionRange(4, 7);
    fire(
      new win.InputEvent('input', {
        inputType: 'insertReplacementText',
        data: 'the',
        bubbles: true,
      }) as unknown as Event
    );
    expect([input.selectionStart, input.selectionEnd]).toEqual([8, 8]);
    type('x');
    expect(texts(e)).toEqual(['say teh x']);
    e.destroy();
  });

  test('a field that something else wrote goes to the model', async () => {
    const { e, input, type, fire } = await open(para(''));
    caret(e, 0, 0);
    type('say teh ');
    // An extension writes the field and sends a plain input event.
    input.value = 'say the ';
    fire(
      new win.InputEvent('input', { inputType: 'insertText', bubbles: true }) as unknown as Event
    );
    expect(texts(e)).toEqual(['say the ']);
    expect(focusOf(e)).toEqual({ paragraph: 0, offset: 8 });
    expect(input.value).toBe('say the ');
    e.destroy();
  });

  for (const [label, data] of [
    ['an emoji picker', '😀'],
    ['dictation', ' and more'],
  ] as const) {
    test(`${label} with no key never types over a selection a native key left`, async () => {
      const { e, input, type, fire, beforeInput } = await open(para(''));
      caret(e, 0, 0);
      type('hello world');
      // A native key (Shift+PageUp) selects in the input; the model caret stays.
      input.focus();
      input.setSelectionRange(0, 11);
      fire(new win.Event('selectionchange') as unknown as Event);
      expect([input.selectionStart, input.selectionEnd]).toEqual([11, 11]);
      // The text arrives with no keydown first.
      beforeInput('insertText', data);
      expect(texts(e)).toEqual([`hello world${data}`]);
      e.destroy();
    });
  }

  test('an input that cannot be canceled goes to the model once, by the field', async () => {
    const { e, input, type, fire } = await open(para(''));
    caret(e, 0, 0);
    type('ab');
    const event = new win.InputEvent('beforeinput', {
      inputType: 'insertText',
      data: 'c',
      bubbles: true,
      cancelable: false,
    }) as unknown as Event;
    fire(event);
    expect(texts(e)).toEqual(['ab']);
    // The browser applies the text to the field; the input event reports it.
    input.value = 'abc';
    fire(
      new win.InputEvent('input', {
        inputType: 'insertText',
        data: 'c',
        bubbles: true,
      }) as unknown as Event
    );
    expect(texts(e)).toEqual(['abc']);
    e.destroy();
  });

  test('a field change of one emoji to another replaces the whole emoji', async () => {
    const { e, input, fire, beforeInput } = await open(para(''));
    caret(e, 0, 0);
    beforeInput('insertText', 'ok 😀');
    input.value = 'ok 😃';
    fire(
      new win.InputEvent('input', { inputType: 'insertText', bubbles: true }) as unknown as Event
    );
    expect(texts(e)).toEqual(['ok 😃']);
    expect(focusOf(e)).toEqual({ paragraph: 0, offset: 5 });
    e.destroy();
  });

  test('a selection over a field that is not the synced text is refused', async () => {
    const { e, input, type, replace } = await open(para(''));
    caret(e, 0, 0);
    type('say teh ');
    const revision = e.session!.store.revision;
    input.value = 'say teh  ';
    replace(4, 7, 'the');
    expect(texts(e)).toEqual(['say teh ']);
    expect(e.snapshot().lastRejection?.reason).toBe('errors.inputUnsupported');
    expect(e.session!.store.revision).toBe(revision);
    e.destroy();
  });

  test('a replacement in a table cell replaces the word in that cell only', async () => {
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const rows = ['teh', 'teh cat']
      .map((t) => `<a:tr h="370840">${cell(t)}${cell(t)}</a:tr>`)
      .join('');
    const bytes = buildPptx({
      slides: [
        {
          shapes: table({
            id: 3,
            x: IN,
            y: IN,
            cx: 2 * IN,
            cy: 741_680,
            columns: [IN, IN],
            tblPr: '<a:tblPr/>',
            rows,
          }),
        },
      ],
    });
    expect((await e.load(bytes)).ok).toBe(true);
    const host = doc.createElement('div') as unknown as HTMLElement;
    doc.body.append(host as unknown as Node);
    e.attach(host);
    const input = host.querySelector('textarea') as HTMLTextAreaElement;
    const part = e.session!.info.slides[0].partName;
    const tree = find(xmlPartOf(e.session!.store.package, part)!.root, 'p:spTree')!;
    const nodeId = elements(tree).find((c) => c.local === 'graphicFrame')!.id;
    const at = { paragraph: 0, offset: 4 };
    e.setSelection({ kind: 'text', nodeId, cell: { row: 1, column: 1 }, anchor: at, focus: at });
    expect(input.value).toBe('teh ');
    input.setSelectionRange(0, 3);
    input.dispatchEvent(
      new win.InputEvent('beforeinput', {
        inputType: 'insertReplacementText',
        data: 'the',
        bubbles: true,
        cancelable: true,
      }) as unknown as Event
    );
    const cells = (): string[] => {
      const tbl = find(xmlPartOf(e.session!.store.package, part)!.root, 'a:tbl')!;
      return elements(tbl)
        .filter((r) => r.local === 'tr')
        .flatMap((r) =>
          elements(r)
            .filter((c) => c.local === 'tc')
            .map((c) => textOf(child(c, 'a:txBody')))
        );
    };
    expect(cells()).toEqual(['teh', 'teh', 'teh cat', 'the cat']);
    expect(paragraphKey(part, nodeId, { row: 1, column: 1 }, 0)).not.toBe(
      paragraphKey(part, nodeId, { row: 1, column: 0 }, 0)
    );
    e.destroy();
  });

  test('insertText over the text the system selected in the input replaces it', async () => {
    const { e, input, type, fire } = await open(para(''));
    caret(e, 0, 0);
    type('cats adn ');
    input.setSelectionRange(5, 8);
    fire(
      new win.InputEvent('beforeinput', {
        inputType: 'insertText',
        data: 'and',
        bubbles: true,
        cancelable: true,
      }) as unknown as Event
    );
    expect(texts(e)).toEqual(['cats and ']);
    expect(focusOf(e)).toEqual({ paragraph: 0, offset: 9 });
    e.destroy();
  });

  test('the target comes from the input selection, or from one target range', () => {
    const event = { inputType: 'insertReplacementText', data: 'the' } as InputEvent;
    const context = { key: 'k', text: 'a teh' };
    const field = { value: 'a teh', selectionStart: 2, selectionEnd: 5 };
    expect(replacementTarget(event, field, context)).toEqual({
      text: 'teh',
      after: 0,
      between: '',
      paragraph: 'k',
    });
    // A selection over a value that is not the context cannot be used: it never matches.
    expect(replacementTarget(event, { ...field, value: 'other' }, context)).toEqual({
      text: '',
      after: -1,
      paragraph: null,
    });
    const text = doc.createTextNode('one teh');
    const range = { startContainer: text, startOffset: 4, endContainer: text, endOffset: 7 };
    const ranged = { ...event, getTargetRanges: () => [range] } as unknown as InputEvent;
    const collapsed = { ...field, selectionStart: 5 };
    expect(replacementTarget(ranged, collapsed, context)).toEqual({
      text: 'teh',
      after: 0,
      paragraph: null,
    });
    // A range in the input's own text: its offsets give the distance to the caret.
    const own = doc.createTextNode('a teh');
    const inInput = { startContainer: own, startOffset: 2, endContainer: own, endOffset: 4 };
    const early = { ...event, getTargetRanges: () => [inInput] } as unknown as InputEvent;
    expect(replacementTarget(early, collapsed, context)).toEqual({
      text: 'te',
      after: 1,
      between: 'h',
      paragraph: 'k',
    });
    const two = { ...event, getTargetRanges: () => [range, range] } as unknown as InputEvent;
    expect(replacementTarget(two, collapsed, context)).toBeNull();
    const calls: unknown[] = [];
    const host = { replaceText: (...args: unknown[]) => calls.push(args), textInput: () => {} };
    dispatchBeforeInput(host as never, event, { text: 'teh', after: 0, paragraph: 'k' });
    expect(calls).toEqual([[{ text: 'teh', after: 0, paragraph: 'k' }, 'the']]);
  });
});

describe('a composition over the text the input holds', () => {
  test('the change against the context keeps the shared characters', () => {
    // A plain composition after the context.
    expect(compositionEdit('abc ', 'abc 日本語', 7, '日本語', 'k')).toEqual({
      target: { text: '', after: 0, between: '', paragraph: 'k' },
      text: '日本語',
    });
    // A keyboard recomposes the word before the caret: only the new letter goes in.
    expect(compositionEdit('a helo', 'a hello', 7, 'hello', 'k')).toEqual({
      target: { text: '', after: 1, between: 'o', paragraph: 'k' },
      text: 'l',
    });
    expect(compositionEdit('a teh', 'a the', 5, 'the', 'k')).toEqual({
      target: { text: 'eh', after: 0, between: '', paragraph: 'k' },
      text: 'he',
    });
    // An empty commit deletes what the composition covered; an unchanged one changes nothing.
    expect(compositionEdit('a hello', 'a ', 2, '', 'k')).toEqual({
      target: { text: 'hello', after: 0, between: '', paragraph: 'k' },
      text: '',
    });
    expect(compositionEdit('a hello', 'a hello', 7, 'hello', 'k')).toEqual({
      target: { text: '', after: 0, between: '', paragraph: 'k' },
      text: '',
    });
    // An emoji recomposed to another is one change of the whole emoji.
    expect(compositionEdit('ok 😀', 'ok 😃', 5, '😃', 'k')).toEqual({
      target: { text: '😀', after: 0, between: '', paragraph: 'k' },
      text: '😃',
    });
    // The field does not hold the committed text at the caret.
    expect(compositionEdit('abc', 'abcに', 4, '日本', 'k')).toBeNull();
  });

  const composition = (type: string, data = ''): Event => {
    const event = new win.CompositionEvent(type, { bubbles: true }) as unknown as Event;
    Object.defineProperty(event, 'data', { value: data });
    return event;
  };
  /** A keyboard composes the input's text again: the field becomes `value`, then commits. */
  async function recompose(text: string, value: string, data: string) {
    const opened = await open(para(run(text)));
    caret(opened.e, 0, text.length);
    const revision = opened.e.session!.store.revision;
    opened.fire(composition('compositionstart'));
    opened.input.value = value;
    opened.input.setSelectionRange(value.length, value.length);
    opened.fire(composition('compositionend', data));
    return { ...opened, revision };
  }

  test('a recomposed word changes once in the model', async () => {
    const { e, input } = await recompose('a helo', 'a hello', 'hello');
    expect(texts(e)).toEqual(['a hello']);
    expect(focusOf(e)).toEqual({ paragraph: 0, offset: 7 });
    expect(input.value).toBe('a hello');
    e.destroy();
  });

  test('a recomposed word that the user deletes goes from the model', async () => {
    const { e, input } = await recompose('a hello', 'a ', '');
    expect(texts(e)).toEqual(['a ']);
    expect(focusOf(e)).toEqual({ paragraph: 0, offset: 2 });
    expect(input.value).toBe('a ');
    e.destroy();
  });

  /** A simulated IME: the first key (229), then a composition at the input's selection. */
  function ime(fire: (e: Event) => boolean, input: HTMLTextAreaElement, data: string, key = true) {
    if (key)
      fire(
        new win.KeyboardEvent('keydown', {
          key: 'Process',
          keyCode: 229,
          bubbles: true,
        }) as unknown as Event
      );
    fire(composition('compositionstart'));
    const { selectionStart: s, selectionEnd: t, value } = input;
    input.value = value.slice(0, s) + data + value.slice(t);
    input.setSelectionRange(s + data.length, s + data.length);
    fire(composition('compositionend', data));
  }

  for (const [label, from, to, key] of [
    ['a stale selection, then an IME key', 0, 11, true],
    ['a stale caret, then an IME key', 0, 0, true],
    ['a stale caret, then a composition with no key', 0, 0, false],
  ] as const) {
    test(`${label}: the composition goes in at the model caret`, async () => {
      const { e, input, type, fire } = await open(para(''));
      caret(e, 0, 0);
      type('hello world');
      // A native key the editor does not handle (Ctrl+P, Ctrl+Shift+P) moved the input caret.
      input.setSelectionRange(from, to);
      ime(fire, input, '日', key);
      expect(texts(e)).toEqual(['hello world日']);
      expect(focusOf(e)).toEqual({ paragraph: 0, offset: 12 });
      e.destroy();
    });
  }

  test('a recomposed word whose target a peer changed keeps the typed text', async () => {
    const { e, input, fire } = await open(para(run('a helo')));
    caret(e, 0, 6);
    fire(composition('compositionstart'));
    // A peer's edit lands while the keyboard composes: the model is no longer the field.
    e.input('text', '!');
    input.value = 'a hello';
    input.setSelectionRange(7, 7);
    fire(composition('compositionend', 'hello'));
    // Nothing is lost and no letter goes in twice: the new letter follows the caret.
    expect(texts(e)).toEqual(['a helo!l']);
    e.destroy();
  });

  test('a recomposition that ends unchanged writes nothing', async () => {
    const { e, revision } = await recompose('a hello', 'a hello', 'hello');
    expect(texts(e)).toEqual(['a hello']);
    expect(e.session!.store.revision).toBe(revision);
    e.destroy();
  });
});
