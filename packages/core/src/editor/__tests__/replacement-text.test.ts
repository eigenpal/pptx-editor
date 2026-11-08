// System text changes on the slide surface: the input holds the text before the caret, and an
// `insertReplacementText` replaces the text the system targeted in the model: a word, the
// double-space period, a correction during a typing burst, and one undo step. A target in
// another paragraph is never matched, and a keyboard that recomposes a word changes it once.

import { afterAll, describe, expect, test } from 'bun:test';
import { buildPptx, cell, para, run, table } from '../../../../../scripts/fixtures/builder.ts';
import { child, elements, find, textOf, xmlPartOf } from '../../store/index.ts';
import { createPptxEditor } from '../index.ts';
import {
  compositionEdit,
  contextBefore,
  CONTEXT_LIMIT,
  paragraphKey,
  replacementRange,
} from '../replacement-text.ts';
import { COMPOSING_CLASS, dispatchBeforeInput, replacementTarget } from '../surface-input.ts';
import { caret, doc, focusOf, IN, open, shapeId, texts, win } from './surface-text-fixture.ts';

afterAll(() => win.happyDOM.close());

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

  test('a replacement a task after its selection never repeats the word', async () => {
    const { e, input, type, fire, beforeInput } = await open(para(''));
    caret(e, 0, 0);
    type('say teh');
    input.focus();
    input.setSelectionRange(4, 7);
    fire(new win.Event('selectionchange') as unknown as Event);
    const revision = e.session!.store.revision;
    beforeInput('insertReplacementText', 'the');
    expect(texts(e)).toEqual(['say teh']);
    expect(e.snapshot().lastRejection?.reason).toBe('errors.inputUnsupported');
    expect(e.session!.store.revision).toBe(revision);
    // After the space that ends the word too.
    type(' ');
    beforeInput('insertReplacementText', 'the');
    expect(texts(e)).toEqual(['say teh ']);
    e.destroy();
  });

  test('a replacement with no target that is no single word goes in at the caret', async () => {
    const { e, type, beforeInput } = await open(para(''));
    caret(e, 0, 0);
    type('say');
    beforeInput('insertReplacementText', ' hello there');
    beforeInput('insertReplacementText', '!');
    expect(texts(e)).toEqual(['say hello there!']);
    type(' ');
    beforeInput('insertReplacementText', '😀');
    expect(texts(e)).toEqual(['say hello there! 😀']);
    // A number is no correction of a word: it goes in too.
    type(' at');
    beforeInput('insertReplacementText', '5');
    expect(texts(e)).toEqual(['say hello there! 😀 at5']);
    e.destroy();
  });

  const uncancelable = (inputType: string, data?: string): Event =>
    new win.InputEvent('beforeinput', {
      inputType,
      data,
      bubbles: true,
      cancelable: false,
    }) as unknown as Event;

  test('a delete that cannot be canceled over a selected range goes to the model', async () => {
    const { e, type, fire } = await open(para(''));
    caret(e, 0, 0);
    type('hello world');
    const at = (offset: number) => ({ paragraph: 0, offset });
    e.setSelection({ kind: 'text', nodeId: shapeId(e), anchor: at(5), focus: at(11) });
    // The input holds no text for a range, so no input event follows.
    fire(uncancelable('deleteContentBackward'));
    expect(texts(e)).toEqual(['hello']);
    e.destroy();
  });

  test('a delete that cannot be canceled at a paragraph start joins the paragraphs', async () => {
    const { e, input, type, fire, beforeInput } = await open(para(''));
    caret(e, 0, 0);
    type('one');
    beforeInput('insertParagraph');
    type('two');
    caret(e, 1, 0);
    expect(input.value).toBe('');
    fire(uncancelable('deleteContentBackward'));
    expect(texts(e)).toEqual(['onetwo']);
    expect(focusOf(e)).toEqual({ paragraph: 0, offset: 3 });
    e.destroy();
  });

  test('an Enter that cannot be canceled splits once, and its field change is no edit', async () => {
    const { e, input, type, fire } = await open(para(''));
    caret(e, 0, 0);
    type('one');
    fire(uncancelable('insertParagraph'));
    // The browser writes a line break into the field, after the model split.
    input.value += '\n';
    fire(
      new win.InputEvent('input', {
        inputType: 'insertParagraph',
        bubbles: true,
      }) as unknown as Event
    );
    expect(texts(e)).toEqual(['one', '']);
    expect(focusOf(e)).toEqual({ paragraph: 1, offset: 0 });
    expect(input.value).toBe('');
    e.destroy();
  });

  test('a field change after a later key is an edit, also when no echo came', async () => {
    const { e, input, type, fire } = await open(para(''));
    caret(e, 0, 0);
    type('ab');
    // A non-cancelable insert with no field change: no input event follows it.
    e.setSelection({
      kind: 'text',
      nodeId: shapeId(e),
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 2 },
    });
    fire(uncancelable('insertText', 'x'));
    expect(texts(e)).toEqual(['x']);
    type('y');
    fire(uncancelable('insertText', 'z'));
    input.value += 'z';
    fire(
      new win.InputEvent('input', {
        inputType: 'insertText',
        data: 'z',
        bubbles: true,
      }) as unknown as Event
    );
    expect(texts(e)).toEqual(['xyz']);
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
      // The browser reports the move a task later; the composition comes after that.
      input.focus();
      input.setSelectionRange(from, to);
      fire(new win.Event('selectionchange') as unknown as Event);
      ime(fire, input, '日', key);
      expect(texts(e)).toEqual(['hello world日']);
      expect(focusOf(e)).toEqual({ paragraph: 0, offset: 12 });
      e.destroy();
    });
  }

  test('a reconversion composes the word the system selected again, in place', async () => {
    const { e, input, type, fire } = await open(para(''));
    caret(e, 0, 0);
    type('hello world');
    // The IME selects the word and starts composing in the same task.
    input.setSelectionRange(6, 11);
    ime(fire, input, 'WORLD', false);
    expect(texts(e)).toEqual(['hello WORLD']);
    expect(focusOf(e)).toEqual({ paragraph: 0, offset: 11 });
    e.destroy();
  });

  test('an edit at the place of a composition ends it, and its text goes in nowhere', async () => {
    const { e, input, fire } = await open(para(run('a helo')));
    caret(e, 0, 6);
    fire(composition('compositionstart'));
    input.value = 'a hel';
    // An edit at the place moves the caret in the paragraph: the place no longer holds.
    e.input('text', '!');
    expect(input.classList.contains(COMPOSING_CLASS)).toBe(false);
    expect(e.snapshot().lastRejection?.reason).toBe('errors.compositionChanged');
    // The browser's late end adds nothing, and no letter goes in twice.
    input.value = 'a hello';
    input.setSelectionRange(7, 7);
    fire(composition('compositionend', 'hello'));
    expect(texts(e)).toEqual(['a helo!']);
    e.destroy();
  });

  /** A keyboard composes the word over `[from, to)` again, as an Android keyboard does. */
  async function composeWord(text: string, from: number, to: number) {
    const opened = await open(para(''));
    caret(opened.e, 0, 0);
    opened.type(text);
    // The keyboard selects the word and starts composing in the same task.
    opened.input.setSelectionRange(from, to);
    opened.fire(composition('compositionstart'));
    return opened;
  }
  const key = (k: string, keyCode = 0, isComposing = false): Event =>
    new win.KeyboardEvent('keydown', {
      key: k,
      keyCode,
      isComposing,
      bubbles: true,
    }) as unknown as Event;

  test('a caret move during a composition ends it at its own place', async () => {
    const { e, input, fire } = await composeWord('hello world', 6, 11);
    // A tap puts the caret at 3; the keyboard has not changed the word.
    caret(e, 0, 3);
    expect(input.classList.contains(COMPOSING_CLASS)).toBe(false);
    expect(input.value).toBe('hel');
    // The old composition's end adds nothing.
    input.value = 'helworl';
    fire(composition('compositionend', 'worl'));
    expect(texts(e)).toEqual(['hello world']);
    // The next key edits at the new caret.
    fire(key('Backspace'));
    fire(
      new win.InputEvent('beforeinput', {
        inputType: 'deleteContentBackward',
        bubbles: true,
        cancelable: true,
      }) as unknown as Event
    );
    expect(texts(e)).toEqual(['helo world']);
    expect(focusOf(e)).toEqual({ paragraph: 0, offset: 2 });
    e.destroy();
  });

  test('text composed before a caret move goes in at the composition, once', async () => {
    const { e, input, type, fire } = await open(para(''));
    caret(e, 0, 0);
    type('abc ');
    fire(composition('compositionstart'));
    input.value = 'abc にほ';
    input.setSelectionRange(6, 6);
    caret(e, 0, 1);
    expect(texts(e)).toEqual(['abc にほ']);
    expect(focusOf(e)).toEqual({ paragraph: 0, offset: 1 });
    fire(composition('compositionend', 'にほ'));
    expect(texts(e)).toEqual(['abc にほ']);
    // A caret after the composition moves with its text. The next composition comes with
    // its first key, so it is a new one at the new caret.
    fire(key('Process', 229));
    fire(composition('compositionstart'));
    input.value = 'aX';
    input.setSelectionRange(2, 2);
    caret(e, 0, 6);
    expect(texts(e)).toEqual(['aXbc にほ']);
    expect(focusOf(e)).toEqual({ paragraph: 0, offset: 7 });
    e.destroy();
  });

  test('the final text the browser sends when the input blurs goes in once, at the place', async () => {
    const { e, input, type, fire } = await open(para(''));
    caret(e, 0, 0);
    type('abc ');
    input.focus();
    fire(composition('compositionstart'));
    input.value = 'abc にほ';
    input.setSelectionRange(6, 6);
    // A browser ends its composition at a blur, with the final text, in the same task.
    const end = () => {
      input.value = 'abc にほん。';
      input.setSelectionRange(8, 8);
      fire(composition('compositionend', 'にほん。'));
    };
    input.addEventListener('blur', end, { once: true });
    caret(e, 0, 1);
    expect(texts(e)).toEqual(['abc にほん。']);
    expect(focusOf(e)).toEqual({ paragraph: 0, offset: 1 });
    expect(doc.activeElement).toBe(input as unknown as Element);
    expect(input.value).toBe('a');
    e.destroy();
  });

  test('a composition after a tap is new, also over characters the input holds', async () => {
    const { e, input, type, fire } = await open(para(''));
    caret(e, 0, 0);
    type('abc ');
    fire(composition('compositionstart'));
    input.value = 'abc wor';
    input.setSelectionRange(7, 7);
    // A tap after "abc": the composed "wor" stays at its place.
    caret(e, 0, 3);
    expect(texts(e)).toEqual(['abc wor']);
    // The keyboard composes "abc" again from a collapsed caret and adds "d".
    fire(composition('compositionstart'));
    input.value = 'abcd';
    input.setSelectionRange(4, 4);
    fire(composition('compositionend', 'abcd'));
    expect(texts(e)).toEqual(['abcd wor']);
    expect(focusOf(e)).toEqual({ paragraph: 0, offset: 4 });
    e.destroy();
  });

  /** A composition of `composed` at the end of `first`, which a tap into `second` ends. */
  async function tapAway(first: string, second: string, composed: string, offset: number) {
    const opened = await open(para(run(first)) + para(run(second)));
    caret(opened.e, 0, first.length);
    opened.fire(key('Process', 229));
    opened.fire(composition('compositionstart'));
    opened.input.value = first + composed;
    opened.input.setSelectionRange(opened.input.value.length, opened.input.value.length);
    caret(opened.e, 1, offset);
    return opened;
  }

  test('after a tap ends a composition, all new input goes in at the new caret', async () => {
    // A second tap, then dictation with no key first.
    const dictated = await tapAway('A ', 'B x', 'th', 3);
    caret(dictated.e, 1, 1);
    dictated.beforeInput('insertText', 'to');
    expect(texts(dictated.e)).toEqual(['A th', 'Bto x']);
    dictated.e.destroy();
    // A new composition with no key first, whose text continues the ended one.
    const next = await tapAway('A ', 'B ', 'th', 2);
    next.fire(composition('compositionstart'));
    next.input.value = 'B the';
    next.input.setSelectionRange(5, 5);
    next.fire(composition('compositionend', 'the'));
    expect(texts(next.e)).toEqual(['A th', 'B the']);
    next.e.destroy();
  });

  test('a reconversion at the new caret changes the word there', async () => {
    const { e, input, fire } = await tapAway('A ', 'B it', 'is', 4);
    input.setSelectionRange(2, 4);
    fire(composition('compositionstart'));
    input.value = 'B its';
    input.setSelectionRange(5, 5);
    fire(composition('compositionend', 'its'));
    expect(texts(e)).toEqual(['A is', 'B its']);
    e.destroy();
  });

  test('input after a composition ended is new input at the caret', async () => {
    const { e, beforeInput } = await composeWord('hello wrold', 6, 11);
    caret(e, 0, 3);
    // No late window: input with no key goes in at the caret, as in a plain text field.
    beforeInput('insertText', 'x');
    expect(texts(e)).toEqual(['helxlo wrold']);
    e.destroy();
  });

  test('a slide change during a composition puts its text on its own slide', async () => {
    const { e, input, type, fire } = await open(para(''), { blankSlides: 1 });
    caret(e, 0, 0);
    type('abc ');
    fire(composition('compositionstart'));
    input.value = 'abc にほ';
    input.setSelectionRange(6, 6);
    e.goToSlide(1);
    expect(input.classList.contains(COMPOSING_CLASS)).toBe(false);
    // The end that the browser sends later adds nothing.
    fire(composition('compositionend', 'にほん'));
    e.goToSlide(0);
    expect(texts(e)).toEqual(['abc にほ']);
    e.destroy();
  });

  test('an undo during a composition ends the composition first', async () => {
    const { e, input, type, fire } = await open(para(''));
    caret(e, 0, 0);
    type('abc');
    e.exec({ type: 'undo' }); // A boundary: the next text is a step of its own.
    e.exec({ type: 'redo' });
    fire(composition('compositionstart'));
    input.value = 'abcにほ';
    input.setSelectionRange(5, 5);
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(input.classList.contains(COMPOSING_CLASS)).toBe(false);
    // The composed text went in as its own step, and the undo took it back.
    expect(texts(e)).toEqual(['abc']);
    expect(e.exec({ type: 'redo' }).ok).toBe(true);
    expect(texts(e)).toEqual(['abcにほ']);
    e.destroy();
  });

  test('after a key, input goes in at the new caret', async () => {
    const { e, fire, beforeInput } = await composeWord('hello world', 6, 11);
    caret(e, 0, 3);
    fire(key('x'));
    beforeInput('insertText', 'x');
    expect(texts(e)).toEqual(['helxlo world']);
    e.destroy();
  });

  test('candidate keys during a composition do not end it', async () => {
    const { e, input, type, fire } = await open(para(''));
    caret(e, 0, 0);
    type('abc ');
    fire(composition('compositionstart'));
    input.value = 'abc にほん';
    for (const k of ['ArrowDown', 'ArrowDown', 'ArrowLeft']) fire(key(k, 229, true));
    expect(input.classList.contains(COMPOSING_CLASS)).toBe(true);
    input.value = 'abc 日本';
    input.setSelectionRange(6, 6);
    fire(composition('compositionend', '日本'));
    expect(texts(e)).toEqual(['abc 日本']);
    e.destroy();
  });

  test('an edit at the caret ends a recomposition, and the word stays', async () => {
    const { e, input, fire } = await composeWord('hello wrold', 6, 11);
    e.input('text', '!');
    expect(input.classList.contains(COMPOSING_CLASS)).toBe(false);
    input.value = 'hello world';
    input.setSelectionRange(11, 11);
    fire(composition('compositionend', 'world'));
    expect(texts(e)).toEqual(['hello wrold!']);
    e.destroy();
  });

  test('a composition clears a wait for the field change of an earlier edit', async () => {
    const { e, input, type, fire } = await open(para(''));
    caret(e, 0, 0);
    type('ab');
    // A non-cancelable insert over a range goes to the model at once.
    const at = (offset: number) => ({ paragraph: 0, offset });
    e.setSelection({ kind: 'text', nodeId: shapeId(e), anchor: at(0), focus: at(2) });
    fire(
      new win.InputEvent('beforeinput', {
        inputType: 'insertText',
        data: 'c',
        bubbles: true,
        cancelable: false,
      }) as unknown as Event
    );
    fire(composition('compositionstart'));
    input.value = 'cd';
    input.setSelectionRange(2, 2);
    fire(composition('compositionend', 'd'));
    expect(texts(e)).toEqual(['cd']);
    // An extension writes the field: that is an edit, not the earlier insert's own change.
    input.value = 'cde';
    fire(
      new win.InputEvent('input', {
        inputType: 'insertText',
        data: 'e',
        bubbles: true,
      }) as unknown as Event
    );
    expect(texts(e)).toEqual(['cde']);
    e.destroy();
  });

  test('a recomposition that ends unchanged writes nothing', async () => {
    const { e, revision } = await recompose('a hello', 'a hello', 'hello');
    expect(texts(e)).toEqual(['a hello']);
    expect(e.session!.store.revision).toBe(revision);
    e.destroy();
  });
});
