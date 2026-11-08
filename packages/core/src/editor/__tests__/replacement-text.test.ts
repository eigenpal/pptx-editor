// System text changes on the slide surface: the input holds the text before the caret, and an
// `insertReplacementText` replaces the text the system targeted in the model: a word, the
// double-space period, a correction during a typing burst, and one undo step. A target in
// another paragraph is never matched, and a keyboard that recomposes a word changes it once.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { attr, elements, find, xmlPartOf } from '../../store/index.ts';
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
    expect(kept.startsWith('word ')).toBe(true);
    expect(long.endsWith(kept)).toBe(true);
  });

  test('the input holds the text before the caret and follows the caret', async () => {
    const { e, input } = await open(para(run('Hello world')) + para(run('Second')));
    // No attribute turns off the system's own text changes.
    expect(input.hasAttribute('autocorrect')).toBe(false);
    expect(input.hasAttribute('autocapitalize')).toBe(false);
    expect(input.hasAttribute('spellcheck')).toBe(false);
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

  test('undo gives the replaced word back, selected', async () => {
    const { e, input, replace } = await open(para(run('Hello wrold')));
    caret(e, 0, 11);
    replace(6, 11, 'world');
    expect(texts(e)).toEqual(['Hello world']);
    e.exec({ type: 'undo' });
    expect(texts(e)).toEqual(['Hello wrold']);
    // Undo selects the word the replacement replaced; the input holds no text for a range.
    expect(e.snapshot().selection).toMatchObject({
      anchor: { paragraph: 0, offset: 6 },
      focus: { paragraph: 0, offset: 11 },
    });
    expect(input.value).toBe('');
    e.destroy();
  });

  test('the target comes from the input selection, or from one target range', () => {
    const event = { inputType: 'insertReplacementText', data: 'the' } as InputEvent;
    const context = { key: 'k', text: 'a teh' };
    const field = { value: 'a teh', selectionStart: 2, selectionEnd: 5 };
    expect(replacementTarget(event, field, context)).toEqual({
      text: 'teh',
      after: 0,
      paragraph: 'k',
    });
    // An input that no longer holds the context names no target.
    expect(replacementTarget(event, { ...field, value: 'other' }, context)).toBeNull();
    const text = doc.createTextNode('one teh');
    const range = { startContainer: text, startOffset: 4, endContainer: text, endOffset: 7 };
    const ranged = { ...event, getTargetRanges: () => [range] } as unknown as InputEvent;
    const collapsed = { ...field, selectionStart: 5 };
    expect(replacementTarget(ranged, collapsed, context)).toEqual({
      text: 'teh',
      after: 0,
      paragraph: null,
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
      target: { text: '', after: 0, paragraph: 'k' },
      text: '日本語',
    });
    // A keyboard recomposes the word before the caret: only the new letter goes in.
    expect(compositionEdit('a helo', 'a hello', 7, 'hello', 'k')).toEqual({
      target: { text: '', after: 1, paragraph: 'k' },
      text: 'l',
    });
    expect(compositionEdit('a teh', 'a the', 5, 'the', 'k')).toEqual({
      target: { text: 'eh', after: 0, paragraph: 'k' },
      text: 'he',
    });
    // The field does not hold the committed text at the caret.
    expect(compositionEdit('abc', 'abcに', 4, '日本', 'k')).toBeNull();
  });

  test('a recomposed word changes once in the model', async () => {
    const { e, input, fire } = await open(para(run('a helo')));
    caret(e, 0, 6);
    const composition = (type: string, data = ''): Event => {
      const event = new win.CompositionEvent(type, { bubbles: true }) as unknown as Event;
      Object.defineProperty(event, 'data', { value: data });
      return event;
    };
    fire(composition('compositionstart'));
    input.value = 'a hello';
    input.setSelectionRange(7, 7);
    fire(composition('compositionend', 'hello'));
    expect(texts(e)).toEqual(['a hello']);
    expect(focusOf(e)).toEqual({ paragraph: 0, offset: 7 });
    expect(input.value).toBe('a hello');
    e.destroy();
  });
});
