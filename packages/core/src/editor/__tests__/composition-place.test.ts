// The place of an IME composition on the slide surface, with no collaboration session: a
// composition over a selected range writes only over that range, and a host's store edit
// during a composition ends it. Text goes in only at a place that the check proves.

import { afterAll, describe, expect, test } from 'bun:test';
import { para, run } from '../../../../../scripts/fixtures/builder.ts';
import { xmlPartOf, type Element } from '../../store/index.ts';
import type { PptxEditorInstance } from '../index.ts';
import { COMPOSING_CLASS } from '../surface-input.ts';
import { caret, focusOf, open, shapeId, texts, win } from './surface-text-fixture.ts';

afterAll(() => win.happyDOM.close());

const composition = (type: string, data = ''): Event => {
  const event = new win.CompositionEvent(type, { bubbles: true }) as unknown as Event;
  Object.defineProperty(event, 'data', { value: data });
  return event;
};
const range = (e: PptxEditorInstance, paragraph: number, from: number, to: number) =>
  e.setSelection({
    kind: 'text',
    nodeId: shapeId(e),
    anchor: { paragraph, offset: from },
    focus: { paragraph, offset: to },
  });
const rejection = (e: PptxEditorInstance) => e.snapshot().lastRejection?.reason ?? null;

/** A host's store edit: the first text node of the slide gets `text`. */
function hostSetsText(e: PptxEditorInstance, text: string): void {
  const part = e.session!.info.slides[0]!.partName;
  type Node = Element | { type: 'text' };
  const first = (n: Node): Element | null =>
    n.type !== 'element'
      ? null
      : n.local === 't'
        ? n
        : (n.children.map((c) => first(c as Node)).find(Boolean) ?? null);
  const t = first(xmlPartOf(e.session!.store.package, part)!.root as Node)!;
  expect(e.session!.store.transact([{ op: 'setText', part, node: t.id, text }]).ok).toBe(true);
}

describe('a composition over a selected range', () => {
  test('writes over the range it started on', async () => {
    const { e, input, fire } = await open(para(run('alpha beta')));
    range(e, 0, 0, 5);
    fire(composition('compositionstart'));
    input.value = 'X';
    fire(composition('compositionend', 'X'));
    expect(texts(e)).toEqual(['X beta']);
    expect(rejection(e)).toBeNull();
    e.destroy();
  });

  test('a move to another range ends it, and its text is refused', async () => {
    const { e, input, fire } = await open(para(run('alpha beta')) + para(run('gamma delta')));
    range(e, 0, 0, 5);
    fire(composition('compositionstart'));
    input.value = 'X';
    range(e, 1, 6, 11);
    expect(input.classList.contains(COMPOSING_CLASS)).toBe(false);
    fire(composition('compositionend', 'X'));
    expect(texts(e)).toEqual(['alpha beta', 'gamma delta']);
    expect(rejection(e)).toBe('errors.compositionChanged');
    e.destroy();
  });

  test('a range that grows, as with Shift+click, ends it, and its text is refused', async () => {
    const { e, input, fire } = await open(para(run('alpha beta')));
    range(e, 0, 0, 5);
    fire(composition('compositionstart'));
    input.value = 'X';
    range(e, 0, 0, 10);
    fire(composition('compositionend', 'X'));
    expect(texts(e)).toEqual(['alpha beta']);
    expect(rejection(e)).toBe('errors.compositionChanged');
    e.destroy();
  });

  test('a move to a caret elsewhere ends it, and its text is refused', async () => {
    const { e, input, fire } = await open(para(run('alpha beta')) + para(run('gamma delta')));
    range(e, 0, 0, 5);
    fire(composition('compositionstart'));
    input.value = 'X';
    caret(e, 1, 3);
    fire(composition('compositionend', 'X'));
    expect(texts(e)).toEqual(['alpha beta', 'gamma delta']);
    expect(rejection(e)).toBe('errors.compositionChanged');
    e.destroy();
  });

  test('an end with no final text after the move is refused too, never silent', async () => {
    const { e, input, fire } = await open(para(run('alpha beta')) + para(run('gamma delta')));
    range(e, 0, 0, 5);
    fire(composition('compositionstart'));
    input.value = 'X';
    // happy-dom sends no compositionend on blur: the text the input shows is the change.
    range(e, 1, 6, 11);
    expect(input.classList.contains(COMPOSING_CLASS)).toBe(false);
    expect(texts(e)).toEqual(['alpha beta', 'gamma delta']);
    expect(rejection(e)).toBe('errors.compositionChanged');
    e.destroy();
  });
});

describe('a host store edit during a composition at a caret', () => {
  test('after the place ends the composition, and its text goes in at its place', async () => {
    const { e, input, fire } = await open(para(run('Hello')));
    caret(e, 0, 3);
    fire(composition('compositionstart'));
    input.value = 'HelX';
    input.setSelectionRange(4, 4);
    hostSetsText(e, 'Hel Hello');
    expect(input.classList.contains(COMPOSING_CLASS)).toBe(false);
    fire(composition('compositionend', 'X'));
    expect(texts(e)).toEqual(['HelX Hello']);
    e.destroy();
  });

  test('before the place, with other text there, refuses its text', async () => {
    const { e, input, fire } = await open(para(run('Hello')));
    caret(e, 0, 5);
    fire(composition('compositionstart'));
    input.value = 'HelloX';
    hostSetsText(e, 'Hi there');
    fire(composition('compositionend', 'X'));
    expect(texts(e)).toEqual(['Hi there']);
    expect(rejection(e)).toBe('errors.compositionChanged');
    e.destroy();
  });

  test('that puts the same text before the place cannot be seen (a known limit)', async () => {
    // With no collaboration session, no caret moves with its characters: the text before the
    // place reads the same, so the check passes and the text goes in after "Hello".
    const { e, input, fire } = await open(para(run('Hello')));
    caret(e, 0, 5);
    fire(composition('compositionstart'));
    input.value = 'HelloX';
    input.setSelectionRange(6, 6);
    hostSetsText(e, 'Hello Hello');
    fire(composition('compositionend', 'X'));
    expect(texts(e)).toEqual(['HelloX Hello']);
    e.destroy();
  });

  test('typing after a host edit goes in at the caret as usual', async () => {
    const { e, beforeInput } = await open(para(run('Hello')));
    caret(e, 0, 5);
    hostSetsText(e, 'Hi');
    beforeInput('insertText', 'Z');
    expect(texts(e)).toEqual(['HiZ']);
    expect(focusOf(e)).toEqual({ paragraph: 0, offset: 3 });
    e.destroy();
  });
});
