// The place of an IME composition on the slide surface, with no collaboration session: a
// composition over a selected range writes only over that range, and a host's store edit
// during a composition ends it. Text goes in only at a place that the check proves.

import { afterAll, describe, expect, test } from 'bun:test';
import { para, run } from '../../../../../scripts/fixtures/builder.ts';
import { xmlPartOf, type Element } from '../../store/index.ts';
import type { PptxEditorInstance } from '../index.ts';
import { PresenceLink } from '../collaboration-presence.ts';
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
/**
 * A store edit from outside the editor: text node `index` of the slide gets `text`. Origin
 * `user` is a host's own transaction; `remote` is a peer's, as a collaboration session sends.
 */
function hostSetsText(
  e: PptxEditorInstance,
  text: string,
  index = 0,
  origin: 'user' | 'remote' = 'user'
): void {
  const part = e.session!.info.slides[0]!.partName;
  type Node = Element | { type: 'text' };
  const all = (n: Node): Element[] =>
    n.type !== 'element' ? [] : n.local === 't' ? [n] : n.children.flatMap((c) => all(c as Node));
  const t = all(xmlPartOf(e.session!.store.package, part)!.root as Node)[index]!;
  const ops = [{ op: 'setText' as const, part, node: t.id, text }];
  expect(e.session!.store.transact(ops, { origin }).ok).toBe(true);
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
    expect(rejection(e)).toBe('errors.compositionRangeChanged');
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
    expect(rejection(e)).toBe('errors.compositionRangeChanged');
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
    expect(rejection(e)).toBe('errors.compositionRangeChanged');
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
    expect(rejection(e)).toBe('errors.compositionRangeChanged');
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

describe('the origin of a change during a composition', () => {
  /** A composition of "日本" at the end of the second of two paragraphs. */
  async function composing() {
    const opened = await open(para(run('one')) + para(run('two')));
    caret(opened.e, 1, 3);
    opened.fire(composition('compositionstart'));
    opened.input.value = 'two日本';
    const busy = () => opened.input.classList.contains(COMPOSING_CLASS);
    return { ...opened, busy };
  }

  test('a peer transaction in another paragraph keeps the composition', async () => {
    const { e, input, fire, busy } = await composing();
    hostSetsText(e, 'ONE', 0, 'remote');
    expect(busy()).toBe(true);
    input.setSelectionRange(5, 5);
    fire(composition('compositionend', '日本'));
    expect(texts(e)).toEqual(['ONE', 'two日本']);
    expect(rejection(e)).toBeNull();
    e.destroy();
  });

  test('a local transaction that a listener makes during a peer change reads as local', async () => {
    const { e, busy } = await composing();
    // A host answers the peer's change, while the store delivers it, with its own edit.
    let answered = false;
    const off = e.subscribe(() => {
      if (answered || e.session!.store.revision === 0) return;
      answered = true;
      hostSetsText(e, 'ONE!', 0);
      // The editor publishes again before the store delivers the host's edit: the surface
      // reads the store while the peer's change is still the one delivered.
      e.setFontFamilies(['Open Sans']);
    });
    hostSetsText(e, 'ONE', 0, 'remote');
    off();
    expect(answered).toBe(true);
    // The host's edit is local: the composition ends, and its text goes in at its place.
    expect(busy()).toBe(false);
    expect(texts(e)).toEqual(['ONE!', 'two日本']);
    expect(rejection(e)).toBeNull();
    e.destroy();
  });

  test('the record of a peer change holds for its revision, until a new deck', () => {
    const link = new PresenceLink({} as never);
    link.delivered({ origin: 'remote', toRevision: 3 });
    expect(link.remoteAt(3)).toBe(true);
    expect(link.remoteAt(4)).toBe(false);
    link.delivered({ origin: 'user', toRevision: 4 });
    expect(link.remoteAt(4)).toBe(false);
    link.delivered({ origin: 'remote', toRevision: 5 });
    // A new deck starts its revisions again: the old record must not match them.
    link.reset();
    expect(link.remoteAt(5)).toBe(false);
  });
});
