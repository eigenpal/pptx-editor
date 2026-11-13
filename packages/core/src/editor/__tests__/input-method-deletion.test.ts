// Deletions that a keyboard names as a range, not as a key: a keyboard that deletes through the
// platform text service (Android `deleteSurroundingText`) sends one `deleteContentBackward`
// for several characters or a word. The range deletes when it ends at the caret, checked
// against the model; any other range deletes one character, as the key does.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { attr, elements, find, textOf, xmlPartOf, type Element } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

const IN = 914_400;
const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const deck = (text: string): Uint8Array =>
  buildPptx({
    slides: [
      {
        shapes: shape({
          id: 2,
          name: 'T',
          x: IN,
          y: IN,
          cx: 6 * IN,
          cy: IN,
          text: para(run(text)),
        }),
      },
    ],
  });

async function attached(text = 'Hello world') {
  const g = globalThis as { ResizeObserver?: unknown };
  g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(deck(text))).ok).toBe(true);
  const host = doc.createElement('div') as unknown as HTMLElement;
  doc.body.append(host as unknown as Node);
  e.attach(host);
  const input = host.querySelector('textarea') as HTMLTextAreaElement;
  const id = elements(find(root(e), 'p:spTree')!).find(
    (c) => attr(find(c, 'p:cNvPr'), 'name') === 'T'
  )!.id;
  const at = { paragraph: 0, offset: text.length };
  e.setSelection({ kind: 'text', nodeId: id, anchor: at, focus: at });
  input.focus();
  return { e, input };
}

const root = (e: PptxEditorInstance): Element =>
  xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!.root;
const stored = (e: PptxEditorInstance): string => textOf(find(root(e), 'p:txBody'));

/**
 * A cancelable `deleteContentBackward` whose target range covers `[start, end)` of the
 * input's text, as Chromium reports a range in a text field: in a text node that holds the
 * field's value.
 */
function deletion(input: HTMLTextAreaElement, start: number, end: number): Event {
  const event = new win.InputEvent('beforeinput', {
    inputType: 'deleteContentBackward',
    bubbles: true,
    cancelable: true,
  }) as unknown as InputEvent;
  const node = doc.createTextNode(input.value);
  const range = { startContainer: node, endContainer: node, startOffset: start, endOffset: end };
  Object.defineProperty(event, 'getTargetRanges', { value: () => [range] });
  return event;
}

describe('input method deletions', () => {
  test('a range of a word at the caret deletes the word, as one undo step', async () => {
    const { e, input } = await attached();
    input.dispatchEvent(deletion(input, 6, 11));
    expect(stored(e)).toBe('Hello ');
    expect(e.snapshot().selection).toMatchObject({ focus: { offset: 6 } });
    // The input holds the new context for the keyboard.
    expect(input.value).toBe('Hello ');
    e.exec({ type: 'undo' });
    expect(stored(e)).toBe('Hello world');
    e.destroy();
  });

  test('several presses in one event delete each character they name', async () => {
    const { e, input } = await attached();
    input.dispatchEvent(deletion(input, 8, 11));
    expect(stored(e)).toBe('Hello wo');
    e.destroy();
  });

  test('a range of one character deletes as the key does', async () => {
    const { e, input } = await attached();
    input.dispatchEvent(deletion(input, 10, 11));
    expect(stored(e)).toBe('Hello worl');
    e.destroy();
  });

  test('one emoji named as two code units deletes once', async () => {
    const { e, input } = await attached('ab😀');
    input.dispatchEvent(deletion(input, 2, 4));
    expect(stored(e)).toBe('ab');
    e.destroy();
  });

  test('a range away from the caret deletes one character at the caret', async () => {
    const { e, input } = await attached();
    // The keyboard's copy of the text drifted: its range names the first word.
    input.dispatchEvent(deletion(input, 0, 5));
    expect(stored(e)).toBe('Hello worl');
    e.destroy();
  });

  test('a range over text the input does not hold deletes one character', async () => {
    const { e, input } = await attached();
    const event = deletion(input, 6, 11);
    input.value = 'Hello there';
    input.dispatchEvent(event);
    expect(stored(e)).toBe('Hello worl');
    e.destroy();
  });

  test('text the keyboard selects before it deletes is deleted', async () => {
    const { e, input } = await attached();
    // The keyboard selects the word and deletes it in the same task. A browser sends
    // `selectionchange` a task later; happy-dom sends it at once, to a focused field only.
    input.blur();
    input.setSelectionRange(6, 11);
    const event = new win.InputEvent('beforeinput', {
      inputType: 'deleteContentBackward',
      bubbles: true,
      cancelable: true,
    }) as unknown as Event;
    input.dispatchEvent(event);
    expect(stored(e)).toBe('Hello ');
    e.destroy();
  });
});
