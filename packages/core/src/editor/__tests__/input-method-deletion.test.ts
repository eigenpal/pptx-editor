// Deletions that a keyboard names as a range, not as a key: a keyboard that deletes through the
// platform text service (Android `deleteSurroundingText`) makes the browser select the
// characters in the text field, then send one `deleteContentBackward`, in one task. A text
// field reports no target range. The selection deletes, checked against the model; one
// grapheme at the caret deletes as the key does.

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
 * The browser selects `[start, end)` of the field for a keyboard's deletion and sends a
 * cancelable `deleteContentBackward`, in one task. A browser sends `selectionchange` a task
 * later; happy-dom sends it at once to a focused field, so the field is not focused here.
 */
function deletion(input: HTMLTextAreaElement, start: number, end: number): Event {
  input.blur();
  input.setSelectionRange(start, end);
  return new win.InputEvent('beforeinput', {
    inputType: 'deleteContentBackward',
    bubbles: true,
    cancelable: true,
  }) as unknown as Event;
}

describe('input method deletions', () => {
  test('a word the keyboard selects at the caret deletes, as one undo step', async () => {
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

  test('one selected character at the caret deletes as the key does', async () => {
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

  test('a word the keyboard selects away from the caret deletes, and the caret stays after', async () => {
    const { e, input } = await attached();
    input.dispatchEvent(deletion(input, 0, 6));
    expect(stored(e)).toBe('world');
    expect(e.snapshot().selection).toMatchObject({ focus: { offset: 5 } });
    e.destroy();
  });

  test('a selection over text the input does not hold deletes one character', async () => {
    const { e, input } = await attached();
    input.value = 'Hello there';
    input.dispatchEvent(deletion(input, 6, 11));
    expect(stored(e)).toBe('Hello worl');
    e.destroy();
  });

  test('a hardware Backspace deletes as the key, whatever a native key left selected', async () => {
    const { e, input } = await attached();
    input.focus();
    input.setSelectionRange(6, 11);
    input.dispatchEvent(
      new win.KeyboardEvent('keydown', { key: 'Backspace', bubbles: true }) as unknown as Event
    );
    input.dispatchEvent(
      new win.InputEvent('beforeinput', {
        inputType: 'deleteContentBackward',
        bubbles: true,
        cancelable: true,
      }) as unknown as Event
    );
    expect(stored(e)).toBe('Hello worl');
    e.destroy();
  });
});
