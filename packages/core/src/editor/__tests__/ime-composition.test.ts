// IME composition on the slide surface, as Chinese, Japanese, and Korean input methods send it:
// the composed text goes in once at composition end, in one undo step; the composing text
// shows at the caret; the key that commits a composition runs no shortcut; a late commit event
// inserts nothing; Hangul syllables composed one after another each go in once; Escape cancels
// a composition and keeps the text selection.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { attr, elements, find, textOf, xmlPartOf, type Element } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import {
  COMPOSING_CLASS,
  bindSurfaceInput,
  isCompositionKey,
  type SurfaceInputHost,
} from '../surface-input.ts';

const IN = 914_400;
const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

type Call = readonly unknown[];

/** A textarea and root bound to a recording host. */
function bound(): { input: HTMLTextAreaElement; root: HTMLElement; calls: Call[] } {
  const root = doc.createElement('div') as unknown as HTMLElement;
  const input = doc.createElement('textarea') as unknown as HTMLTextAreaElement;
  root.append(input);
  doc.body.append(root as unknown as Node);
  const calls: Call[] = [];
  const host: SurfaceInputHost = {
    textInput: (kind, data) => calls.push(['text', kind, data]),
    shortcut: (key) => (calls.push(['key', key]), true),
    copy: () => false,
    paste: () => false,
    run: (command) => calls.push(['run', command.type]),
    beginHistoryGroup: () => {
      throw new Error('no gesture here');
    },
    refuse: (reason, detail) => calls.push(['refuse', reason, detail]),
  };
  bindSurfaceInput(input, root, host);
  return { input, root, calls };
}

const fire = (target: EventTarget, event: Event): boolean => {
  target.dispatchEvent(event);
  return event.defaultPrevented;
};
const composition = (type: string, data = ''): Event => {
  // happy-dom 21 has no `data` in its init dictionary: set it on the event.
  const event = new win.CompositionEvent(type, { bubbles: true }) as unknown as Event;
  Object.defineProperty(event, 'data', { value: data });
  return event;
};
const beforeInput = (inputType: string, data?: string) =>
  new win.InputEvent('beforeinput', {
    inputType,
    data,
    bubbles: true,
    cancelable: true,
  }) as unknown as Event;
const keyDown = (key: string, init: { keyCode?: number; isComposing?: boolean } = {}) =>
  new win.KeyboardEvent('keydown', {
    key,
    bubbles: true,
    cancelable: true,
    ...init,
  }) as unknown as Event;

/** One composition as an input method sends it: updates, then the committed text. */
function compose(input: HTMLTextAreaElement, steps: readonly string[], commit: string): void {
  fire(input, composition('compositionstart'));
  for (const step of steps) {
    // The browser shows the composing text in the field; `beforeinput` cannot be prevented.
    fire(input, beforeInput('insertCompositionText', step));
    input.value = step;
    fire(input, new win.Event('input', { bubbles: true }) as unknown as Event);
    fire(input, composition('compositionupdate', step));
  }
  fire(input, composition('compositionend', commit));
}

describe('the surface input during an IME composition', () => {
  test('a Japanese composition goes in once, at composition end', () => {
    const { input, calls } = bound();
    compose(input, ['n', 'に', 'にほ', 'にほんご', '日本語'], '日本語');
    expect(calls).toEqual([['text', 'text', '日本語']]);
    expect(input.value).toBe('');
  });

  test('the composing text shows at the caret, and the field hides again after', () => {
    const { input } = bound();
    fire(input, composition('compositionstart'));
    expect(input.classList.contains(COMPOSING_CLASS)).toBe(true);
    input.value = 'にほん';
    fire(input, new win.Event('input', { bubbles: true }) as unknown as Event);
    expect(input.style.width).toMatch(/px$/);
    fire(input, composition('compositionend', '日本'));
    expect(input.classList.contains(COMPOSING_CLASS)).toBe(false);
    expect(input.style.width).toBe('');
  });

  test('Hangul syllables composed one after another each go in once', () => {
    const { input, calls } = bound();
    compose(input, ['ㅎ', '하', '한'], '한');
    compose(input, ['ㄱ', '그', '글'], '글');
    // A space ends the last syllable and goes in as typed text.
    fire(input, beforeInput('insertText', ' '));
    expect(calls).toEqual([
      ['text', 'text', '한'],
      ['text', 'text', '글'],
      ['text', 'text', ' '],
    ]);
  });

  test('keys during a composition, and the key that commits it, run no shortcut', () => {
    const { input, root, calls } = bound();
    fire(input, composition('compositionstart'));
    expect(fire(input, keyDown('ArrowLeft', { isComposing: true }))).toBe(false);
    expect(fire(input, keyDown('Escape', { isComposing: true }))).toBe(false);
    fire(input, composition('compositionend', '漢字'));
    // Some browsers send the committing Enter after `compositionend`, with key code 229.
    expect(fire(input, keyDown('Enter', { keyCode: 229 }))).toBe(false);
    expect(fire(root, keyDown('Enter', { keyCode: 229 }))).toBe(false);
    expect(calls).toEqual([['text', 'text', '漢字']]);
    // A plain Enter after the composition is a shortcut again.
    expect(fire(input, keyDown('Enter'))).toBe(true);
    expect(calls[calls.length - 1]).toEqual(['key', 'Enter']);
  });

  test('a commit event after composition end inserts nothing and refuses nothing', () => {
    const { input, calls } = bound();
    compose(input, ['中'], '中');
    expect(fire(input, beforeInput('insertFromComposition', '中'))).toBe(true);
    expect(fire(input, beforeInput('deleteCompositionText'))).toBe(true);
    expect(calls).toEqual([['text', 'text', '中']]);
  });

  test('a cancelled composition inserts nothing', () => {
    const { input, calls } = bound();
    compose(input, ['ni'], '');
    expect(calls).toEqual([]);
    expect(input.classList.contains(COMPOSING_CLASS)).toBe(false);
  });

  test('isCompositionKey: composing keys and key code 229 only', () => {
    expect(isCompositionKey({ isComposing: true, keyCode: 13 } as KeyboardEvent)).toBe(true);
    expect(isCompositionKey({ isComposing: false, keyCode: 229 } as KeyboardEvent)).toBe(true);
    expect(isCompositionKey({ isComposing: false, keyCode: 13 } as KeyboardEvent)).toBe(false);
  });
});

describe('IME composition in the editor', () => {
  const deck = (): Uint8Array =>
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
            text: para(run('AB', 'altLang="ja-JP"')),
          }),
        },
      ],
    });

  async function attached(): Promise<{ e: PptxEditorInstance; input: HTMLTextAreaElement }> {
    const g = globalThis as { ResizeObserver?: unknown };
    g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await e.load(deck())).ok).toBe(true);
    const host = doc.createElement('div') as unknown as HTMLElement;
    doc.body.append(host as unknown as Node);
    e.attach(host);
    return { e, input: host.querySelector('textarea') as HTMLTextAreaElement };
  }

  const root = (e: PptxEditorInstance): Element =>
    xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!.root;
  const id = (e: PptxEditorInstance): string =>
    elements(find(root(e), 'p:spTree')!).find((c) => attr(find(c, 'p:cNvPr'), 'name') === 'T')!.id;
  const text = (e: PptxEditorInstance): string => textOf(find(root(e), 'p:txBody'));

  test('the composed text goes in at the caret once, as one undo step', async () => {
    const { e, input } = await attached();
    const nodeId = id(e);
    e.setSelection({
      kind: 'text',
      nodeId,
      anchor: { paragraph: 0, offset: 1 },
      focus: { paragraph: 0, offset: 1 },
    });
    compose(input, ['に', 'にほん', '日本'], '日本');
    expect(text(e)).toBe('A日本B');
    expect(e.snapshot().selection).toMatchObject({ focus: { paragraph: 0, offset: 3 } });
    e.exec({ type: 'undo' });
    expect(text(e)).toBe('AB');
    e.destroy();
  });

  test('a composition replaces the selected text', async () => {
    const { e, input } = await attached();
    e.setSelection({
      kind: 'text',
      nodeId: id(e),
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 2 },
    });
    compose(input, ['한', '한국'], '한국');
    expect(text(e)).toBe('한국');
    e.destroy();
  });
});
