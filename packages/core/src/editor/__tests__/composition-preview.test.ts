// The slide shows the text an IME composes, in place, before the composition writes it: the
// store, the undo history, and the file hold nothing of it until the composition ends.

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
import { attr, elements, find, textOf, xmlPartOf, type Element } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import {
  COMPOSING_CLASS,
  PREVIEWED_CLASS,
  bindSurfaceInput,
  type SurfaceInputHost,
} from '../surface-input.ts';

const IN = 914_400;
const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const fire = (target: EventTarget, event: Event): void => void target.dispatchEvent(event);
const composition = (type: string, data = ''): Event => {
  // happy-dom 21 has no `data` in its init dictionary: set it on the event.
  const event = new win.CompositionEvent(type, { bubbles: true }) as unknown as Event;
  Object.defineProperty(event, 'data', { value: data });
  return event;
};
/** The browser shows `value` in the field as the input method composes. */
function update(input: HTMLTextAreaElement, value: string): void {
  input.value = value;
  input.setSelectionRange(value.length, value.length);
  fire(input, new win.Event('input', { bubbles: true }) as unknown as Event);
}

const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        shapes:
          shape({
            id: 2,
            name: 'T',
            x: IN,
            y: IN,
            cx: 6 * IN,
            cy: IN,
            text: para(run('AB', 'altLang="ja-JP" sz="2400"')),
          }) +
          table({
            id: 3,
            x: IN,
            y: 4 * IN,
            cx: 2 * IN,
            cy: 370_840,
            columns: [2 * IN],
            tblPr: '<a:tblPr/>',
            rows: `<a:tr h="370840">${cell('xy')}</a:tr>`,
          }),
      },
    ],
  });

async function attached() {
  const g = globalThis as { ResizeObserver?: unknown };
  g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(deck())).ok).toBe(true);
  const host = doc.createElement('div') as unknown as HTMLElement;
  doc.body.append(host as unknown as Node);
  e.attach(host);
  const input = host.querySelector('textarea') as HTMLTextAreaElement;
  return { e, host, input };
}

const root = (e: PptxEditorInstance): Element =>
  xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!.root;
const shapeId = (e: PptxEditorInstance): string =>
  elements(find(root(e), 'p:spTree')!).find((c) => attr(find(c, 'p:cNvPr'), 'name') === 'T')!.id;
const stored = (e: PptxEditorInstance): string => textOf(find(root(e), 'p:txBody'));
const cellText = (e: PptxEditorInstance): string => textOf(find(root(e), 'a:tc'));
const tableId = (e: PptxEditorInstance): string =>
  elements(find(root(e), 'p:spTree')!).find((c) => c.local === 'graphicFrame')!.id;
/** The text the slide paints. */
const painted = (host: HTMLElement): string =>
  host.querySelector('.pptx-surface__slide')?.textContent?.replace(/\s+/g, '') ?? '';
const underline = (host: HTMLElement) => host.querySelectorAll('.pptx-surface__composition line');

function caret(e: PptxEditorInstance, from: number, to = from): void {
  e.setSelection({
    kind: 'text',
    nodeId: shapeId(e),
    anchor: { paragraph: 0, offset: from },
    focus: { paragraph: 0, offset: to },
  });
}

describe('the composition preview', () => {
  test('the slide shows the composing text in place; the model takes it at the end', async () => {
    const { e, host, input } = await attached();
    caret(e, 1);
    const revision = e.session!.store.revision;
    fire(input, composition('compositionstart'));
    update(input, 'Aにほん');
    expect(painted(host)).toContain('AにほんB');
    expect(underline(host).length).toBe(1);
    expect(input.classList.contains(COMPOSING_CLASS)).toBe(true);
    expect(input.classList.contains(PREVIEWED_CLASS)).toBe(true);
    // Nothing reaches the store, so nothing reaches the undo history or a peer.
    expect(stored(e)).toBe('AB');
    expect(e.session!.store.revision).toBe(revision);
    expect(e.snapshot().selection).toMatchObject({ focus: { offset: 1 } });
    update(input, 'A日本');
    expect(painted(host)).toContain('A日本B');
    fire(input, composition('compositionend', '日本'));
    expect(stored(e)).toBe('A日本B');
    expect(painted(host)).toContain('A日本B');
    expect(underline(host).length).toBe(0);
    expect(input.classList.contains(PREVIEWED_CLASS)).toBe(false);
    e.exec({ type: 'undo' });
    expect(stored(e)).toBe('AB');
    e.destroy();
  });

  test('a composition over a selected range shows in place of the range', async () => {
    const { e, host, input } = await attached();
    caret(e, 0, 2);
    fire(input, composition('compositionstart'));
    update(input, '한');
    expect(painted(host)).toContain('한');
    expect(painted(host)).not.toContain('AB');
    expect(stored(e)).toBe('AB');
    fire(input, composition('compositionend', '한'));
    expect(stored(e)).toBe('한');
    e.destroy();
  });

  test('a recomposed word that the keyboard erases leaves the slide and the model', async () => {
    const { e, host, input } = await attached();
    caret(e, 2);
    // The keyboard composes the word before the caret again, then erases it.
    input.setSelectionRange(0, 2);
    fire(input, composition('compositionstart'));
    update(input, 'A');
    expect(painted(host)).not.toContain('B');
    update(input, '');
    expect(painted(host)).not.toContain('A');
    expect(stored(e)).toBe('AB');
    fire(input, composition('compositionend', ''));
    expect(stored(e)).toBe('');
    e.destroy();
  });

  test('a press during a composition puts the composed text in first', async () => {
    const { e, input, host } = await attached();
    caret(e, 2);
    fire(input, composition('compositionstart'));
    update(input, 'AB日本');
    const stage = host.querySelector('.pptx-surface__stage')!;
    const press = new win.PointerEvent('pointerdown', { bubbles: true, button: 0 });
    fire(stage, press as unknown as Event);
    expect(stored(e)).toBe('AB日本');
    e.destroy();
  });

  test('the input shows the composed text when the slide cannot', () => {
    const root = doc.createElement('div') as unknown as HTMLElement;
    const input = doc.createElement('textarea') as unknown as HTMLTextAreaElement;
    root.append(input);
    doc.body.append(root as unknown as Node);
    const place = { at: { kind: 'none' as const }, ids: null, text: '', revision: 0 };
    const host: SurfaceInputHost = {
      textInput: () => {},
      selectionPlace: () => place,
      composeSelection: () => true,
      shortcut: () => false,
      copy: () => false,
      paste: () => false,
      run: () => {},
      beginHistoryGroup: () => {
        throw new Error('no gesture here');
      },
      refuse: () => {},
    };
    const shown: unknown[] = [];
    bindSurfaceInput(input, root, host, (preview) => (shown.push(preview), false));
    fire(input, composition('compositionstart'));
    update(input, 'にほん');
    expect(input.classList.contains(COMPOSING_CLASS)).toBe(true);
    expect(input.classList.contains(PREVIEWED_CLASS)).toBe(false);
    expect(input.style.width).toMatch(/px$/);
    fire(input, composition('compositionend', '日本'));
    expect(shown.at(-1)).toBeNull();
    expect(input.style.width).toBe('');
  });

  test('undo, redo, and the snapshot see nothing of a composition until it ends', async () => {
    const { e, input } = await attached();
    caret(e, 1);
    const before = e.snapshot();
    fire(input, composition('compositionstart'));
    update(input, 'Aにほん');
    const during = e.snapshot();
    expect(during.revision).toBe(before.revision);
    expect(during.canUndo).toBe(before.canUndo);
    expect(during.selection).toEqual(before.selection);
    fire(input, composition('compositionend', 'にほん'));
    expect(e.snapshot().canUndo).toBe(true);
    e.destroy();
  });

  test('a cancelled composition leaves the slide as it was', async () => {
    const { e, host, input } = await attached();
    caret(e, 1);
    fire(input, composition('compositionstart'));
    update(input, 'Aにほん');
    expect(painted(host)).toContain('AにほんB');
    // Escape: the browser takes the composed text back out of the field.
    update(input, 'A');
    fire(input, composition('compositionend', ''));
    expect(stored(e)).toBe('AB');
    expect(painted(host)).toContain('AB');
    expect(painted(host)).not.toContain('に');
    expect(underline(host).length).toBe(0);
    e.destroy();
  });

  test('a composition that a peer edit takes away leaves the slide', async () => {
    const { e, host, input } = await attached();
    caret(e, 1);
    fire(input, composition('compositionstart'));
    update(input, 'Aにほん');
    // A peer replaces the text before the place: the composition is lost, its text refused.
    const part = e.session!.info.slides[0]!.partName;
    const t = find(root(e), 'a:t')!;
    const ops = [{ op: 'setText' as const, part, node: t.id, text: 'ZZ' }];
    expect(e.session!.store.transact(ops, { origin: 'remote' }).ok).toBe(true);
    fire(input, composition('compositionend', 'にほん'));
    expect(stored(e)).toBe('ZZ');
    expect(painted(host)).not.toContain('に');
    e.destroy();
  });

  test('a composition in a table cell shows in the cell', async () => {
    const { e, host, input } = await attached();
    const at = { paragraph: 0, offset: 1 };
    e.setSelection({
      kind: 'text',
      nodeId: tableId(e),
      cell: { row: 0, column: 0 },
      anchor: at,
      focus: at,
    });
    fire(input, composition('compositionstart'));
    update(input, 'x中');
    expect(painted(host)).toContain('x中y');
    expect(underline(host).length).toBe(1);
    expect(cellText(e)).toBe('xy');
    fire(input, composition('compositionend', '中'));
    expect(cellText(e)).toBe('x中y');
    e.destroy();
  });

  test('a word composed again is underlined whole', async () => {
    const { e, host, input } = await attached();
    caret(e, 2);
    input.setSelectionRange(0, 2);
    fire(input, composition('compositionstart'));
    update(input, 'AC');
    const line = underline(host)[0]!;
    // The underline starts at the word start, not at the one letter that changed.
    const x1 = Number(line.getAttribute('x1'));
    const x2 = Number(line.getAttribute('x2'));
    update(input, 'AB');
    expect(underline(host).length).toBe(1);
    expect(Number(underline(host)[0]!.getAttribute('x1'))).toBe(x1);
    expect(x2).toBeGreaterThan(x1);
    fire(input, composition('compositionend', 'AC'));
    expect(stored(e)).toBe('AC');
    e.destroy();
  });

  test('a right-click during a composition puts the composed text in first', async () => {
    const { e, input, host } = await attached();
    caret(e, 2);
    fire(input, composition('compositionstart'));
    update(input, 'AB日本');
    const stage = host.querySelector('.pptx-surface__stage')!;
    fire(
      stage,
      new win.PointerEvent('pointerdown', { bubbles: true, button: 2 }) as unknown as Event
    );
    expect(stored(e)).toBe('AB日本');
    e.destroy();
  });
});
