// Caret stops, word steps, and the double-click word of a slide paragraph. The clusters and the
// words come from the UAX #29 segmentation shared with DOCX Editor (`@eigenpal/ooxml-core`).

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { attr, elements, find, textOf, xmlPartOf, type Element } from '../../store/index.ts';
import type { ResolvedParagraph } from '../../cascade/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { caretAt } from '../surface-geometry.ts';
import { moveCaret } from '../text-input.ts';
import { paragraphUnits, wordSelectionAt } from '../../automation/commands/text-units.ts';

const IN = 914_400;

/** A paragraph of text runs, as the cascade resolves it, for the unit functions. */
const paragraph = (...texts: string[]): ResolvedParagraph =>
  ({ runs: texts.map((text) => ({ kind: 'text', text })) }) as unknown as ResolvedParagraph;

async function open(text: string): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    slides: [{ shapes: shape({ id: 2, name: 'T', x: IN, y: IN, cx: 6 * IN, cy: IN, text }) }],
  });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const rootOf = (e: PptxEditorInstance): Element =>
  xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!.root;
const textShapeId = (e: PptxEditorInstance): string =>
  elements(find(rootOf(e), 'p:spTree')!).find((c) => attr(find(c, 'p:cNvPr'), 'name') === 'T')!.id;
const itemOf = (e: PptxEditorInstance) => {
  const id = textShapeId(e);
  const item = e.session!.layout(0)!.items.find((i) => i.element.nodeId === id)!;
  if (item.kind !== 'shape') throw new Error('no shape');
  return item;
};

describe('paragraph units', () => {
  test('an emoji ZWJ sequence and a field are one stop each; a break ends a word', () => {
    const family = '\u{1F469}\u200D\u{1F469}\u200D\u{1F467}';
    const p = {
      runs: [
        { kind: 'text', text: `a${family}b` },
        { kind: 'field', text: '12' },
        { kind: 'text', text: 'cd' },
        { kind: 'break' },
        { kind: 'text', text: 'ef' },
      ],
    } as unknown as ResolvedParagraph;
    const u = paragraphUnits(p);
    const f = family.length;
    expect(u.stops).toEqual([0, 1, 1 + f, 2 + f, 4 + f, 5 + f, 6 + f, 7 + f, 8 + f, 9 + f]);
    expect(wordSelectionAt(u, 2 + f + 1)).toEqual([2 + f, 4 + f]);
    expect(wordSelectionAt(u, 4 + f + 1)).toEqual([4 + f, 6 + f]);
  });

  test('a word that crosses a run boundary is one word', () => {
    const u = paragraphUnits(paragraph('Bo', 'ld text'));
    expect(u.words.filter((w) => w.word).map((w) => [w.start, w.end])).toEqual([
      [0, 4],
      [5, 9],
    ]);
  });

  test('a combining mark in its own run stays in its cluster', () => {
    const u = paragraphUnits(paragraph('e', '́t'));
    expect(u.stops).toEqual([0, 2, 3]);
  });

  test('Thai text splits into dictionary words', () => {
    // "สวัสดีครับ": "สวัสดี" (6 units) and "ครับ" (4 units).
    const u = paragraphUnits(paragraph('สวัสดีครับ'));
    expect(u.words.map((w) => [w.start, w.end])).toEqual([
      [0, 6],
      [6, 10],
    ]);
  });

  test('a decimal number is one word (UAX #29 WB11, WB12)', () => {
    const u = paragraphUnits(paragraph('pi 3.14'));
    expect(u.words.at(-1)).toEqual({ start: 3, end: 7, word: true });
  });
});

describe('double-click word', () => {
  const u = paragraphUnits(paragraph('well-known words  here'));

  test('selects the word under the pointer; a hyphen ends a word (UAX #29)', () => {
    expect(wordSelectionAt(u, 2)).toEqual([0, 4]);
    expect(wordSelectionAt(u, 7)).toEqual([5, 10]);
  });

  test('at a word end, selects that word rather than the space after it', () => {
    expect(wordSelectionAt(u, 16)).toEqual([11, 16]);
  });

  test('between words, selects the run of spaces there', () => {
    expect(wordSelectionAt(u, 17)).toEqual([16, 18]);
  });
});

describe('word steps', () => {
  test('Alt or Ctrl with Right steps past a word and its spaces; Left to a word start', async () => {
    const e = await open(para(run('alpha beta')) + para(run('gamma')));
    const item = itemOf(e);
    const right = (paragraph: number, offset: number) =>
      moveCaret(item, { paragraph, offset }, 'ArrowRight', { altKey: true });
    const left = (paragraph: number, offset: number) =>
      moveCaret(item, { paragraph, offset }, 'ArrowLeft', { ctrlKey: true });
    // Right lands at the start of the next word, past the space.
    expect(right(0, 0)).toEqual({ paragraph: 0, offset: 6 });
    // From a word end, the step crosses only the spaces, to the next word.
    expect(right(0, 5)).toEqual({ paragraph: 0, offset: 6 });
    expect(right(0, 6)).toEqual({ paragraph: 0, offset: 10 });
    expect(right(0, 10)).toEqual({ paragraph: 1, offset: 0 });
    expect(left(0, 10)).toEqual({ paragraph: 0, offset: 6 });
    expect(left(0, 6)).toEqual({ paragraph: 0, offset: 0 });
    expect(left(1, 0)).toEqual({ paragraph: 0, offset: 10 });
    // A plain arrow still steps one character.
    expect(moveCaret(item, { paragraph: 0, offset: 0 }, 'ArrowRight')).toEqual({
      paragraph: 0,
      offset: 1,
    });
    // Cmd alone with Left and Right goes to the paragraph edges. With Alt or Ctrl too, the
    // word step wins, as in DOCX Editor's key map.
    const meta = { metaKey: true };
    expect(moveCaret(item, { paragraph: 0, offset: 3 }, 'ArrowRight', meta).offset).toBe(10);
    expect(moveCaret(item, { paragraph: 0, offset: 3 }, 'ArrowLeft', meta).offset).toBe(0);
    expect(
      moveCaret(item, { paragraph: 0, offset: 3 }, 'ArrowRight', { metaKey: true, altKey: true })
    ).toEqual({ paragraph: 0, offset: 6 });
  });

  test('Ctrl or Cmd with Home, End, Up, and Down goes to the text body edges', async () => {
    const e = await open(para(run('alpha beta')) + para(run('gamma')));
    const item = itemOf(e);
    const at = { paragraph: 0, offset: 3 };
    const end = { paragraph: 1, offset: 5 };
    const start = { paragraph: 0, offset: 0 };
    expect(moveCaret(item, at, 'End', { ctrlKey: true })).toEqual(end);
    expect(moveCaret(item, at, 'ArrowDown', { metaKey: true })).toEqual(end);
    expect(moveCaret(item, end, 'Home', { metaKey: true })).toEqual(start);
    expect(moveCaret(item, end, 'ArrowUp', { metaKey: true })).toEqual(start);
    // Without Ctrl or Cmd, Home and End stay in the paragraph.
    expect(moveCaret(item, at, 'End')).toEqual({ paragraph: 0, offset: 10 });
  });

  test('Alt or Ctrl with Up and Down goes to a paragraph start', async () => {
    const e = await open(para(run('alpha beta')) + para(run('gamma')) + para(run('delta')));
    const item = itemOf(e);
    const up = (paragraph: number, offset: number, mods: { altKey?: boolean; ctrlKey?: boolean }) =>
      moveCaret(item, { paragraph, offset }, 'ArrowUp', mods);
    const down = (paragraph: number, offset: number) =>
      moveCaret(item, { paragraph, offset }, 'ArrowDown', { altKey: true });
    // Inside a paragraph, Up goes to its start; at its start, to the start of the one before.
    expect(up(1, 3, { altKey: true })).toEqual({ paragraph: 1, offset: 0 });
    expect(up(1, 0, { ctrlKey: true })).toEqual({ paragraph: 0, offset: 0 });
    expect(up(0, 0, { altKey: true })).toEqual({ paragraph: 0, offset: 0 });
    // Down goes to the start of the next paragraph, and in the last one to its end.
    expect(down(0, 3)).toEqual({ paragraph: 1, offset: 0 });
    expect(down(2, 1)).toEqual({ paragraph: 2, offset: 5 });
  });

  test('Home and End go to the edges of the laid-out line, as Cmd with Left and Right', async () => {
    const text = 'word '.repeat(40).trim();
    const e = await open(para(run(text)));
    const item = itemOf(e);
    const lines = item.text!.lines;
    expect(lines.length).toBeGreaterThan(2);
    const [first, second] = [lines[0]!, lines[1]!];
    const last = lines[lines.length - 1]!;
    const inSecond = { paragraph: 0, offset: second.start + 2 };
    expect(moveCaret(item, inSecond, 'Home')).toEqual({ paragraph: 0, offset: second.start });
    expect(moveCaret(item, inSecond, 'ArrowLeft', { metaKey: true }).offset).toBe(second.start);
    // A wrapped line ends before the space it wraps at, so the caret stays on that line.
    expect(text[first.end - 1]).toBe(' ');
    const firstEnd = moveCaret(item, { paragraph: 0, offset: 1 }, 'End');
    expect(firstEnd).toEqual({ paragraph: 0, offset: first.end - 1 });
    expect(caretAt(item, firstEnd)!.y1).toBe(caretAt(item, { paragraph: 0, offset: 1 })!.y1);
    // The last line ends at the end of the paragraph.
    const inLast = { paragraph: 0, offset: last.start + 1 };
    expect(moveCaret(item, inLast, 'End')).toEqual({ paragraph: 0, offset: text.length });
    expect(moveCaret(item, inLast, 'ArrowRight', { metaKey: true }).offset).toBe(text.length);
  });

  test('End goes before a line break, and after the last character of a wrap in a word', async () => {
    const br = await open(para(run('abc') + '<a:br/>' + run('def')));
    const brItem = itemOf(br);
    expect(moveCaret(brItem, { paragraph: 0, offset: 1 }, 'End')).toEqual({
      paragraph: 0,
      offset: 3,
    });
    const long = await open(para(run('x'.repeat(400))));
    const item = itemOf(long);
    const first = item.text!.lines[0]!;
    expect(item.text!.lines.length).toBeGreaterThan(1);
    // A wrap inside a word: End goes after the line's last character, so typed text goes in
    // there (the caret draws on the next line: it has no side).
    expect(moveCaret(item, { paragraph: 0, offset: 1 }, 'End').offset).toBe(first.end);
  });

  test('a word step at the text body edges stays there; an empty paragraph is one stop', async () => {
    const e = await open(para(run('one')) + para('') + para(run('two')));
    const item = itemOf(e);
    const alt = { altKey: true };
    expect(moveCaret(item, { paragraph: 0, offset: 0 }, 'ArrowLeft', alt)).toEqual({
      paragraph: 0,
      offset: 0,
    });
    expect(moveCaret(item, { paragraph: 2, offset: 3 }, 'ArrowRight', alt)).toEqual({
      paragraph: 2,
      offset: 3,
    });
    expect(moveCaret(item, { paragraph: 0, offset: 3 }, 'ArrowRight', alt)).toEqual({
      paragraph: 1,
      offset: 0,
    });
    expect(moveCaret(item, { paragraph: 1, offset: 0 }, 'ArrowRight', alt)).toEqual({
      paragraph: 2,
      offset: 0,
    });
  });

  test('a word step and a word deletion reach the same offset', async () => {
    const e = await open(para(run('one two three')));
    const id = textShapeId(e);
    const to = moveCaret(itemOf(e), { paragraph: 0, offset: 13 }, 'ArrowLeft', {
      altKey: true,
    });
    e.setSelection({
      kind: 'text',
      nodeId: id,
      anchor: { paragraph: 0, offset: 13 },
      focus: { paragraph: 0, offset: 13 },
    });
    expect(e.input('deleteWordBackward').ok).toBe(true);
    expect(textOf(find(rootOf(e), 'p:txBody'))).toBe('one two '.slice(0, to.offset));
  });
});

// A DOM for this file only: the engine test run registers no global DOM.
const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

describe('the surface key and pointer path', () => {
  /** Attach the editor; happy-dom lays nothing out, so the stage is 1280 px wide. */
  async function attached(text: string) {
    const g = globalThis as { ResizeObserver?: unknown };
    g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
    const e = await open(text);
    const host = doc.createElement('div') as unknown as HTMLElement;
    doc.body.append(host as unknown as Node);
    e.attach(host);
    const stage = host.querySelector('.pptx-surface__stage') as HTMLElement;
    stage.getBoundingClientRect = () =>
      ({ left: 0, top: 0, right: 1280, bottom: 720, width: 1280, height: 720 }) as DOMRect;
    const input = host.querySelector('textarea') as HTMLTextAreaElement;
    type Mods = { altKey?: boolean; ctrlKey?: boolean; metaKey?: boolean; shiftKey?: boolean };
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
    return { e, stage, key };
  }
  const caret = (paragraph: number, offset: number) => ({ paragraph, offset });

  test('Alt, Ctrl, and Shift with the arrows step and extend by word, and the event is taken', async () => {
    const { e, key } = await attached(para(run('alpha beta')) + para(run('gamma')));
    const nodeId = textShapeId(e);
    e.setSelection({ kind: 'text', nodeId, anchor: caret(0, 0), focus: caret(0, 0) });
    expect(key('ArrowRight', { altKey: true })).toBe(true);
    expect(e.snapshot().selection).toMatchObject({ anchor: caret(0, 6), focus: caret(0, 6) });
    key('ArrowRight', { ctrlKey: true, shiftKey: true });
    expect(e.snapshot().selection).toMatchObject({ anchor: caret(0, 6), focus: caret(0, 10) });
    key('ArrowRight', { altKey: true, shiftKey: true });
    expect(e.snapshot().selection).toMatchObject({ anchor: caret(0, 6), focus: caret(1, 0) });
    // Alt+Left is the browser's Back key on some platforms: the editor takes the event. On the
    // range from 6, the word step starts at the range start.
    expect(key('ArrowLeft', { altKey: true })).toBe(true);
    expect(e.snapshot().selection).toMatchObject({ anchor: caret(0, 0), focus: caret(0, 0) });
    key('End', { ctrlKey: true });
    expect(e.snapshot().selection).toMatchObject({ focus: caret(1, 5) });
    key('ArrowUp', { metaKey: true, shiftKey: true });
    expect(e.snapshot().selection).toMatchObject({ anchor: caret(1, 5), focus: caret(0, 0) });
    e.destroy();
  });

  test('on a range, a plain arrow collapses it; a word key steps on from its edge', async () => {
    const { e, key } = await attached(para(run('one two three four')));
    const nodeId = textShapeId(e);
    const range = { kind: 'text' as const, nodeId, anchor: caret(0, 5), focus: caret(0, 7) };
    e.setSelection(range);
    key('ArrowRight');
    expect(e.snapshot().selection).toMatchObject({ anchor: caret(0, 7), focus: caret(0, 7) });
    e.setSelection(range);
    key('ArrowRight', { altKey: true });
    expect(e.snapshot().selection).toMatchObject({ anchor: caret(0, 8), focus: caret(0, 8) });
    e.setSelection(range);
    key('ArrowLeft', { ctrlKey: true });
    expect(e.snapshot().selection).toMatchObject({ anchor: caret(0, 4), focus: caret(0, 4) });
    e.destroy();
  });

  test('a double-click selects the Unicode word under the pointer', async () => {
    const { e, stage } = await attached(para(run('well-known words')));
    const item = itemOf(e);
    // A point inside "known" (offsets 5 to 10), between its second and third letters.
    const c = caretAt(item, caret(0, 7))!;
    const width = e.session!.layout(0)!.slide.width;
    const px = (emu: number) => (emu / width) * 1280;
    const at = { clientX: px(c.x1), clientY: px((c.y1 + c.y2) / 2), bubbles: true };
    stage.dispatchEvent(new win.MouseEvent('dblclick', at) as unknown as Event);
    expect(e.snapshot().selection).toMatchObject({
      kind: 'text',
      nodeId: textShapeId(e),
      anchor: caret(0, 5),
      focus: caret(0, 10),
    });
    e.destroy();
  });
});
