// The caret, the selection, the caret keys, and hit tests on lines in visual order
// (visual-caret.ts, visual-caret-keys.ts). Offsets stay logical; a caret place is an offset and
// the side of a direction boundary it shows on (`TextPoint.affinity`).

import { describe, expect, test } from 'bun:test';
import { GlobalWindow, Window } from 'happy-dom';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '../../store/index.ts';
import { resolveSlide } from '../../cascade/index.ts';
import {
  createFixedMeasurer,
  layoutSlide,
  runGlyphs,
  type LayoutItem,
  type Line,
} from '../../layout/index.ts';
import type { TextPoint } from '../../contracts/index.ts';
import { apply, invert } from '../../layout/index.ts';
import { caretAt, pointAt, textFrameMatrix } from '../surface-geometry.ts';
import { selectionRects } from '../surface-text-selection.ts';
import { moveCaret, moveTextSelection } from '../text-input.ts';
import { caretXOn, lineStopsOf, type VisualCaretStop } from '../visual-caret.ts';
import { createPptxEditor } from '../index.ts';

type ShapeItem = Extract<LayoutItem, { kind: 'shape' }>;

const shape = (body: string, cx = 6_000_000): string =>
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
  `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="4000000"/></a:xfrm>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
  `<p:txBody><a:bodyPr/><a:lstStyle/>${body}</p:txBody></p:sp>`;
const r = (t: string, lang = 'en-US'): string =>
  `<a:r><a:rPr lang="${lang}" sz="1800"/><a:t xml:space="preserve">${t}</a:t></a:r>`;
const rtl = (runs: string): string => `<a:p><a:pPr rtl="1" algn="r"/>${runs}</a:p>`;
const ltr = (runs: string): string => `<a:p>${runs}</a:p>`;

function lay(body: string, cx?: number): ShapeItem {
  const res = readPptx(buildPptx({ slides: [{ shapes: shape(body, cx), layout: 'blank' }] }));
  if (!res.ok) throw new Error('open failed');
  const item = layoutSlide(resolveSlide(res.pkg, 0)!, createFixedMeasurer()).items[0];
  if (item.kind !== 'shape' || !item.text) throw new Error('no text');
  return item;
}

// Hebrew around a Latin word; Arabic with a Latin word and digits; Hebrew with a number.
const HEBREW_LATIN = rtl(r('אבג ', 'he-IL') + r('PPTX') + r(' דהו.', 'he-IL'));
const ARABIC_LATIN_DIGITS = rtl(r('سلام ', 'ar-SA') + r('PPTX') + r(' دار 42.', 'ar-SA'));
const HEBREW_NUMBER = rtl(r('שלום 120.50 תאריך', 'he-IL'));
const LATIN_IN_RTL = rtl(r('abc PPTX def.'));
const CASES = [
  ['Hebrew with a Latin word', HEBREW_LATIN],
  ['Arabic with a Latin word and digits', ARABIC_LATIN_DIGITS],
  ['Hebrew with a number', HEBREW_NUMBER],
  ['Latin text in a right-to-left paragraph', LATIN_IN_RTL],
] as const;

/** Every visual edge of a line: the edges of its units. */
function edgesOf(line: Line): number[] {
  const xs = [...line.caretX, ...line.bidi!.trailingX];
  return [...new Set(xs)].sort((a, b) => a - b);
}

/** The caret's x in the text frame. */
function caretX(item: ShapeItem, point: TextPoint): number {
  const c = caretAt(item, point)!;
  return apply(invert(textFrameMatrix(item)!), c.x1, c.y1).x;
}

/** A slide point at text-frame x on the first line. */
function slidePoint(item: ShapeItem, x: number): { x: number; y: number } {
  const line = item.text!.lines[0];
  return apply(textFrameMatrix(item)!, x, line.top + line.height / 2);
}

const place = (s: VisualCaretStop) => (s.affinity ? `${s.offset}^` : `${s.offset}`);

describe('caret places on a line in visual order', () => {
  test('a left-to-right line has no visual order data', () => {
    const line = lay(ltr(r('plain text'))).text!.lines[0];
    expect(line.bidi).toBeUndefined();
  });

  for (const [name, body] of CASES)
    test(`${name}: each visual edge has one place, and each offset shows`, () => {
      const item = lay(body);
      const line = item.text!.lines[0];
      const stops = lineStopsOf(item, line);
      expect(stops.map((s) => s.x)).toEqual(edgesOf(line));
      const offsets = new Set(stops.map((s) => s.offset));
      for (let o = line.start; o <= line.end; o += 1) expect(offsets.has(o)).toBe(true);
      // The painted caret of each place is at the place.
      for (const s of stops) expect(caretX(item, { paragraph: 0, ...s })).toBe(s.x);
    });

  for (const [name, body] of CASES)
    test(`${name}: each unit's caret cell is the cell paint draws its glyph in`, () => {
      const line = lay(body).text!.lines[0];
      const bidi = line.bidi!;
      for (const run of line.runs) {
        if (run.kind !== 'text') continue;
        for (const glyph of runGlyphs(run)) {
          const i = run.start + glyph.index - line.start;
          expect(Math.min(line.caretX[i]!, bidi.trailingX[i]!)).toBe(glyph.x);
        }
      }
    });

  test('the offset after a Latin word in Hebrew shows after the word when upstream', () => {
    const item = lay(HEBREW_LATIN);
    const line = item.text!.lines[0];
    const word = line.runs.find((run) => run.text === 'PPTX')!;
    // Before P (offset 4) and after X (offset 8) share P's left edge downstream.
    expect(caretXOn(line, 4)).toBe(word.x[0]);
    expect(caretXOn(line, 8)).toBe(word.x[0]);
    // Upstream, offset 8 is the trailing edge of X: the right edge of the word.
    expect(caretXOn(line, 8, 'upstream')).toBe(word.x[4]);
    // Right to left, from offset 0 at the right edge.
    expect(lineStopsOf(item, line).map(place).reverse()).toEqual([
      '0',
      '1',
      '2',
      '3',
      '8^',
      '7',
      '6',
      '5',
      '4',
      '9',
      '10',
      '11',
      '12',
      '13',
    ]);
  });

  test('Arabic with digits: the digits read left to right inside the Arabic', () => {
    const item = lay(ARABIC_LATIN_DIGITS);
    const line = item.text!.lines[0];
    const digits = line.runs.find((run) => run.text === '42')!;
    // `4` (offset 14) starts at the digits' left edge; after `2` (16) is their right edge.
    expect(caretXOn(line, 14)).toBe(digits.x[0]);
    expect(caretXOn(line, 16, 'upstream')).toBe(digits.x[2]);
  });
});

describe('selection bands', () => {
  const doc = new Window().document as unknown as Document;
  const bands = (item: ShapeItem, from: number, to: number) => {
    const [g] = selectionRects(
      doc,
      item,
      { paragraph: 0, offset: from },
      { paragraph: 0, offset: to },
      'sel'
    );
    return [...g.querySelectorAll('rect')].map((rect) => {
      const x = Number(rect.getAttribute('x'));
      return [x, x + Number(rect.getAttribute('width'))];
    });
  };

  test('a range across a direction boundary covers one band for each visual stretch', () => {
    const item = lay(HEBREW_LATIN);
    const line = item.text!.lines[0];
    const word = line.runs.find((run) => run.text === 'PPTX')!;
    const hebrew = line.runs.find((run) => run.text === 'אבג')!;
    // `אבג ` and `P`: the Hebrew word with its space, and the first letter of the Latin word.
    expect(bands(item, 0, 5)).toEqual([
      [word.x[0], word.x[1]],
      [word.x[4], hebrew.x[3]],
    ]);
    // The whole Latin word is one band; reversed ends give the same bands.
    expect(bands(item, 4, 8)).toEqual([[word.x[0], word.x[4]]]);
    expect(bands(item, 8, 4)).toEqual(bands(item, 4, 8));
    // The whole line is one band.
    expect(bands(item, 0, 13)).toEqual([[line.caretX[13], line.caretX[0]]]);
  });

  test('Hebrew with a number: a range from the number into the Hebrew is two bands', () => {
    const item = lay(HEBREW_NUMBER);
    const line = item.text!.lines[0];
    const number = line.runs.find((run) => run.text === '120.50')!;
    // `.50 ` (offsets 8-12) and `ת` (12-13).
    const result = bands(item, 8, 13);
    expect(result).toHaveLength(2);
    expect(result[1]![0]).toBe(number.x[3]);
  });

  test('a left-to-right line keeps one rectangle from caret to caret', () => {
    const item = lay(ltr(r('plain text')));
    const line = item.text!.lines[0];
    expect(bands(item, 2, 7)).toEqual([[line.caretX[2], line.caretX[7]]]);
  });
});

describe('caret keys on a line in visual order', () => {
  const walk = (item: ShapeItem, from: TextPoint, key: string, mods = {}): TextPoint[] => {
    const out: TextPoint[] = [from];
    for (let i = 0; i < 40; i += 1) {
      const next = moveCaret(item, out[out.length - 1]!, key, mods);
      const prev = out[out.length - 1]!;
      if (next.offset === prev.offset && next.affinity === prev.affinity) break;
      out.push(next);
    }
    return out;
  };

  for (const [name, body] of CASES)
    test(`${name}: Left and Right visit every place in screen order`, () => {
      const item = lay(body);
      const line = item.text!.lines[0];
      const stops = lineStopsOf(item, line);
      const right = stops[stops.length - 1]!;
      const left = stops[0]!;
      const leftward = walk(item, { paragraph: 0, ...right }, 'ArrowLeft');
      expect(leftward.map((p) => caretX(item, p))).toEqual([...stops.map((s) => s.x)].reverse());
      const rightward = walk(item, { paragraph: 0, ...left }, 'ArrowRight');
      expect(rightward.map((p) => caretX(item, p))).toEqual(stops.map((s) => s.x));
      // Cmd with Left or Right goes to the visual edge on the arrow's side.
      const meta = { metaKey: true };
      const at = { paragraph: 0, offset: 2 };
      expect(caretX(item, moveCaret(item, at, 'ArrowLeft', meta))).toBe(left.x);
      expect(caretX(item, moveCaret(item, at, 'ArrowRight', meta))).toBe(right.x);
      // With Shift, the range grows to that edge from its anchor.
      const sel = { kind: 'text' as const, nodeId: 'n', anchor: at, focus: at };
      const grown = moveTextSelection(item, sel, 'ArrowLeft', true, meta);
      expect(grown.anchor).toEqual(at);
      expect(caretX(item, grown.focus)).toBe(left.x);
      // Home and End go to the logical start and end of the line.
      expect(moveCaret(item, at, 'Home').offset).toBe(line.start);
      expect(moveCaret(item, at, 'End').offset).toBe(line.end);
    });

  describe('Home and End on lines that start or end in the other direction', () => {
    const cases = [
      ['a left-to-right line that starts in Hebrew', ltr(r('שלום world', 'he-IL')), 10],
      ['a right-to-left line of Latin text', rtl(r('abc def')), 7],
      ['a right-to-left line that ends in Latin', rtl(r('אבג דהו abc', 'he-IL')), 11],
      ['a right-to-left line that starts in Latin', rtl(r('PPTX') + r(' אבג', 'he-IL')), 8],
    ] as const;
    for (const [name, body, length] of cases)
      test(`${name}: Home is offset 0 and End the text end`, () => {
        const item = lay(body);
        const line = item.text!.lines[0];
        for (let offset = 0; offset <= length; offset += 1) {
          const at = { paragraph: 0, offset };
          const home = moveCaret(item, at, 'Home');
          expect(home).toEqual({ paragraph: 0, offset: 0 });
          expect(caretX(item, home)).toBe(line.caretX[0]);
          const end = moveCaret(item, at, 'End');
          expect(end.offset).toBe(length);
          // The caret shows after the last character, at its trailing edge.
          expect(caretX(item, end)).toBe(line.bidi!.trailingX[length - 1]);
          const shifted = moveTextSelection(
            item,
            { kind: 'text', nodeId: 'n', anchor: at, focus: at },
            'End',
            true
          );
          expect([shifted.anchor, shifted.focus.offset]).toEqual([at, length]);
        }
      });

    test('End before the spaces of a wrap shows after the last word', () => {
      const words = 'אבג דהו זחט טיכ למנ סעפ צקר שתא בגד';
      const item = lay(rtl(r('PPTX', 'en-US') + r(` ${words}`, 'he-IL')), 1_400_000);
      const line = item.text!.lines[0];
      expect(item.text!.lines.length).toBeGreaterThan(1);
      const end = moveCaret(item, { paragraph: 0, offset: 0 }, 'End');
      expect(end.offset).toBeLessThan(line.end);
      expect(caretX(item, end)).toBe(line.bidi!.trailingX[end.offset - 1 - line.start]);
    });
  });

  test('Shift extends a selection on screen; the anchor stays', () => {
    const item = lay(HEBREW_LATIN);
    const at = { paragraph: 0, offset: 3 };
    const sel = { kind: 'text' as const, nodeId: 'n', anchor: at, focus: at };
    const one = moveTextSelection(item, sel, 'ArrowLeft', true);
    expect(one.anchor).toEqual(at);
    expect(one.focus).toEqual({ paragraph: 0, offset: 8, affinity: 'upstream' });
    const end = moveTextSelection(item, one, 'End', true);
    expect(end.anchor).toEqual(at);
    expect(end.focus).toEqual({ paragraph: 0, offset: 13 });
    const word = moveTextSelection(item, end, 'ArrowRight', true, { altKey: true });
    expect(word.anchor).toEqual(at);
    expect(word.focus.offset).toBe(9);
  });

  test('Ctrl or Alt with an arrow walks by word on screen', () => {
    const item = lay(HEBREW_LATIN);
    const word = (p: TextPoint, key: string) => moveCaret(item, p, key, { ctrlKey: true });
    // From the right edge, through `אבג`: the walk stops before the jump to the Latin word.
    expect(word({ paragraph: 0, offset: 0 }, 'ArrowLeft')).toEqual({ paragraph: 0, offset: 3 });
    // Across the Latin word, to its left edge.
    expect(word({ paragraph: 0, offset: 3 }, 'ArrowLeft').offset).toBe(4);
    // Back to the right: the Latin word's right edge, then the Hebrew word's start.
    const back = word({ paragraph: 0, offset: 4 }, 'ArrowRight');
    expect(caretX(item, back)).toBe(item.text!.lines[0].runs.find((x) => x.text === 'PPTX')!.x[4]);
    expect(word({ paragraph: 0, offset: 3 }, 'ArrowRight')).toEqual({ paragraph: 0, offset: 0 });
  });

  test('a plain arrow collapses a range to its outer band edge on the key side', () => {
    const item = lay(HEBREW_LATIN);
    const sel = {
      kind: 'text' as const,
      nodeId: 'n',
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 5 },
    };
    expect(moveTextSelection(item, sel, 'ArrowLeft', false).focus.offset).toBe(4);
    expect(moveTextSelection(item, sel, 'ArrowRight', false).focus.offset).toBe(0);
    // `אבג ` only: Left collapses to the band's left edge, the right edge of the Latin word.
    const hebrew = { ...sel, focus: { paragraph: 0, offset: 4 } };
    const left = moveTextSelection(item, hebrew, 'ArrowLeft', false).focus;
    expect(caretX(item, left)).toBe(item.text!.lines[0].runs.find((x) => x.text === 'PPTX')!.x[4]);
  });

  test('past the line edge, the keys follow reading order into the next paragraph', () => {
    const item = lay(HEBREW_LATIN + ltr(r('next')));
    // Left at the left edge of the right-to-left paragraph reads on into the next one.
    expect(moveCaret(item, { paragraph: 0, offset: 13 }, 'ArrowLeft')).toEqual({
      paragraph: 1,
      offset: 0,
    });
    // Right at its right edge goes back to the paragraph before: none, so the caret stays.
    expect(moveCaret(item, { paragraph: 0, offset: 0 }, 'ArrowRight')).toEqual({
      paragraph: 0,
      offset: 0,
    });
    // From the left-to-right paragraph, Left goes to the end of the paragraph before.
    expect(moveCaret(item, { paragraph: 1, offset: 0 }, 'ArrowLeft')).toEqual({
      paragraph: 0,
      offset: 13,
    });
  });

  test('a wrapped right-to-left paragraph steps from the left edge onto the next line', () => {
    const item = lay(rtl(r('אבג דהו זחט טיכ למנ סעפ צקר שתא בגד', 'he-IL')), 1_400_000);
    const lines = item.text!.lines;
    expect(lines.length).toBeGreaterThan(1);
    const first = lineStopsOf(item, lines[0]);
    const next = moveCaret(item, { paragraph: 0, ...first[0]! }, 'ArrowLeft');
    // The entry is the right edge of the second line.
    const second = lineStopsOf(item, lines[1]);
    expect(next.offset).toBe(second[second.length - 1]!.offset);
    // Up from there keeps the x on the first line.
    const up = moveCaret(item, next, 'ArrowUp');
    expect(up.paragraph).toBe(0);
    expect(up.offset).toBeLessThan(lines[0].end);
  });
});

describe('hit tests on a line in visual order', () => {
  for (const [name, body] of CASES)
    test(`${name}: a press at each place gives that place`, () => {
      const item = lay(body);
      for (const s of lineStopsOf(item, item.text!.lines[0])) {
        const p = slidePoint(item, s.x);
        const at = pointAt(item, p.x, p.y)!;
        expect(caretX(item, at)).toBe(s.x);
        expect(at.offset).toBe(s.offset);
      }
    });

  test('beside the Latin word, a press chooses the side of the boundary', () => {
    const item = lay(HEBREW_LATIN);
    const word = item.text!.lines[0].runs.find((run) => run.text === 'PPTX')!;
    const right = slidePoint(item, word.x[4] + 1000);
    expect(pointAt(item, right.x, right.y)).toEqual({
      paragraph: 0,
      offset: 8,
      affinity: 'upstream',
    });
    const left = slidePoint(item, word.x[0] - 1000);
    expect(pointAt(item, left.x, left.y)).toEqual({ paragraph: 0, offset: 4 });
  });
});

describe('typing on a line in visual order', () => {
  test('the caret stays after the typed character at a direction boundary', async () => {
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const bytes = buildPptx({ slides: [{ shapes: shape(HEBREW_LATIN), layout: 'blank' }] });
    expect((await e.load(bytes)).ok).toBe(true);
    const item = e.session!.layout(0)!.items.find((i) => i.kind === 'shape') as ShapeItem;
    const nodeId = item.element.nodeId;
    const at = { paragraph: 0, offset: 8, affinity: 'upstream' as const };
    e.setSelection({ kind: 'text', nodeId, anchor: at, focus: at });
    expect(e.input('text', 'Y').ok).toBe(true);
    const sel = e.snapshot().selection;
    if (sel.kind !== 'text') throw new Error('no text selection');
    expect(sel.focus).toEqual({ paragraph: 0, offset: 9, affinity: 'upstream' });
    const next = e.session!.layout(0)!.items.find((i) => i.kind === 'shape') as ShapeItem;
    const word = next.text!.lines[0].runs.find((run) => run.text === 'PPTXY')!;
    expect(caretX(next, sel.focus)).toBe(word.x[5]);
    e.destroy();
  });
});

describe('the surface on a line in visual order', () => {
  test('keys move the painted caret on screen, and a range paints a band for each stretch', async () => {
    const win = new GlobalWindow();
    const g = globalThis as { ResizeObserver?: unknown };
    g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const bytes = buildPptx({ slides: [{ shapes: shape(HEBREW_LATIN), layout: 'blank' }] });
    expect((await e.load(bytes)).ok).toBe(true);
    const doc = win.document as unknown as Document;
    const host = doc.createElement('div');
    doc.body.append(host);
    e.attach(host);
    const input = host.querySelector('textarea') as HTMLTextAreaElement;
    const key = (k: string, shiftKey = false): void => {
      const ev = new win.KeyboardEvent('keydown', { key: k, bubbles: true, shiftKey });
      input.dispatchEvent(ev as unknown as Event);
    };
    const item = e.session!.layout(0)!.items.find((i) => i.kind === 'shape') as ShapeItem;
    const nodeId = item.element.nodeId;
    const word = item.text!.lines[0].runs.find((run) => run.text === 'PPTX')!;
    const m = textFrameMatrix(item)!;
    const at = { paragraph: 0, offset: 3 };
    e.setSelection({ kind: 'text', nodeId, anchor: at, focus: at });
    key('ArrowLeft');
    let sel = e.snapshot().selection;
    expect(sel.kind === 'text' && sel.focus).toEqual({
      paragraph: 0,
      offset: 8,
      affinity: 'upstream',
    });
    const line = host.querySelector('.pptx-surface__caret');
    expect(Number(line?.getAttribute('x1'))).toBeCloseTo(apply(m, word.x[4], 0).x, 0);
    // Shift+Left twice selects `X` and `T`, from the anchor at offset 8.
    key('ArrowLeft', true);
    key('ArrowLeft', true);
    sel = e.snapshot().selection;
    expect(sel.kind === 'text' && [sel.anchor.offset, sel.focus.offset]).toEqual([8, 6]);
    e.setSelection({
      kind: 'text',
      nodeId,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 5 },
    });
    expect(host.querySelectorAll('.pptx-surface__text-selection').length).toBe(2);
    e.destroy();
    await win.happyDOM.close();
  });
});
