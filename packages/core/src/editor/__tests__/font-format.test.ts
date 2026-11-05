// Superscript and subscript, Change Case, character spacing, and the Font dialog's command. Each
// writes the schema value, applies to the selected text only, lays out, undoes as one
// step, and survives save and reopen.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { attr, child, elements, find, textOf, xmlPartOf, type Element } from '../../store/index.ts';
import {
  createPptxEditor,
  runToolbarCommand,
  toolbarCommandState,
  type PptxEditorInstance,
} from '../index.ts';
import { changeCaseText } from '../../automation/commands/font-format.ts';

const IN = 914_400;

function deck(text: string): Uint8Array {
  return buildPptx({
    slides: [{ shapes: shape({ id: 2, name: 'T', x: IN, y: IN, cx: 6 * IN, cy: 2 * IN, text }) }],
  });
}

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const root = (e: PptxEditorInstance): Element =>
  xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!.root;
const shapeId = (e: PptxEditorInstance): string =>
  elements(find(root(e), 'p:spTree')!).find((c) => attr(find(c, 'p:cNvPr'), 'name') === 'T')!.id;
const runs = (e: PptxEditorInstance, p = 0): Element[] =>
  elements(elements(find(root(e), 'p:txBody')!).filter((c) => c.local === 'p')[p]!).filter(
    (c) => c.local === 'r'
  );
const texts = (e: PptxEditorInstance, p = 0): string[] =>
  runs(e, p).map((r) => textOf(child(r, 'a:t')));
const rPr = (r: Element, name: string): string | undefined => attr(child(r, 'a:rPr'), name);
const select = (
  e: PptxEditorInstance,
  from: number,
  to: number,
  paragraph = 0,
  toParagraph = paragraph
) =>
  e.setSelection({
    kind: 'text',
    nodeId: shapeId(e),
    anchor: { paragraph, offset: from },
    focus: { paragraph: toParagraph, offset: to },
  });

describe('superscript and subscript', () => {
  test('write the preset baselines on the selected text, and toggle off', async () => {
    const e = await open(deck(para(run('H2O x2'))));
    select(e, 1, 2);
    expect(e.exec({ type: 'toggleSubscript' }).ok).toBe(true);
    expect(texts(e)).toEqual(['H', '2', 'O x2']);
    expect(rPr(runs(e)[1]!, 'baseline')).toBe('-25000');
    expect(rPr(runs(e)[0]!, 'baseline')).toBeUndefined();
    expect(e.snapshot().format?.subscript).toBe(true);
    expect(toolbarCommandState(e, 'script.sub').active).toBe(true);
    // Superscript replaces subscript.
    expect(e.exec({ type: 'toggleSuperscript' }).ok).toBe(true);
    expect(rPr(runs(e)[1]!, 'baseline')).toBe('30000');
    expect(toolbarCommandState(e, 'script.super').active).toBe(true);
    expect(toolbarCommandState(e, 'script.sub').active).toBe(false);
    expect(e.exec({ type: 'toggleSuperscript' }).ok).toBe(true);
    expect(rPr(runs(e)[1]!, 'baseline')).toBe('0');
    expect(e.snapshot().format?.superscript).toBe(false);
  });

  test('raise the run in layout and shrink it, and undo as one step each', async () => {
    const e = await open(deck(para(run('ab'))));
    const glyphs = () => {
      const item = e.session!.layout(0)!.items.find((x) => x.kind === 'shape' && x.text)!;
      return item.kind === 'shape' ? item.text!.lines[0]!.runs : [];
    };
    select(e, 1, 2);
    e.exec({ type: 'toggleSuperscript' });
    const [plain, raised] = glyphs();
    expect(raised!.rise).toBeGreaterThan(0);
    expect(raised!.sizePt).toBeLessThan(plain!.sizePt);
    e.exec({ type: 'undo' });
    expect(texts(e)).toEqual(['ab']);
    expect(e.snapshot().canUndo).toBe(false);
  });
});

describe('font size stepper', () => {
  test('the step buttons grow and shrink the size; a number sets it', async () => {
    const e = await open(deck(para(run('Size', 'sz="1800"'))));
    select(e, 0, 4);
    expect(runToolbarCommand(e, 'font.size', { step: 1 }).ok).toBe(true);
    const grown = Number(rPr(runs(e)[0]!, 'sz'));
    expect(grown).toBeGreaterThan(1800);
    expect(runToolbarCommand(e, 'font.size', { step: -1 }).ok).toBe(true);
    expect(rPr(runs(e)[0]!, 'sz')).toBe('1800');
    expect(runToolbarCommand(e, 'font.size', 1).ok).toBe(true);
    expect(rPr(runs(e)[0]!, 'sz')).toBe('100');
    expect(runToolbarCommand(e, 'font.size', { step: 2 } as never).ok).toBe(false);
  });
});

describe('character spacing', () => {
  test('writes spc in hundredths of a point and widens the text', async () => {
    const e = await open(deck(para(run('spread'))));
    const width = () => {
      const item = e.session!.layout(0)!.items.find((x) => x.kind === 'shape' && x.text)!;
      const xs = item.kind === 'shape' ? item.text!.lines[0]!.caretX : [];
      return xs[xs.length - 1]! - xs[0]!;
    };
    const before = width();
    select(e, 0, 6);
    expect(runToolbarCommand(e, 'text.characterSpacing', 3).ok).toBe(true);
    expect(rPr(runs(e)[0]!, 'spc')).toBe('300');
    expect(Math.round(width() - before)).toBe(6 * 3 * 12_700);
    expect(toolbarCommandState(e, 'text.characterSpacing').value).toBe('3');
    expect(e.exec({ type: 'setCharacterSpacing', points: -1.5 }).ok).toBe(true);
    expect(rPr(runs(e)[0]!, 'spc')).toBe('-150');
    expect(e.snapshot().format?.font?.characterSpacingPt).toBe(-1.5);
    for (const points of [Number.NaN, 2000, -2000])
      expect(e.exec({ type: 'setCharacterSpacing', points }).ok).toBe(false);
  });
});

describe('change case', () => {
  test('text rules keep the length and find sentences and words in the whole text', () => {
    const text = "it's done. NEXT one! ok";
    expect(changeCaseText(text, 0, text.length, 'sentence')).toBe("It's done. Next one! Ok");
    expect(changeCaseText(text, 0, text.length, 'capitalize')).toBe("It's Done. Next One! Ok");
    expect(changeCaseText(text, 0, text.length, 'upper')).toBe("IT'S DONE. NEXT ONE! OK");
    expect(changeCaseText(text, 0, text.length, 'lower')).toBe("it's done. next one! ok");
    expect(changeCaseText('aBc', 0, 3, 'toggle')).toBe('AbC');
    // Mid-sentence, a range keeps the sentence going: no capital.
    expect(changeCaseText('one TWO', 4, 7, 'sentence')).toBe('one two');
    // A case change that changes the length (ß to SS) keeps the character.
    expect(changeCaseText('straße', 0, 6, 'upper')).toBe('STRAßE');
  });

  test('changes only the selected letters and keeps each run formatting', async () => {
    const e = await open(deck(para(run('hello ') + run('big', 'b="1"') + run(' world'))));
    select(e, 3, 12);
    expect(runToolbarCommand(e, 'text.changeCase', 'upper').ok).toBe(true);
    expect(texts(e)).toEqual(['helLO ', 'BIG', ' WOrld']);
    expect(rPr(runs(e)[1]!, 'b')).toBe('1');
    e.exec({ type: 'undo' });
    expect(texts(e)).toEqual(['hello ', 'big', ' world']);
  });

  test('a caret inside a word changes that word; a caret between words changes nothing', async () => {
    const e = await open(deck(para(run('ONE  two')) + para(run('three'))));
    select(e, 2, 2);
    expect(e.exec({ type: 'changeCase', case: 'capitalize' }).ok).toBe(true);
    expect(texts(e, 0)).toEqual(['One  two']);
    expect(texts(e, 1)).toEqual(['three']);
    select(e, 4, 4);
    expect(e.exec({ type: 'changeCase', case: 'upper' })).toMatchObject({
      ok: false,
      reason: 'disabled.selectWord',
    });
    expect(e.exec({ type: 'changeCase', case: 'shout' as never }).ok).toBe(false);
  });

  test('a caret at a word end takes the word before it, at a word start the word after', async () => {
    const e = await open(deck(para(run('one two  x'))));
    // Right after "one".
    select(e, 3, 3);
    expect(e.exec({ type: 'changeCase', case: 'upper' }).ok).toBe(true);
    expect(texts(e)).toEqual(['ONE two  x']);
    // Right before "two".
    select(e, 4, 4);
    expect(e.exec({ type: 'changeCase', case: 'upper' }).ok).toBe(true);
    expect(texts(e)).toEqual(['ONE TWO  x']);
    // A one-letter word at the end of the text.
    select(e, 10, 10);
    expect(e.exec({ type: 'changeCase', case: 'upper' }).ok).toBe(true);
    expect(texts(e)).toEqual(['ONE TWO  X']);
    // Between two spaces nothing is reached: Change Case and Insert link are refused.
    select(e, 8, 8);
    expect(e.can({ type: 'changeCase', case: 'lower' })).toMatchObject({
      ok: false,
      reason: 'disabled.selectWord',
    });
    expect(e.can({ type: 'setLink', url: 'https://x.test' })).toMatchObject({
      ok: false,
      reason: 'disabled.selectText',
    });
    // At the end of "TWO", Insert link links that word.
    select(e, 7, 7);
    expect(e.exec({ type: 'setLink', url: 'https://x.test' }).ok).toBe(true);
    expect(texts(e)).toEqual(['ONE ', 'TWO', '  X']);
  });
});

describe('superscript off with an inherited offset', () => {
  test('writes baseline="0", so a list style superscript turns off', async () => {
    const e = await open(
      buildPptx({
        slides: [
          {
            shapes: shape({
              id: 2,
              name: 'T',
              x: IN,
              y: IN,
              cx: 6 * IN,
              cy: 2 * IN,
              lstStyle: '<a:lvl1pPr><a:defRPr baseline="30000"/></a:lvl1pPr>',
              text: para(run('raised')),
            }),
          },
        ],
      })
    );
    select(e, 0, 6);
    expect(e.snapshot().format?.superscript).toBe(true);
    expect(e.exec({ type: 'toggleSuperscript' }).ok).toBe(true);
    expect(rPr(runs(e)[0]!, 'baseline')).toBe('0');
    expect(e.snapshot().format?.superscript).toBe(false);
  });
});

describe('formatting at a caret', () => {
  test('inside a word, the word is formatted, not the whole text body', async () => {
    const e = await open(deck(para(run('E = mc2 here')) + para(run('Second'))));
    select(e, 6, 6);
    expect(e.exec({ type: 'toggleSuperscript' }).ok).toBe(true);
    expect(texts(e)).toEqual(['E = ', 'mc2', ' here']);
    expect(rPr(runs(e)[1]!, 'baseline')).toBe('30000');
    expect(rPr(runs(e)[0]!, 'baseline')).toBeUndefined();
    expect(runs(e, 1).map((r) => rPr(r, 'baseline'))).toEqual([undefined]);
  });

  test('at a word edge, the format is armed for the text typed next, and undoes in two steps', async () => {
    const e = await open(deck(para(run('Water is H')) + para(run('Second'))));
    select(e, 10, 10);
    expect(e.exec({ type: 'toggleSubscript' }).ok).toBe(true);
    // Nothing typed yet: the document does not change (no store step; Undo ends the arm), and
    // the button shows the armed state.
    expect(texts(e)).toEqual(['Water is H']);
    expect(e.session!.store.canUndo).toBe(false);
    expect(e.snapshot().format?.subscript).toBe(true);
    expect(e.input('text', '2').ok).toBe(true);
    expect(texts(e)).toEqual(['Water is H', '2']);
    expect(rPr(runs(e)[1]!, 'baseline')).toBe('-25000');
    expect(runs(e, 1).map((r) => rPr(r, 'baseline'))).toEqual([undefined]);
    // As in DOCX Editor: the first undo removes the text, and the format is armed again at the
    // caret. The second undo ends the arm.
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(texts(e)).toEqual(['Water is H']);
    expect(e.snapshot().format?.subscript).toBe(true);
    expect(e.snapshot().canUndo).toBe(true);
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(e.snapshot().canUndo).toBe(false);
    expect(e.snapshot().format?.subscript).toBe(false);
    // Redo brings back the text with its format.
    expect(e.exec({ type: 'redo' }).ok).toBe(true);
    expect(texts(e)).toEqual(['Water is H', '2']);
    expect(rPr(runs(e)[1]!, 'baseline')).toBe('-25000');
  });

  test('a font family picked at a caret undoes after the text typed with it', async () => {
    const e = await open(deck(para(run('PPTX Editor'))));
    select(e, 5, 5);
    expect(e.exec({ type: 'setFontFamily', family: 'Arial' }).ok).toBe(true);
    expect(e.input('text', 'Z').ok).toBe(true);
    expect(e.input('text', 'q').ok).toBe(true);
    expect(texts(e).join('')).toBe('PPTX ZqEditor');
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(texts(e).join('')).toBe('PPTX Editor');
    expect(e.snapshot().format?.fontFamily).toBe('Arial');
    expect(e.snapshot().canUndo).toBe(true);
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(e.snapshot().format?.fontFamily).not.toBe('Arial');
    expect(e.snapshot().canUndo).toBe(false);
  });

  test('an undo back to where plain typing started does not arm an old format', async () => {
    const e = await open(deck(para(run('Hello'))));
    select(e, 5, 5);
    e.exec({ type: 'toggleBold' });
    e.input('text', 'a');
    e.exec({ type: 'undo' });
    e.exec({ type: 'undo' });
    // The arm ended; new typing from the same package makes its own history entry.
    e.input('text', 'b');
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(texts(e).join('')).toBe('Hello');
    expect(e.snapshot().format?.bold).toBe(false);
    expect(e.snapshot().canUndo).toBe(false);
  });

  test('the armed format ends when the caret moves, and never reaches the file', async () => {
    const e = await open(deck(para(run('Water is H'))));
    select(e, 10, 10);
    e.exec({ type: 'toggleBold' });
    expect(e.snapshot().format?.bold).toBe(true);
    // Saved while armed: no empty run is written.
    const again = await open(e.save()!);
    expect(texts(again)).toEqual(['Water is H']);
    select(e, 0, 0);
    select(e, 10, 10);
    expect(e.snapshot().format?.bold).toBe(false);
    e.input('text', 'Z');
    expect(texts(e)).toEqual(['Water is HZ']);
    expect(rPr(runs(e)[0]!, 'b')).toBeUndefined();
  });

  test('Enter carries the arm: a new empty paragraph takes it on its paragraph mark', async () => {
    const e = await open(deck(para(run('Hello'))));
    select(e, 5, 5);
    e.exec({ type: 'toggleBold' });
    expect(e.input('paragraph').ok).toBe(true);
    // Nothing empty is left in the first paragraph; the new paragraph mark is bold.
    expect(texts(e, 0)).toEqual(['Hello']);
    const p1 = elements(find(root(e), 'p:txBody')!).filter((c) => c.local === 'p')[1]!;
    expect(attr(child(p1, 'a:endParaRPr'), 'b')).toBe('1');
    expect(e.snapshot().format?.bold).toBe(true);
    // Bold turns off there, and the text typed next follows.
    e.exec({ type: 'toggleBold' });
    expect(e.snapshot().format?.bold).toBe(false);
    e.input('text', 'x');
    expect(rPr(runs(e, 1)[0]!, 'b')).not.toBe('1');
    // A newline typed as text acts as Enter.
    const n = await open(deck(para(run('Hi'))));
    select(n, 2, 2);
    n.exec({ type: 'toggleItalic' });
    n.input('text', '\n');
    expect(texts(n, 0)).toEqual(['Hi']);
  });

  test('Shift+Enter, Backspace, and Delete keep the arm as session state', async () => {
    const e = await open(deck(para(run('Hello'))));
    select(e, 5, 5);
    e.exec({ type: 'toggleBold' });
    expect(e.input('lineBreak').ok).toBe(true);
    expect(e.snapshot().format?.bold).toBe(true);
    // No empty run is written before the break.
    expect(texts(e)).toEqual(['Hello']);
    e.input('text', 'x');
    expect(
      rPr(
        runs(e).find((r) => textOf(child(r, 'a:t')) === 'x')!,
        'b'
      )
    ).toBe('1');
    const d = await open(deck(para(run('Hello'))));
    select(d, 5, 5);
    d.exec({ type: 'toggleItalic' });
    expect(d.input('deleteBackward').ok).toBe(true);
    expect(d.snapshot().format?.italic).toBe(true);
    d.input('text', 'p');
    expect(texts(d)).toEqual(['Hell', 'p']);
    expect(rPr(runs(d)[1]!, 'i')).toBe('1');
  });

  test('an armed toggle or size step stays what it was when it armed, after Enter and Backspace', async () => {
    // Bold armed off after a bold word, then Backspace onto plain text: still off.
    const a = await open(deck(para(run('Hello ') + run('W', 'b="1"'))));
    select(a, 7, 7);
    a.exec({ type: 'toggleBold' });
    expect(a.snapshot().format?.bold).toBe(false);
    a.input('deleteBackward');
    expect(a.snapshot().format?.bold).toBe(false);
    a.input('text', 'x');
    expect(
      rPr(
        runs(a).find((r) => textOf(child(r, 'a:t')) === 'x')!,
        'b'
      )
    ).toBe('0');
    // Bold armed on before a bold run, then Enter: still on.
    const b = await open(deck(para(run('Hello') + run(' world', 'b="1"'))));
    select(b, 5, 5);
    b.exec({ type: 'toggleBold' });
    b.input('paragraph');
    expect(b.snapshot().format?.bold).toBe(true);
    // A size step armed at 40 pt stays 44 pt after Enter onto 10 pt text.
    const c = await open(deck(para(run('Big', 'sz="4000"') + run(' small', 'sz="1000"'))));
    select(c, 3, 3);
    c.exec({ type: 'stepFontSize', direction: 1 });
    const armed = c.snapshot().format?.fontSize;
    c.input('lineBreak');
    expect(c.snapshot().format?.fontSize).toBe(armed);
  });

  test('a deletion keeps the arm inside a word, and an emptied paragraph takes it on its mark', async () => {
    const e = await open(deck(para(run('ab cd'))));
    select(e, 3, 3);
    e.exec({ type: 'toggleItalic' });
    e.input('deleteBackward');
    expect(e.snapshot().format?.italic).toBe(true);
    e.input('text', 'x');
    expect(texts(e)).toEqual(['ab', 'x', 'cd']);
    expect(rPr(runs(e)[1]!, 'i')).toBe('1');
    const f = await open(deck(para(run('a'))));
    select(f, 1, 1);
    f.exec({ type: 'toggleBold' });
    f.input('deleteBackward');
    const p0 = elements(find(root(f), 'p:txBody')!).filter((c) => c.local === 'p')[0]!;
    expect(attr(child(p0, 'a:endParaRPr'), 'b')).toBe('1');
    f.input('text', 'x');
    expect(rPr(runs(f)[0]!, 'b')).toBe('1');
  });

  test('Undo is on while an arm is held, and the size steps refuse a mixed size', async () => {
    const e = await open(deck(para(run('Hi '))));
    select(e, 3, 3);
    expect(e.can({ type: 'undo' }).ok).toBe(false);
    e.exec({ type: 'toggleBold' });
    expect(e.snapshot().canUndo).toBe(true);
    expect(e.can({ type: 'undo' }).ok).toBe(true);
    const m = await open(deck(para(run('Big', 'sz="2800"') + run('small', 'sz="1400"'))));
    select(m, 0, 8);
    expect(m.exec({ type: 'stepFontSize', direction: 1 })).toMatchObject({
      ok: false,
      reason: 'disabled.mixedSize',
    });
  });

  test('the first undo ends the arm alone, as in DOCX Editor', async () => {
    const e = await open(deck(para(run('Hello'))));
    select(e, 5, 5);
    e.input('text', 'abc');
    e.exec({ type: 'toggleBold' });
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(texts(e).join('')).toBe('Helloabc');
    expect(e.snapshot().format?.bold).toBe(false);
    e.exec({ type: 'undo' });
    expect(texts(e).join('')).toBe('Hello');
  });

  test('an arm on one slide does not leak into the toggle state of another', async () => {
    const e = await open(
      buildPptx({
        slides: [
          {
            shapes: shape({
              id: 2,
              name: 'T',
              x: IN,
              y: IN,
              cx: 4 * IN,
              cy: IN,
              text: para(run('Bold', 'b="1"')),
            }),
          },
          {
            shapes: shape({
              id: 2,
              name: 'T',
              x: IN,
              y: IN,
              cx: 4 * IN,
              cy: IN,
              text: para(run('Plain')),
            }),
          },
        ],
      })
    );
    const caretAt = (offset: number) => {
      const nodeId = e.session!.layout(e.snapshot().currentSlide)!.items[0]!.element.nodeId;
      e.setSelection({
        kind: 'text',
        nodeId,
        anchor: { paragraph: 0, offset },
        focus: { paragraph: 0, offset },
      });
    };
    e.goToSlide(1);
    caretAt(5);
    e.exec({ type: 'toggleItalic' });
    e.goToSlide(0);
    caretAt(4);
    // Bold text: the toggle turns bold off for the text typed next.
    e.exec({ type: 'toggleBold' });
    e.input('text', 'x');
    expect(
      rPr(
        runs(e).find((r) => textOf(child(r, 'a:t')) === 'x')!,
        'b'
      )
    ).toBe('0');
  });

  test('two toggles at one caret cancel, and other formats add up', async () => {
    const e = await open(deck(para(run('Hi '))));
    select(e, 3, 3);
    e.exec({ type: 'toggleBold' });
    e.exec({ type: 'toggleItalic' });
    e.exec({ type: 'toggleBold' });
    expect([e.snapshot().format?.bold, e.snapshot().format?.italic]).toEqual([false, true]);
    e.input('text', 'x');
    expect(texts(e)).toEqual(['Hi ', 'x']);
    expect([rPr(runs(e)[1]!, 'b'), rPr(runs(e)[1]!, 'i')]).toEqual(['0', '1']);
  });

  test('in an empty paragraph the paragraph mark takes the format at once', async () => {
    const e = await open(deck(para(run('Top')) + '<a:p><a:endParaRPr lang="en-US"/></a:p>'));
    select(e, 0, 0, 1);
    expect(e.exec({ type: 'setFontSize', points: 40 }).ok).toBe(true);
    expect(e.snapshot().canUndo).toBe(true);
    const p1 = elements(find(root(e), 'p:txBody')!).filter((c) => c.local === 'p')[1]!;
    expect(attr(child(p1, 'a:endParaRPr'), 'sz')).toBe('4000');
    e.input('text', 'big');
    expect(rPr(runs(e, 1)[0]!, 'sz')).toBe('4000');
  });

  test('Clear Formatting at the caret of an empty paragraph clears its paragraph mark', async () => {
    const e = await open(
      deck(para(run('Top')) + '<a:p><a:endParaRPr lang="en-US" b="1" sz="4000"/></a:p>')
    );
    select(e, 0, 0, 1);
    expect(e.exec({ type: 'clearFormatting' }).ok).toBe(true);
    const p1 = elements(find(root(e), 'p:txBody')!).filter((c) => c.local === 'p')[1]!;
    const mark = child(p1, 'a:endParaRPr')!;
    expect([attr(mark, 'b'), attr(mark, 'sz'), attr(mark, 'lang')]).toEqual([
      undefined,
      undefined,
      'en-US',
    ]);
  });

  test('in a paragraph that holds only an empty run, the run takes the format too', async () => {
    const e = await open(
      deck(para(run('Top')) + '<a:p><a:r><a:rPr lang="en-US"/><a:t></a:t></a:r></a:p>')
    );
    select(e, 0, 0, 1);
    expect(e.exec({ type: 'toggleBold' }).ok).toBe(true);
    expect(e.snapshot().format?.bold).toBe(true);
    e.input('text', 'x');
    expect(rPr(runs(e, 1)[0]!, 'b')).toBe('1');
    expect(textOf(child(runs(e, 1)[0]!, 'a:t'))).toBe('x');
    expect(e.snapshot().format?.bold).toBe(true);
  });

  test('after a line break, next to a field, and between surrogate pairs', async () => {
    const e = await open(deck(para(run('ab') + '<a:br><a:rPr lang="en-US"/></a:br>' + run('cd'))));
    select(e, 3, 3);
    e.exec({ type: 'toggleBold' });
    e.input('text', 'X');
    expect(texts(e)).toEqual(['ab', 'X', 'cd']);
    expect(rPr(runs(e)[1]!, 'b')).toBe('1');
    const field =
      '<a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld>';
    const f = await open(deck(para(run('Page ') + field)));
    select(f, 5, 5);
    expect(f.exec({ type: 'toggleItalic' }).ok).toBe(true);
    expect(f.snapshot().format?.italic).toBe(true);
    const g = await open(deck(para(run('\u{1D4B3}\u{1D4B4} z'))));
    select(g, 2, 2);
    g.exec({ type: 'toggleBold' });
    expect(texts(g)).toEqual(['\u{1D4B3}\u{1D4B4}', ' z']);
    expect(rPr(runs(g)[0]!, 'b')).toBe('1');
  });

  test('a field ends a word, a Thai word is a dictionary word, and a range formats its field', async () => {
    const field =
      '<a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-US"/><a:t>12</a:t></a:fld>';
    const e = await open(deck(para(run('Page') + field + run(' end'))));
    // The caret between "Page" and the field is at a word edge: it arms, nothing changes.
    select(e, 4, 4);
    e.exec({ type: 'setFontSize', points: 40 });
    expect(rPr(runs(e)[0]!, 'sz')).toBeUndefined();
    select(e, 0, 9);
    e.exec({ type: 'setFontSize', points: 30 });
    const fld = elements(
      elements(find(root(e), 'p:txBody')!).filter((c) => c.local === 'p')[0]!
    ).find((c) => c.local === 'fld')!;
    expect(rPr(fld, 'sz')).toBe('3000');
    expect(e.snapshot().format?.font?.disagrees.fontSizePt).toBe(false);
    const thai = await open(deck(para(run('สวัสดีครับ'))));
    select(thai, 2, 2);
    thai.exec({ type: 'toggleBold' });
    expect(texts(thai)[0]).toBe('สวัสดี');
    expect(rPr(runs(thai)[0]!, 'b')).toBe('1');
  });

  test('a font family from the Font dialog is asked of the fonts origins', async () => {
    const asked: string[] = [];
    const e = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      fonts: (request) => (asked.push(...request.families), undefined),
    });
    expect((await e.load(deck(para(run('Hello'))))).ok).toBe(true);
    select(e, 0, 5);
    expect(e.exec({ type: 'setFontFormat', fontFamily: 'Zapfino Extra' }).ok).toBe(true);
    await new Promise((r) => setTimeout(r, 0));
    expect(asked).toContain('Zapfino Extra');
  });
});

describe('setFontFormat', () => {
  test('writes every Font dialog field as one undo step', async () => {
    const e = await open(deck(para(run('Effects'))));
    select(e, 0, 7);
    const r = e.exec({
      type: 'setFontFormat',
      fontFamily: 'Georgia',
      fontSizePt: 20,
      bold: true,
      italic: false,
      color: { theme: 'accent2' },
      underline: 'wavy',
      strike: 'double',
      baseline: 50,
      caps: 'small',
      characterSpacingPt: 1.2,
      kerningPt: 12,
      underlineColor: '#c00000',
    });
    expect(r.ok).toBe(true);
    const run0 = runs(e)[0]!;
    const props = child(run0, 'a:rPr')!;
    expect(Object.fromEntries(props.attrs.map((a) => [a.local, a.value]))).toMatchObject({
      sz: '2000',
      b: '1',
      i: '0',
      u: 'wavy',
      strike: 'dblStrike',
      baseline: '50000',
      cap: 'small',
      spc: '120',
      kern: '1200',
    });
    expect(elements(props).map((c) => c.local)).toEqual(['solidFill', 'uFill', 'latin']);
    expect(attr(find(child(props, 'a:uFill')!, 'a:srgbClr'), 'val')).toBe('C00000');
    expect(e.snapshot().format?.font).toEqual({
      fontFamily: 'Georgia',
      fontSizePt: 20,
      bold: true,
      italic: false,
      underline: 'wavy',
      strike: 'double',
      baseline: 50,
      caps: 'small',
      characterSpacingPt: 1.2,
      kerningPt: 12,
      underlineColor: '#C00000',
      disagrees: {
        fontFamily: false,
        fontSizePt: false,
        bold: false,
        italic: false,
        underline: false,
        strike: false,
        baseline: false,
        caps: false,
        characterSpacingPt: false,
        kerningPt: false,
        underlineColor: false,
      },
    });
    e.exec({ type: 'undo' });
    expect(
      child(runs(e)[0]!, 'a:rPr')!
        .attrs.map((a) => a.local)
        .sort()
    ).toEqual(['dirty', 'lang']);
    expect(e.snapshot().canUndo).toBe(false);
  });

  test('offset 0 writes baseline="0", and kerning off writes kern="0"', async () => {
    const e = await open(deck(para(run('x', 'baseline="30000" kern="1200"'))));
    select(e, 0, 1);
    expect(e.exec({ type: 'setFontFormat', baseline: 0, kerningPt: null }).ok).toBe(true);
    expect(rPr(runs(e)[0]!, 'baseline')).toBe('0');
    expect(rPr(runs(e)[0]!, 'kern')).toBe('0');
    expect(e.snapshot().format?.font?.kerningPt).toBeNull();
  });

  test('the underline takes its own color in layout; null returns it to the text color', async () => {
    const e = await open(
      deck(para(run('Under', 'u="sng"', '<a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>')))
    );
    select(e, 0, 5);
    const bars = (): string[] => {
      const item = e.session!.layout(0)!.items.find((x) => x.kind === 'shape' && x.text)!;
      const run0 = item.kind === 'shape' ? item.text!.lines[0]!.runs[0]! : undefined;
      return [run0?.props.underlineFill?.type ?? 'text'];
    };
    expect(bars()).toEqual(['text']);
    expect(e.snapshot().format?.font?.underlineColor).toBeNull();
    expect(e.exec({ type: 'setFontFormat', underlineColor: { theme: 'accent2' } }).ok).toBe(true);
    expect(bars()).toEqual(['solid']);
    expect(e.snapshot().format?.font?.underlineColor).toMatch(/^#[0-9A-F]{6}$/);
    expect(e.exec({ type: 'setFontFormat', underlineColor: null }).ok).toBe(true);
    expect(child(child(runs(e)[0]!, 'a:rPr')!, 'a:uFill')).toBeUndefined();
    expect(child(child(runs(e)[0]!, 'a:rPr')!, 'a:uFillTx')).toBeDefined();
    expect(bars()).toEqual(['text']);
  });

  test('a:uFillTx makes the underline follow the text over an inherited a:uFill', async () => {
    const red = '<a:uFill><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:uFill>';
    const e = await open(
      buildPptx({
        slides: [
          {
            shapes: shape({
              id: 2,
              name: 'T',
              x: IN,
              y: IN,
              cx: 6 * IN,
              cy: 2 * IN,
              lstStyle: `<a:lvl1pPr><a:defRPr>${red}</a:defRPr></a:lvl1pPr>`,
              text: para(run('x', 'u="sng"', '<a:uFillTx/>')) + para(run('y', 'u="sng"')),
            }),
          },
        ],
      })
    );
    select(e, 0, 1);
    expect(e.snapshot().format?.font?.underlineColor).toBeNull();
    select(e, 0, 1, 1);
    expect(e.snapshot().format?.font?.underlineColor).toBe('#FF0000');
    // Automatic writes a:uFillTx, so the underline follows the text over the list style.
    expect(e.exec({ type: 'setFontFormat', underlineColor: null }).ok).toBe(true);
    expect(e.snapshot().format?.font?.underlineColor).toBeNull();
  });

  test('refuses values outside the schema or editor limits', async () => {
    const e = await open(deck(para(run('x'))));
    select(e, 0, 1);
    for (const fields of [
      {},
      { fontFamily: '<script>' },
      { fontSizePt: 0 },
      { underline: 'zigzag' },
      { strike: 'triple' },
      { baseline: 150 },
      { caps: 'title' },
      { characterSpacingPt: 5000 },
      { kerningPt: -1 },
      { kerningPt: 0 },
      { kerningPt: 0.004 },
      { color: 'red' },
      { underlineColor: { theme: 'nope' } },
    ]) {
      const r = e.exec({ type: 'setFontFormat', ...fields } as never);
      expect({ fields, ok: r.ok }).toEqual({ fields, ok: false });
    }
    expect(e.snapshot().canUndo).toBe(false);
  });

  test('the values survive save and reopen', async () => {
    const e = await open(deck(para(run('Saved'))));
    select(e, 0, 5);
    e.exec({
      type: 'setFontFormat',
      caps: 'all',
      strike: 'single',
      characterSpacingPt: -0.5,
      underline: 'double',
      underlineColor: '#00B050',
    });
    e.exec({ type: 'toggleSubscript' });
    const again = await open(e.save()!);
    select(again, 0, 5);
    const font = again.snapshot().format?.font;
    expect([
      font?.caps,
      font?.strike,
      font?.characterSpacingPt,
      font?.baseline,
      font?.underlineColor,
    ]).toEqual(['all', 'single', -0.5, -25, '#00B050']);
    expect(toolbarCommandState(again, 'text.dialog').enabled).toBe(true);
  });
});

describe('the Font read over a mixed selection', () => {
  test('a field the runs disagree about reads null, and one value sets them all', async () => {
    const e = await open(deck(para(run('Hello ', 'sz="1800" b="1"') + run('World', 'sz="2400"'))));
    select(e, 0, 11);
    const font = e.snapshot().format!.font!;
    expect([font.fontSizePt, font.bold]).toEqual([null, null]);
    expect(font.disagrees).toMatchObject({ fontSizePt: true, bold: true, italic: false });
    expect(e.exec({ type: 'setFontFormat', fontSizePt: 18, bold: true, italic: true }).ok).toBe(
      true
    );
    const after = e.snapshot().format!.font!;
    expect([after.fontSizePt, after.bold, after.italic]).toEqual([18, true, true]);
  });
});

describe('text.dialog slot and caret words', () => {
  test('runs setFontFormat with fields, never another command, and names its dialog bare', async () => {
    const e = await open(deck(para(run('Hello')) + para(run('two'))));
    select(e, 0, 5);
    expect(runToolbarCommand(e, 'text.dialog', { type: 'delete', bold: true } as never).ok).toBe(
      true
    );
    expect(rPr(runs(e)[0]!, 'b')).toBe('1');
    expect(elements(find(root(e), 'p:txBody')!).filter((c) => c.local === 'p')).toHaveLength(2);
    expect(runToolbarCommand(e, 'text.dialog')).toEqual({
      ok: false,
      reason: 'disabled.dialogSlot',
    });
    expect(runToolbarCommand(e, 'text.characterSpacing', 'dialog' as never)).toEqual({
      ok: false,
      reason: 'disabled.dialogSlot',
    });
  });

  test('an apostrophe stays inside a word; highlight at a caret arms instead', async () => {
    const e = await open(deck(para(run("say don't go"))));
    select(e, 6, 6);
    expect(e.exec({ type: 'toggleBold' }).ok).toBe(true);
    expect(texts(e)).toEqual(['say ', "don't", ' go']);
    const e2 = await open(deck(para(run('hello world'))));
    select(e2, 2, 2);
    expect(e2.exec({ type: 'setHighlight', color: '#FFFF00' }).ok).toBe(true);
    // The word keeps no highlight; the text typed at the caret takes it.
    expect(texts(e2)).toEqual(['hello world']);
    e2.input('text', 'X');
    expect(texts(e2)).toEqual(['he', 'X', 'llo world']);
    expect(child(child(runs(e2)[1]!, 'a:rPr')!, 'a:highlight')).toBeDefined();
  });
});
