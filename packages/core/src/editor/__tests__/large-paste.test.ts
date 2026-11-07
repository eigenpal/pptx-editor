// Pasted text of many lines goes in at a cost that grows with the text, not with its square:
// text with 30,000 newlines (the browser's paste into the text field) and HTML with 5,000
// paragraphs. The text that goes in is the text that typing each line and pressing Enter after
// it gives.

import { expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { canonicalForm } from '../../store/package/serialize.ts';
import { presentationInfo, xmlPartOf } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import type { TextPoint } from '../../contracts/index.ts';
import { paragraphSpecs, richFromHtml } from '../clipboard-rich.ts';

const IN = 914_400;

/**
 * One text box: a bold run, a plain run, and a paragraph mark of its own size; a run in German;
 * and a run with no run properties before a German paragraph mark.
 */
const deck = () =>
  buildPptx({
    slides: [
      {
        shapes: shape({
          id: 2,
          name: 'Body',
          x: IN,
          y: IN,
          cx: 8 * IN,
          cy: 4 * IN,
          bodyPr: '<a:bodyPr wrap="square"/>',
          text:
            para(run('bold ', 'b="1"') + run('plain', 'sz="2000"'), '<a:pPr algn="ctr"/>') +
            para(run('second', 'i="1"')) +
            para(run('dritte', 'lang="de-DE" sz="2400"').replace(' lang="en-US"', '')) +
            '<a:p><a:r><a:t>vierte</a:t></a:r><a:endParaRPr lang="de-DE" b="1"/></a:p>',
        }),
      },
    ],
  });

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false, autoCorrect: false });
  expect((await editor.load(deck())).ok).toBe(true);
  return editor;
}

function select(editor: PptxEditorInstance, anchor: TextPoint, focus = anchor): void {
  const item = editor.session!.layout(0)!.items.find((i) => i.element.name === 'Body')!;
  editor.setSelection({ kind: 'text', nodeId: item.element.nodeId, anchor, focus });
}

/** The paragraphs of the text box. */
function paragraphCount(editor: PptxEditorInstance): number {
  const item = editor.session!.layout(0)!.items[0]!;
  return item.kind === 'shape' ? (item.element.text?.paragraphs.length ?? 0) : 0;
}

/** The slide's shape tree, with no node ids. */
function slideXml(editor: PptxEditorInstance): string {
  const pkg = editor.session!.store.package;
  return canonicalForm(xmlPartOf(pkg, presentationInfo(pkg).slides[0]!.partName)!.root);
}

const CASES: [string, TextPoint, TextPoint?][] = [
  ['the end of a paragraph', { paragraph: 0, offset: 10 }],
  ['inside a run', { paragraph: 0, offset: 7 }],
  ['the start of a paragraph', { paragraph: 1, offset: 0 }],
  ['a range across paragraphs', { paragraph: 0, offset: 2 }, { paragraph: 1, offset: 3 }],
  ['a run in another language', { paragraph: 2, offset: 3 }],
  ['the start of a run in another language', { paragraph: 2, offset: 0 }],
  ['a run with no run properties', { paragraph: 3, offset: 2 }],
  ['the start of a run with no run properties', { paragraph: 3, offset: 0 }],
];

for (const [name, anchor, focus] of CASES)
  test(`lines pasted at ${name} equal each line typed with Enter after it`, async () => {
    const lines = ['one', '', 'two\u000Bbreak', 'three', '\u000Blead', 'four', 'last'];
    const pasted = await open();
    select(pasted, anchor, focus);
    expect(pasted.input('text', lines.join('\n')).ok).toBe(true);
    const typed = await open();
    select(typed, anchor, focus);
    lines.forEach((line, i) => {
      if (i > 0) expect(typed.input('paragraph').ok).toBe(true);
      if (line) expect(typed.input('text', line).ok).toBe(true);
    });
    expect(slideXml(pasted)).toBe(slideXml(typed));
    expect(pasted.snapshot().selection).toEqual(typed.snapshot().selection);
    pasted.destroy();
    typed.destroy();
  });

/** The faster of two pastes of `count` lines, in milliseconds. */
async function linePasteMs(count: number): Promise<number> {
  const text = Array.from({ length: count }, (_, i) => `Line ${i}`).join('\n');
  let best = Infinity;
  for (let k = 0; k < 2; k += 1) {
    const editor = await open();
    select(editor, { paragraph: 0, offset: 10 });
    const t0 = performance.now();
    expect(editor.input('text', text).ok).toBe(true);
    best = Math.min(best, performance.now() - t0);
    expect(paragraphCount(editor)).toBe(count + 3);
    editor.destroy();
  }
  return best;
}

// A cost that grows with the square of the lines takes 100 times as long for 10 times the lines.
// Before the fix, 30,000 lines took about 19 seconds in a browser.
test('30,000 pasted lines cost less than 20 times what 3,000 lines cost', async () => {
  const small = await linePasteMs(3_000);
  const large = await linePasteMs(30_000);
  expect(large).toBeLessThan(small * 20);
});

/** The fastest of five pastes of `count` HTML paragraphs, in milliseconds. */
async function htmlPasteMs(count: number): Promise<number> {
  const html = Array.from({ length: count }, (_, i) => `<p><b>p${i}</b> text</p>`).join('');
  const paragraphs = paragraphSpecs(richFromHtml(html)!);
  let best = Infinity;
  for (let k = 0; k < 5; k += 1) {
    const editor = await open();
    select(editor, { paragraph: 0, offset: 10 });
    const t0 = performance.now();
    expect(editor.exec({ type: 'pasteText', paragraphs }).ok).toBe(true);
    best = Math.min(best, performance.now() - t0);
    expect(paragraphCount(editor)).toBe(count + 3);
    editor.destroy();
  }
  return best;
}

test('pasted HTML paragraphs cost what they insert: four times the paragraphs take less than five times the time', async () => {
  const small = await htmlPasteMs(1_250);
  const large = await htmlPasteMs(5_000);
  expect(large).toBeLessThan(small * 5);
});
