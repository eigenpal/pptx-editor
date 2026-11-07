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

/**
 * The fastest of `turns` runs each of `paste(small)` and `paste(large)`, taken in turn, so a
 * busy moment of the machine slows both.
 */
async function fastest(
  paste: (count: number) => Promise<number>,
  small: number,
  large: number,
  turns: number
): Promise<[number, number]> {
  let [a, b] = [Infinity, Infinity];
  for (let k = 0; k < turns; k += 1) {
    a = Math.min(a, await paste(small));
    b = Math.min(b, await paste(large));
  }
  return [a, b];
}

/** The time of one paste of `count` lines, in milliseconds. */
async function linePasteMs(count: number): Promise<number> {
  const text = Array.from({ length: count }, (_, i) => `Line ${i}`).join('\n');
  const editor = await open();
  select(editor, { paragraph: 0, offset: 10 });
  const t0 = performance.now();
  expect(editor.input('text', text).ok).toBe(true);
  const ms = performance.now() - t0;
  expect(paragraphCount(editor)).toBe(count + 3);
  editor.destroy();
  return ms;
}

// A cost that grows with the square of the lines takes 16 times as long for 4 times the lines.
// Before the fix, 30,000 lines took about 19 seconds in a browser.
test('30,000 pasted lines cost less than 10 times what 7,500 lines cost', async () => {
  const [small, large] = await fastest(linePasteMs, 7_500, 30_000, 3);
  expect(large).toBeLessThan(small * 10);
}, 30_000);

/** The time of one paste of `count` HTML paragraphs, in milliseconds. */
async function htmlPasteMs(count: number): Promise<number> {
  const html = Array.from({ length: count }, (_, i) => `<p><b>p${i}</b> text</p>`).join('');
  const paragraphs = paragraphSpecs(richFromHtml(html)!);
  const editor = await open();
  select(editor, { paragraph: 0, offset: 10 });
  const t0 = performance.now();
  expect(editor.exec({ type: 'pasteText', paragraphs }).ok).toBe(true);
  const ms = performance.now() - t0;
  expect(paragraphCount(editor)).toBe(count + 3);
  editor.destroy();
  return ms;
}

// A paste holds at most 5,000 paragraphs. A cost that grows with the square of the paragraphs
// takes 6.25 times as long for 2.5 times the paragraphs. The limit of 5 leaves room for a busy
// machine and still fails that cost.
test('pasted HTML paragraphs cost what they insert: 2.5 times the paragraphs take less than five times the time', async () => {
  const [small, large] = await fastest(htmlPasteMs, 2_000, 5_000, 5);
  expect(large).toBeLessThan(small * 5);
}, 30_000);
