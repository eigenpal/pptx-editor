// Formatted text and pictures on the clipboard: HTML from other applications in
// (`clipboard-rich.ts`), the `pasteText` command (`clipboard-text.ts`), the picture of copied
// shapes (`clipboard-picture.ts`), and what the clipboard host reads and writes.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import {
  attr,
  child,
  children,
  elements,
  find,
  textOf,
  xmlPartOf,
  type Element,
} from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import {
  cssDeclarations,
  decodeEntities,
  paragraphSpecs,
  richFromHtml,
  richFromText,
  richHtml,
} from '../clipboard-rich.ts';
import { cleanParagraphs, TEXT_MIME } from '../clipboard-text.ts';
import { pngResolution, withPngResolution } from '../clipboard-picture.ts';
import { createClipboardHost, type ClipboardEditor } from '../clipboard-host.ts';
import { CLIPBOARD_MIME } from '../clipboard.ts';
import { readSlidesClipboard, SLIDES_MIME } from '../slide-clipboard.ts';
import type { EditorCommand } from '../../contracts/index.ts';

const IN = 914_400;

/** HTML as presentation applications write a bulleted and numbered list (abridged). */
const LIST_HTML = `<html><head><meta http-equiv=Content-Type content="text/html; charset=utf-8"></head>
<body>
<!--StartFragment-->
<div style='line-height:90%;margin-top:10.0pt;margin-left:.25in;text-indent:-.25in;
text-align:left'><span style='font-size:28.0pt'><span style='mso-special-format:bullet;
font-family:"Liberation Sans"'>&#8226;</span></span><span style='font-size:28.0pt;
font-family:Carlito;mso-ascii-font-family:Carlito;mso-ascii-theme-font:minor-latin;color:black;
mso-style-textfill-type:solid;mso-style-textfill-fill-themecolor:text1'>First level bullet
from the master body style</span></div>
<div class=O1 style='margin-left:.75in;text-indent:-.25in'><span style='font-size:24.0pt'><span
style='mso-special-format:bullet;font-family:"Liberation Sans"'>&#8226;</span></span><span
style='font-size:24.0pt;font-family:Carlito;color:black'>Second
level</span></div>
<div style='margin-left:.38in;text-indent:-.38in'><span style='font-size:28.0pt'><span
style='mso-special-format:"numbullet15\\,1";color:#2E7D32;font-family:"Liberation Sans"'>I)</span></span><span
style='font-size:28.0pt;font-family:Carlito;color:black'>Roman
with color</span></div>
<p style='margin-left:0in;text-indent:0in;text-align:center'><span
style='font-size:28.0pt;font-family:Carlito;color:black'>No
bullet</span></p>
<!--EndFragment-->
</body></html>`;

describe('HTML from other applications', () => {
  test('a presentation list keeps its bullets, levels, numbers, sizes, and theme formatting', () => {
    const p = richFromHtml(LIST_HTML)!;
    expect(p.map((x) => x.runs.map((r) => r.text).join(''))).toEqual([
      'First level bullet from the master body style',
      'Second level',
      'Roman with color',
      'No bullet',
    ]);
    expect(p[0]!.bullet).toEqual({ kind: 'char', char: '•', font: 'Liberation Sans' });
    expect(p[0]!.runs[0]).toMatchObject({ size: 28, scheme: 'tx1', fontTheme: 'minor' });
    expect(p[0]!.marL).toBe(IN / 4);
    expect(p[0]!.indent).toBe(-IN / 4);
    expect(p[1]!.level).toBe(1);
    expect(p[1]!.runs[0]).toMatchObject({ size: 24, color: '000000', font: 'Carlito' });
    expect(p[2]!.bullet).toEqual({ kind: 'number', scheme: 'romanUcParenR', startAt: 1 });
    // A paragraph without a bullet mark has none, where the slide's list style gives one.
    expect(p[3]!.bullet).toEqual({ kind: 'none' });
    expect(p[3]!.align).toBe('ctr');
  });

  test('web HTML: inline formatting, lists, breaks, entities, and collapsed spaces', () => {
    const p = richFromHtml(
      '<p>Hello   <b>bold</b> <i>it&amp;alic</i><br>next<sup>2</sup></p>' +
        '<ul><li>one</li><li style="color:rgb(255, 0, 0)">two<ul><li>deep</li></ul></li></ul>' +
        '<ol start="3"><li>three</li></ol><h2>Title</h2>'
    )!;
    expect(p[0]!.runs).toEqual([
      { text: 'Hello ' },
      { text: 'bold', bold: true },
      { text: ' ' },
      { text: 'it&alic', italic: true },
      { text: '', br: true },
      { text: 'next' },
      { text: '2', baseline: 30000 },
    ]);
    expect(p[1]).toMatchObject({ level: 0, bullet: { kind: 'char', char: '•' } });
    expect(p[2]!.runs[0]).toMatchObject({ text: 'two', color: 'FF0000' });
    expect(p[3]).toMatchObject({ level: 1, runs: [{ text: 'deep' }] });
    expect(p[4]).toMatchObject({ bullet: { kind: 'number', scheme: 'arabicPeriod', startAt: 3 } });
    expect(p[5]!.runs[0]).toMatchObject({ text: 'Title', bold: true });
    // Web lists mark only their items: other paragraphs keep the slide's own bullet.
    expect(p[0]!.bullet).toBeUndefined();
  });

  test('scripts, styles, and pictures are not text; a page without text is null', () => {
    expect(richFromHtml('<script>alert(1)</script><style>p{}</style><img src="x.png">')).toBeNull();
    const p = richFromHtml('<svg><text>no</text></svg><p onclick="x()">yes</p>')!;
    expect(p.map((x) => x.runs.map((r) => r.text).join(''))).toEqual(['yes']);
  });

  test('broken and hostile HTML ends, with bounded output', () => {
    expect(richFromHtml('<p>open <b>never closed')!.length).toBe(1);
    expect(richFromHtml('</p></b>text<')!.map((x) => x.runs.map((r) => r.text).join(''))).toEqual([
      'text<',
    ]);
    const deep = '<span>'.repeat(100_000) + 'x';
    expect(richFromHtml(deep)![0]!.runs[0]!.text).toBe('x');
    const many = '<p>x</p>'.repeat(20_000);
    expect(richFromHtml(many)!.length).toBe(5_000);
    expect(decodeEntities('&#0;&#xD800;&#x1F600;&bogus;')).toBe('😀&bogus;');
  });

  test('style declarations split outside quotes', () => {
    const css = cssDeclarations(`mso-special-format:"numbullet3\\,1;x";font-family:'A;B'`);
    expect(css.get('mso-special-format')).toBe('"numbullet3\\,1;x"');
    expect(css.get('font-family')).toBe("'A;B'");
  });

  test('plain text: line feeds start paragraphs, U+000B breaks a line', () => {
    expect(richFromText('a\r\nb\u000Bc')).toEqual([
      { runs: [{ text: 'a' }] },
      { runs: [{ text: 'b' }, { text: '', br: true }, { text: 'c' }] },
    ]);
  });
});

describe('copied text as HTML', () => {
  test('HTML from the editor reads back as the same paragraphs', () => {
    const html = richHtml([
      {
        runs: [{ text: 'A <b>', bold: true, size: 20, color: '1F4E79', font: 'Open Sans' }],
        align: 'ctr',
      },
      { runs: [{ text: 'item' }], level: 1, bullet: { kind: 'char', char: '•' } },
      { runs: [{ text: 'n' }], bullet: { kind: 'number', scheme: 'alphaUcParenR', startAt: 2 } },
    ]);
    const p = richFromHtml(html)!;
    expect(p[0]).toMatchObject({
      align: 'ctr',
      runs: [{ text: 'A <b>', bold: true, size: 20, color: '1F4E79', font: 'Open Sans' }],
    });
    expect(p[1]).toMatchObject({ level: 1, bullet: { kind: 'char', char: '•' } });
    expect(p[2]!.bullet).toEqual({ kind: 'number', scheme: 'alphaUcParenR', startAt: 2 });
  });
});

describe('pasted paragraphs are untrusted', () => {
  test('links, pictures, and relationship ids are left out; a field becomes text', () => {
    const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const el = (local: string, kids: unknown[] = [], attrs: unknown[] = []) => ({
      ns: A,
      local,
      prefix: 'a',
      attrs,
      children: kids,
    });
    const p = cleanParagraphs([
      el('p', [
        el('pPr', [el('buBlip', [el('blip', [], [{ ns: R, local: 'embed', value: 'rId9' }])])]),
        el('r', [
          el('rPr', [el('hlinkClick', [], [{ ns: R, local: 'id', value: 'rId3' }])]),
          el('t', [{ text: 'x' }]),
        ]),
        el('fld', [el('t', [{ text: '7' }])], [{ ns: '', local: 'type', value: 'slidenum' }]),
        { ns: 'urn:other', local: 'thing', children: [] },
      ]),
    ])!;
    const json = JSON.stringify(p);
    expect(json).not.toContain('hlinkClick');
    expect(json).not.toContain('buBlip');
    expect(json).not.toContain('rId');
    expect(json).not.toContain('thing');
    expect(p[0]!.children!.map((c) => ('local' in c ? c.local : ''))).toEqual(['pPr', 'r', 'r']);
    expect(cleanParagraphs([{ ns: A, local: 'sp' }])).toBeNull();
    expect(cleanParagraphs('nope')).toBeNull();
  });
});

function deck(): Uint8Array {
  return buildPptx({
    slides: [
      {
        shapes:
          shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 4 * IN,
            cy: IN,
            text: para(run('Hello world')) + para(run('Second')),
          }) + shape({ id: 3, x: IN, y: 4 * IN, cx: IN, cy: IN }),
      },
    ],
  });
}

async function open(): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(deck())).ok).toBe(true);
  return e;
}

const tree = (e: PptxEditorInstance): Element =>
  find(xmlPartOf(e.session!.store.package, e.session!.info.slides[0]!.partName)!.root, 'p:spTree')!;
const shapes = (e: PptxEditorInstance): Element[] =>
  elements(tree(e)).filter((c) => c.local === 'sp');
const texts = (sp: Element): string[] =>
  children(child(sp, 'p:txBody'), 'a:p').map((p) =>
    children(p, 'a:r')
      .map((r) => textOf(child(r, 'a:t')))
      .join('')
  );
const caret = (
  e: PptxEditorInstance,
  paragraph: number,
  offset: number,
  to = { paragraph, offset }
) =>
  e.setSelection({
    kind: 'text',
    nodeId: shapes(e)[0]!.id,
    anchor: { paragraph, offset },
    focus: to,
  });

describe('the pasteText command', () => {
  const pasted = paragraphSpecs([
    { runs: [{ text: 'One', bold: true }] },
    { runs: [{ text: 'Two' }], level: 1 },
    { runs: [{ text: 'Three', color: 'C00000' }] },
  ]);

  test('at a caret: the first paragraph joins the caret, the last one runs into the rest', async () => {
    const e = await open();
    caret(e, 0, 6);
    expect(e.exec({ type: 'pasteText', paragraphs: pasted }).ok).toBe(true);
    expect(texts(shapes(e)[0]!)).toEqual(['Hello One', 'Two', 'Threeworld', 'Second']);
    const sel = e.snapshot().selection;
    expect(sel.kind === 'text' && sel.focus).toEqual({ paragraph: 2, offset: 5 });
    const ps = children(child(shapes(e)[0]!, 'p:txBody'), 'a:p');
    expect(attr(child(children(ps[0]!, 'a:r')[1]!, 'a:rPr'), 'b')).toBe('1');
    expect(attr(child(ps[1]!, 'a:pPr'), 'lvl')).toBe('1');
    // One undo step.
    e.exec({ type: 'undo' });
    expect(texts(shapes(e)[0]!)).toEqual(['Hello world', 'Second']);
  });

  test('over a selection: the selected text goes; one paragraph stays in its paragraph', async () => {
    const e = await open();
    caret(e, 0, 0, { paragraph: 0, offset: 5 });
    const one = paragraphSpecs([{ runs: [{ text: 'Bye' }] }]);
    expect(e.exec({ type: 'pasteText', paragraphs: one }).ok).toBe(true);
    expect(texts(shapes(e)[0]!)).toEqual(['Bye world', 'Second']);
  });

  test('with shapes selected: a new text box holds the text, selected', async () => {
    const e = await open();
    e.setSelection({ kind: 'shapes', nodeIds: [shapes(e)[1]!.id] });
    expect(e.exec({ type: 'pasteText', paragraphs: pasted }).ok).toBe(true);
    const box = shapes(e)[2]!;
    expect(texts(box)).toEqual(['One', 'Two', 'Three']);
    expect(attr(child(find(box, 'p:nvSpPr')!, 'p:cNvSpPr'), 'txBox')).toBe('1');
    expect(e.snapshot().selection).toEqual({ kind: 'shapes', nodeIds: [box.id] });
  });

  test('refuses what is not paragraphs', async () => {
    const e = await open();
    const r = e.exec({ type: 'pasteText', paragraphs: [{ ns: 'x', local: 'p' }] });
    expect(r).toMatchObject({ ok: false, reason: 'errors.clipboard' });
  });
});

describe('picture resolution', () => {
  // A 1 × 1 PNG.
  const PNG = Uint8Array.from(
    atob(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
    ),
    (c) => c.charCodeAt(0)
  );

  test('a PNG states the resolution written to it', () => {
    expect(pngResolution(PNG)).toBeUndefined();
    const out = withPngResolution(PNG, 144);
    expect(pngResolution(out)).toBeCloseTo(144, 0);
    expect(out.length).toBe(PNG.length + 21);
    expect(withPngResolution(new Uint8Array([1, 2, 3]), 144)).toEqual(new Uint8Array([1, 2, 3]));
  });
});

/** A clipboard event's data, as a browser gives it. */
function transfer(data: Record<string, string>, files: File[] = []): DataTransfer {
  return {
    getData: (t: string) => data[t] ?? '',
    setData: (t: string, v: string) => void (data[t] = v),
    get types() {
      return Object.keys(data);
    },
    items: files.map((f) => ({ kind: 'file', type: f.type, getAsFile: () => f })),
    files,
  } as unknown as DataTransfer;
}

/** A clipboard host on an editor, recording the commands it runs. */
async function host(select?: (e: PptxEditorInstance) => void) {
  const e = await open();
  select?.(e);
  const ran: EditorCommand[] = [];
  const editor: ClipboardEditor = {
    session: () => e.session,
    slide: () => 0,
    selection: () => e.snapshot().selection,
    layout: () => e.session!.layout(0),
    readOnly: () => false,
    emit: () => undefined,
    deleteText: () => undefined,
    cutShapes: () => undefined,
    insertText: () => undefined,
    refuse: () => undefined,
    exec: (command) => void ran.push(command),
    selectedSlides: () => e.snapshot().selectedSlides,
  };
  return { e, ran, clip: createClipboardHost(editor) };
}

describe('the clipboard host', () => {
  test('copied text carries its runs, HTML, and plain text, and pastes back formatted', async () => {
    const { e, ran, clip } = await host((e) => caret(e, 0, 0, { paragraph: 1, offset: 3 }));
    const data: Record<string, string> = {};
    expect(clip.copyTo(transfer(data), false)).toBe(true);
    expect(data['text/plain']).toBe('Hello world\nSec');
    expect(data['text/html']).toContain('Hello world');
    const own = JSON.parse(data[TEXT_MIME]!) as { paragraphs: unknown[] };
    expect(own.paragraphs).toHaveLength(2);
    e.setSelection({ kind: 'none' });
    expect(clip.pasteFrom(transfer(data))).toBe(true);
    expect(ran[0]).toMatchObject({ type: 'pasteText' });
  });

  test('HTML text pastes as formatted text; a picture alone pastes as a picture', async () => {
    const { ran, clip } = await host();
    expect(clip.pasteFrom(transfer({ 'text/html': LIST_HTML, 'text/plain': 'x' }))).toBe(true);
    expect(ran[0]).toMatchObject({ type: 'pasteText' });
    const png = new File([new Uint8Array([137, 80, 78, 71])], 'image.png', { type: 'image/png' });
    expect(clip.pasteFrom(transfer({ 'text/html': '<img src="data:,">' }, [png]))).toBe(true);
  });

  test('plain text with nothing selected goes in a text box; in text editing it types', async () => {
    const plain = await host();
    expect(plain.clip.pasteFrom(transfer({ 'text/plain': 'a\nb' }))).toBe(true);
    expect(plain.ran[0]).toMatchObject({ type: 'pasteText' });
    const typing = await host((e) => caret(e, 0, 0));
    expect(typing.clip.pasteFrom(transfer({ 'text/plain': 'a' }))).toBe(false);
  });

  test('copied shapes paste as shapes; text boxes leave as HTML text', async () => {
    const { e, ran, clip } = await host((e) =>
      e.setSelection({ kind: 'shapes', nodeIds: [shapes(e)[1]!.id] })
    );
    const data: Record<string, string> = {};
    expect(clip.copyTo(transfer(data), false)).toBe(true);
    expect(data[CLIPBOARD_MIME]).toBeTruthy();
    e.setSelection({ kind: 'none' });
    expect(clip.pasteFrom(transfer(data))).toBe(true);
    expect(ran[0]).toMatchObject({ type: 'paste' });
  });

  test('slides copied in the slide list paste after the current slide, in another editor too', async () => {
    const { ran, clip } = await host();
    const data: Record<string, string> = {};
    expect(clip.copyTo(transfer(data), false, true)).toBe(true);
    expect(Object.keys(data)).toEqual([SLIDES_MIME]);
    expect(clip.pasteFrom(transfer(data))).toBe(true);
    const command = ran[0]!;
    expect(command).toMatchObject({ type: 'pasteSlides', slides: [0] });
    const other = await open();
    expect(other.exec(command).ok).toBe(true);
    expect(other.session!.info.slides).toHaveLength(2);
    expect(other.snapshot().currentSlide).toBe(1);
    other.exec({ type: 'undo' });
    expect(other.session!.info.slides).toHaveLength(1);
  });

  test('pasteSlides refuses a file that is not a presentation and indexes it does not have', async () => {
    const e = await open();
    const bad = e.exec({ type: 'pasteSlides', file: new Uint8Array([1, 2, 3]), slides: [0] });
    expect(bad).toMatchObject({ ok: false, reason: 'errors.clipboard' });
    const file = e.save()!;
    expect(e.exec({ type: 'pasteSlides', file, slides: [3] })).toMatchObject({ ok: false });
    expect(readSlidesClipboard('{"version":1,"slides":[-1],"file":"AA=="}')).toBeNull();
  });
});
