// A placeholder without `p:txBody` (the text body is optional in CT_Shape, ECMA-376 Part 1,
// 19.3.1.43): a press in it places the caret, its prompt shows, and the first typed character
// adds the text body, which inherits the body properties of the layout and master
// placeholders. A press without typing leaves the file as it was.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { attr, child, elements, find, textOf, xmlPartOf, type Element } from '../../store/index.ts';
import { readPptx } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { PlaceholderPrompts } from '../placeholder-prompts.ts';
import { placeCaret } from '../paragraph-place.ts';
import { caretParagraph } from '../replacement-text.ts';
import { paragraphSpecs } from '../clipboard-rich.ts';

const ph = (id: number, name: string, type: string, body = '') =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/>` +
  `</p:cNvSpPr><p:nvPr><p:ph type="${type}"/></p:nvPr></p:nvSpPr><p:spPr/>${body}` +
  '<p:extLst><p:ext uri="{00000000-0000-0000-0000-000000000001}"/></p:extLst></p:sp>';

const IN = 914_400;
const picture = (id: number) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Picture"/><p:cNvSpPr/><p:nvPr>` +
  `<p:ph type="pic" idx="${id}"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${IN}" ` +
  `y="${3 * IN}"/><a:ext cx="${3 * IN}" cy="${2 * IN}"/></a:xfrm></p:spPr></p:sp>`;

/** Slide 1: a title without a text body and a picture placeholder. Slide 2: an empty title. */
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      { shapes: ph(2, 'Title', 'title') + picture(3) },
      { shapes: ph(2, 'Title', 'title', '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>') },
    ],
  });

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

async function attached() {
  const g = globalThis as { ResizeObserver?: unknown };
  g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false, zoom: 1 });
  expect((await e.load(deck())).ok).toBe(true);
  const host = doc.createElement('div') as unknown as HTMLElement;
  doc.body.append(host as unknown as Node);
  e.attach(host);
  const stage = host.querySelector('.pptx-surface__stage') as HTMLElement;
  stage.getBoundingClientRect = () =>
    ({ left: 0, top: 0, right: 1280, bottom: 720, width: 1280, height: 720 }) as DOMRect;
  const width = e.session!.layout(0)!.slide.width;
  const px = (emu: number) => (emu / width) * 1280;
  const mouse = (type: string, x: number, y: number) =>
    new win.MouseEvent(type, {
      clientX: px(x),
      clientY: px(y),
      button: 0,
      bubbles: true,
      cancelable: true,
    }) as unknown as Event;
  /** A press and a release at the center of the shape named `name`. */
  const click = (name: string) => {
    const item = e.session!.layout(0)!.items.find((i) => i.element.name === name)!;
    const { cx, cy } = item.element.xfrm;
    const [a, b, c, d, ex, f] = item.matrix;
    const x = a * (cx / 2) + c * (cy / 2) + ex;
    const y = b * (cx / 2) + d * (cy / 2) + f;
    stage.dispatchEvent(mouse('pointerdown', x, y));
    doc.dispatchEvent(new win.MouseEvent('pointerup', { bubbles: true }) as unknown as Event);
  };
  return { e, click };
}

/** The `p:sp` named `name` on slide `index` of the editor's package, or of saved bytes. */
function shapeOf(source: PptxEditorInstance | Uint8Array, name: string, index = 0): Element {
  const pkg =
    source instanceof Uint8Array
      ? (() => {
          const r = readPptx(source);
          if (!r.ok) throw new Error('the saved file does not open');
          return r.pkg;
        })()
      : source.session!.store.package;
  const part = `/ppt/slides/slide${index + 1}.xml`;
  const tree = find(xmlPartOf(pkg, part)!.root, 'p:spTree')!;
  return elements(tree).find((c) => attr(find(c, 'p:cNvPr'), 'name') === name)!;
}

const caret = { paragraph: 0, offset: 0 };

describe('a placeholder without a text body', () => {
  test('a press places the caret, and the prompt shows', async () => {
    const { e, click } = await attached();
    const title = shapeOf(e, 'Title');
    expect(child(title, 'p:txBody')).toBeUndefined();
    const prompts = new PlaceholderPrompts();
    const painted = prompts.paint(e.session, e.session!.layout(0), { kind: 'none' }, false);
    const titleItem = painted!.items.find((i) => i.element.name === 'Title')!;
    const text = titleItem.kind === 'shape' ? titleItem.text?.lines[0]?.runs : undefined;
    expect(text?.map((r) => (r.kind === 'text' ? r.text : '')).join('')).toBe('Click to add title');
    click('Title');
    expect(e.snapshot().selection).toEqual({
      kind: 'text',
      nodeId: title.id,
      anchor: caret,
      focus: caret,
    });
    // The press alone writes nothing.
    expect(child(shapeOf(e, 'Title'), 'p:txBody')).toBeUndefined();
    e.destroy();
  });

  test('the first typed character adds the body, which inherits its body properties', async () => {
    const { e, click } = await attached();
    click('Title');
    expect(e.input('text', 'H').ok).toBe(true);
    expect(e.input('text', 'i').ok).toBe(true);
    const saved = e.save()!;
    const title = shapeOf(saved, 'Title');
    const body = child(title, 'p:txBody')!;
    // CT_Shape order: the body comes before `p:extLst`.
    const order = elements(title).map((c) => c.local);
    expect(order).toEqual(['nvSpPr', 'spPr', 'txBody', 'extLst']);
    // Plain body properties: the layout and master placeholders give the rest.
    expect(attr(child(body, 'a:bodyPr')!, 'anchor')).toBeUndefined();
    const read = (index: number) => {
      const el = e.session!.layout(index)!.items.find((i) => i.element.name === 'Title')!;
      return el.kind === 'shape' ? el.element.text : undefined;
    };
    expect(
      read(0)!
        .paragraphs[0]!.runs.map((r) => r.text)
        .join('')
    ).toBe('Hi');
    expect(read(0)!.body).toEqual(read(1)!.body);
    expect(read(0)!.body.anchor).toBe('ctr');
    // One undo removes the text and the body that it added.
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(child(shapeOf(e, 'Title'), 'p:txBody')).toBeUndefined();
    e.destroy();
  });

  test('a delete key at the caret writes nothing', async () => {
    const { e, click } = await attached();
    click('Title');
    expect(e.input('deleteBackward').ok).toBe(true);
    expect(e.input('deleteForward').ok).toBe(true);
    expect(child(shapeOf(e, 'Title'), 'p:txBody')).toBeUndefined();
    expect(e.snapshot().canUndo).toBe(false);
    e.destroy();
  });

  test('a paste at the caret adds the body', async () => {
    const { e, click } = await attached();
    click('Title');
    const paragraphs = paragraphSpecs([{ runs: [{ text: 'Pasted' }] }]);
    expect(e.exec({ type: 'pasteText', paragraphs }).ok).toBe(true);
    const body = child(shapeOf(e, 'Title'), 'p:txBody')!;
    expect(textOf(body)).toBe('Pasted');
    e.destroy();
  });

  test('a composition place holds while the placeholder has no body, and ends when it gets one', async () => {
    const { e, click } = await attached();
    click('Title');
    const part = e.session!.info.slides[0]!.partName;
    const sel = e.snapshot().selection;
    // Every read gives the same paragraph id, so the place of a composition holds.
    const first = caretParagraph(e.session!, part, sel)!;
    expect(caretParagraph(e.session!, part, sel)!.id).toBe(first.id);
    const place = { at: sel, paragraph: first.id, before: '' };
    expect(placeCaret(e.session!.probeBatch(), part, place)).toEqual(sel as never);
    expect(child(shapeOf(e, 'Title'), 'p:txBody')).toBeUndefined();
    // Once an edit adds the body, the place no longer holds.
    expect(e.input('text', 'A').ok).toBe(true);
    expect(placeCaret(e.session!.probeBatch(), part, place)).toBeNull();
    e.destroy();
  });

  test('a picture placeholder without a body still takes a press as a move', async () => {
    const { e, click } = await attached();
    click('Picture');
    expect(e.snapshot().selection).toEqual({
      kind: 'shapes',
      nodeIds: [shapeOf(e, 'Picture').id],
    });
    e.destroy();
  });
});
