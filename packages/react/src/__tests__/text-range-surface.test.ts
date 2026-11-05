import { afterEach, expect, test } from 'bun:test';
import { createPptxEditor, type PptxEditorInstance } from '../../../core/src/editor/index.ts';
import type { EditorSelection } from '../../../core/src/contracts/index.ts';
import { child, children, elements, textOf, xmlPartOf } from '../../../core/src/store/index.ts';
import { buildPptx, cell, para, run, shape, table } from '../../../../scripts/fixtures/builder.ts';

type TextSelection = Extract<EditorSelection, { kind: 'text' }>;
const IN = 914400;
const text = 'A👩‍💻e\u0301Z';
const editors: PptxEditorInstance[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
  document.body.replaceChildren();
});

async function setup(inTable: boolean) {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false, zoom: 1 });
  editors.push(editor);
  const paragraphs = para(run(text)) + para(run('beta'));
  const content = inTable
    ? table({
        id: 2,
        x: IN,
        y: IN,
        cx: 6 * IN,
        cy: 2 * IN,
        columns: [3 * IN, 3 * IN],
        tblPr: '<a:tblPr/>',
        rows: `<a:tr h="${2 * IN}">${cell(text + '\nbeta')}${cell('untouched')}</a:tr>`,
      })
    : shape({ id: 2, x: IN, y: IN, cx: 6 * IN, cy: 2 * IN, text: paragraphs });
  expect(
    (await editor.load(buildPptx({ slides: [{ layout: 'blank', shapes: content }] }))).ok
  ).toBe(true);
  const host = document.createElement('div');
  document.body.append(host);
  editor.attach(host);
  const nodeId = editor.session!.layout(0)!.items[0].element.nodeId;
  const select = (backward: boolean, across: boolean) => {
    const start = { paragraph: 0, offset: 1 };
    const end = across ? { paragraph: 1, offset: 2 } : { paragraph: 0, offset: 8 };
    const selection: TextSelection = {
      kind: 'text',
      nodeId,
      ...(inTable ? { cell: { row: 0, column: 0 } } : {}),
      anchor: backward ? end : start,
      focus: backward ? start : end,
    };
    editor.setSelection(selection);
    return { selection, start, end };
  };
  const key = (key: string, shiftKey = false) => {
    const event = new KeyboardEvent('keydown', { key, shiftKey, bubbles: true, cancelable: true });
    host.querySelector('.pptx-surface')!.dispatchEvent(event);
    expect(event.defaultPrevented).toBe(true);
  };
  const bodyText = () => {
    const part = editor.session!.info.slides[0].partName;
    const root = xmlPartOf(editor.session!.store.package, part)!.root;
    const target = child(child(root, 'p:cSld'), 'p:spTree')!;
    const item = elements(target).find((node) => node.id === nodeId)!;
    if (!inTable) return children(child(item, 'p:txBody'), 'a:p').map(textOf);
    const tbl = child(child(child(item, 'a:graphic'), 'a:graphicData'), 'a:tbl');
    return children(child(child(tbl, 'a:tr'), 'a:tc'), 'a:txBody').flatMap((body) =>
      children(body, 'a:p').map(textOf)
    );
  };
  return { editor, select, key, bodyText };
}

for (const inTable of [false, true]) {
  for (const backward of [false, true]) {
    for (const across of [false, true]) {
      for (const keyName of ['ArrowLeft', 'ArrowRight']) {
        test(`surface collapse: table=${inTable}, backward=${backward}, across=${across}, ${keyName}`, async () => {
          const { editor, select, key, bodyText } = await setup(inTable);
          const { selection, start, end } = select(backward, across);
          const at = keyName === 'ArrowLeft' ? start : end;
          const before = editor.save()!;
          const revision = editor.snapshot().revision;
          key(keyName);
          expect(editor.snapshot().selection).toEqual({ ...selection, anchor: at, focus: at });
          expect(editor.snapshot().revision).toBe(revision);
          expect(editor.save()).toEqual(before);
          expect(editor.input('text', 'X').ok).toBe(true);
          const expected =
            keyName === 'ArrowLeft'
              ? ['AX👩‍💻e\u0301Z', 'beta']
              : across
                ? [text, 'beXta']
                : ['A👩‍💻e\u0301XZ', 'beta'];
          expect(bodyText()).toEqual(expected);
          expect(editor.exec({ type: 'undo' }).ok).toBe(true);
          expect(bodyText()).toEqual([text, 'beta']);
          expect(editor.snapshot().selection).toEqual({ ...selection, anchor: at, focus: at });
          expect(editor.exec({ type: 'redo' }).ok).toBe(true);
          expect(bodyText()).toEqual(expected);
        });
      }
    }
    test(`Shift extends through whole graphemes: table=${inTable}, backward=${backward}`, async () => {
      const { editor, select, key } = await setup(inTable);
      const { selection } = select(backward, false);
      key(backward ? 'ArrowRight' : 'ArrowLeft', true);
      expect(editor.snapshot().selection).toEqual({
        ...selection,
        focus: { paragraph: 0, offset: 6 },
      });
      key(backward ? 'ArrowLeft' : 'ArrowRight', true);
      expect(editor.snapshot().selection).toEqual(selection);
      key('ArrowRight');
      key('ArrowLeft');
      expect(editor.snapshot().selection).toMatchObject({
        anchor: { paragraph: 0, offset: 6 },
        focus: { paragraph: 0, offset: 6 },
      });
    });
  }
}
