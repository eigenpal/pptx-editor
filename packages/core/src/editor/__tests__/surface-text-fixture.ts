// The slide surface for the text input tests: an editor with one text shape, attached to a
// DOM, and the events of its input. Each test file closes `win` when it ends.

import { expect } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import { attr, elements, find, xmlPartOf } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

export const IN = 914_400;
export const win = new GlobalWindow();
export const doc = win.document as unknown as Document;

/** An editor with one text shape `T` whose body is `text`, attached to the DOM. */
export async function open(
  text: string,
  options: { autoCorrect?: boolean; blankSlides?: number } = {}
) {
  const g = globalThis as { ResizeObserver?: unknown };
  g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
  const { blankSlides = 0, ...editorOptions } = options;
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false, ...editorOptions });
  const first = { shapes: shape({ id: 2, name: 'T', x: IN, y: IN, cx: 6 * IN, cy: 2 * IN, text }) };
  const blank = Array.from({ length: blankSlides }, () => ({ shapes: '' }));
  const bytes = buildPptx({ slides: [first, ...blank] });
  expect((await e.load(bytes)).ok).toBe(true);
  const host = doc.createElement('div') as unknown as HTMLElement;
  doc.body.append(host as unknown as Node);
  e.attach(host);
  const input = host.querySelector('textarea') as HTMLTextAreaElement;
  const fire = (event: Event): boolean => {
    input.dispatchEvent(event);
    return event.defaultPrevented;
  };
  const beforeInput = (inputType: string, data?: string) =>
    fire(
      new win.InputEvent('beforeinput', {
        inputType,
        data,
        bubbles: true,
        cancelable: true,
      }) as unknown as Event
    );
  /** Type through the input, one `beforeinput` for each character, as a keyboard does. */
  const type = (text: string): void => {
    for (const ch of text) beforeInput('insertText', ch);
  };
  /** The system replaces `[from, to)` of the input's text with `text`. */
  const replace = (from: number, to: number, text: string): boolean => {
    input.setSelectionRange(from, to);
    return beforeInput('insertReplacementText', text);
  };
  return { e, input, type, replace, beforeInput, fire };
}

export const shapeId = (e: PptxEditorInstance): string => {
  const part = e.session!.info.slides[0].partName;
  const tree = find(xmlPartOf(e.session!.store.package, part)!.root, 'p:spTree')!;
  return elements(tree).find((c) => attr(find(c, 'p:cNvPr'), 'name') === 'T')!.id;
};
export const texts = (e: PptxEditorInstance): string[] => {
  const item = e.session!.layout(0)!.items.find((x) => x.kind === 'shape' && x.element.text);
  const paragraphs = item?.kind === 'shape' ? item.element.text!.paragraphs : [];
  return paragraphs.map((p) => p.runs.map((r) => (r.kind === 'break' ? '\n' : r.text)).join(''));
};
export const caret = (e: PptxEditorInstance, paragraph: number, offset: number) =>
  e.setSelection({
    kind: 'text',
    nodeId: shapeId(e),
    anchor: { paragraph, offset },
    focus: { paragraph, offset },
  });
export const focusOf = (e: PptxEditorInstance) => {
  const sel = e.snapshot().selection;
  return sel.kind === 'text' ? sel.focus : null;
};
