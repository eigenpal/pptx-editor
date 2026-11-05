// Shared helpers for the toolbar tests: one deck, a React and a Vue toolbar on the same editor,
// and a canonical DOM that both adapters must produce (generated ids and the attribute order
// left out), after DOCX Editor's cross-adapter differential test.

import { expect } from 'bun:test';
import { createElement, type ReactNode } from 'react';
import { render } from '@testing-library/react';
import { createApp, defineComponent, h, nextTick, shallowRef, type App } from 'vue';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { PNG } from 'pngjs';
import { buildPptx, cell, para, run, shape, table } from '../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../../react/src/context.ts';
import { providePptxEditor } from '../src/composables.ts';
import { canonicalDom } from './canonical-dom.ts';

export const IN = 914_400;

/** A one-slide deck with one text box. */
export async function openDeck(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run('Hi')) }),
      },
    ],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

/** A one-slide deck with one picture, selected. */
export async function openPictureDeck(): Promise<PptxEditorInstance> {
  const img = new PNG({ width: 4, height: 3 });
  img.data.fill(200);
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        rels: [['rIdImg', R_IMAGE, '../media/image1.png']],
        shapes:
          '<p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture 1"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
          '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
          `<p:spPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${4 * IN}" cy="${3 * IN}"/></a:xfrm>` +
          '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>',
      },
    ],
    media: [
      { name: 'image1.png', bytes: new Uint8Array(PNG.sync.write(img)), contentType: 'image/png' },
    ],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  selectShape(editor);
  return editor;
}

/** A one-slide deck with a 1 × 2 table, the caret in its second cell. */
export async function openTableDeck(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: table({
          id: 2,
          x: IN,
          y: IN,
          cx: 2 * IN,
          cy: IN,
          columns: [IN, IN],
          tblPr: '<a:tblPr firstRow="1"/>',
          rows: `<a:tr h="${IN}">${cell('A')}${cell('B')}</a:tr>`,
        }),
      },
    ],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  const frame = editor.session!.layout(0)!.items[0]!.element.nodeId;
  editor.setSelection({
    kind: 'text',
    nodeId: frame,
    cell: { row: 0, column: 1 },
    anchor: { paragraph: 0, offset: 0 },
    focus: { paragraph: 0, offset: 0 },
  });
  return editor;
}

/** Select the deck's text box as a shape. */
export function selectShape(editor: PptxEditorInstance): void {
  const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
  editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] });
}

/** Put the caret in the deck's text box: the bar shows its text groups. */
export function selectText(editor: PptxEditorInstance): void {
  const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
  const at = { paragraph: 0, offset: 0 };
  editor.setSelection({ kind: 'text', nodeId, anchor: at, focus: at });
}

/** Render a React node over `editor`, and return its container. */
export function renderReact(editor: PptxEditorInstance, node: ReactNode): HTMLElement {
  return render(createElement(PptxEditorContext.Provider, { value: editor }, node)).container;
}

/** Mount a Vue render function over `editor`. */
export async function mountVue(
  editor: PptxEditorInstance,
  renderFn: () => unknown
): Promise<{ host: HTMLElement; app: App }> {
  const host = document.createElement('div');
  document.body.append(host);
  const Child = defineComponent({ setup: () => renderFn as () => ReturnType<typeof h> });
  const app = createApp(
    defineComponent({
      setup() {
        providePptxEditor(shallowRef(editor));
        return () => h(Child);
      },
    })
  );
  app.mount(host);
  await nextTick();
  return { host, app };
}

/**
 * The canonical markup of a toolbar element itself, with whitespace-only text left out. Two
 * adapters that render the same chrome give the same text.
 */
export const canonical = (el: Element): string => canonicalDom(el, { self: true, trimText: true });

/** Make every element `px` wide, so the More menu fit runs as in a browser. */
export function fixWidths(barWidth: () => number, offset = 9): () => void {
  const client = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth');
  const off = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'offsetWidth');
  Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
    configurable: true,
    get: barWidth,
  });
  Object.defineProperty(HTMLElement.prototype, 'offsetWidth', {
    configurable: true,
    get: () => offset,
  });
  return () => {
    if (client) Object.defineProperty(HTMLElement.prototype, 'clientWidth', client);
    else delete (HTMLElement.prototype as { clientWidth?: number }).clientWidth;
    if (off) Object.defineProperty(HTMLElement.prototype, 'offsetWidth', off);
    else delete (HTMLElement.prototype as { offsetWidth?: number }).offsetWidth;
  };
}
