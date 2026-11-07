// Clipboard events in the editor's chrome around the slide: the slide list and toolbar take
// them; fields, dialogs, menus, the slide itself, and other editors keep them.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { bindEditorClipboard, isEditorChromeEvent } from '../editor-clipboard.ts';

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
const saved = globalThis.Element;
globalThis.Element = win.Element as unknown as typeof Element;
afterAll(() => {
  globalThis.Element = saved;
  win.happyDOM.close();
});

describe('clipboard events in the editor chrome', () => {
  doc.body.innerHTML = `
    <div class="pptx-editor" id="one">
      <div class="pptx-rail"><div id="thumb" tabindex="0"></div></div>
      <button id="tool"></button>
      <input id="field">
      <div role="dialog"><button id="in-dialog"></button></div>
      <div role="menu"><div id="in-menu" role="menuitem"></div></div>
      <div id="surface"><div id="on-slide"></div></div>
      <div class="pptx-editor" id="nested"><button id="nested-tool"></button></div>
    </div>
    <button id="outside"></button>`;
  const surface = doc.getElementById('surface') as HTMLElement;
  const at = (id: string) => isEditorChromeEvent(surface, doc.getElementById(id));

  test('the slide list and the toolbar are the editor chrome', () => {
    expect(at('thumb')).toBe(true);
    expect(at('tool')).toBe(true);
  });

  test('fields, dialogs, menus, the slide, other editors, and the page are not', () => {
    for (const id of ['field', 'in-dialog', 'in-menu', 'on-slide', 'nested-tool', 'outside'])
      expect(at(id)).toBe(false);
    expect(isEditorChromeEvent(surface, null)).toBe(false);
  });

  test('a paste in the slide list tells the host it comes from there', () => {
    const calls: [string, boolean | undefined][] = [];
    const unbind = bindEditorClipboard(surface, {
      copy: () => false,
      paste: (_data, slideList) => (calls.push(['paste', slideList]), true),
    });
    const fire = (id: string) => {
      const e = new win.Event('paste', { bubbles: true, cancelable: true });
      Object.defineProperty(e, 'clipboardData', { value: {} });
      doc.getElementById(id)!.dispatchEvent(e as unknown as Event);
    };
    fire('thumb');
    fire('tool');
    unbind();
    fire('thumb');
    expect(calls).toEqual([
      ['paste', true],
      ['paste', false],
    ]);
  });
});
