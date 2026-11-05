import { afterAll, afterEach, beforeAll, describe, expect, spyOn, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import { useState } from 'react';
import {
  createPptxEditor,
  resetDevWarnings,
  ribbonTabs,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../context.ts';
import { PptxEditor, Toolbar, useEditorState } from '../index.ts';

// `PptxEditor.Ribbon`: tabs over the toolbar's groups,
// the Quick Access Toolbar, contextual tabs, and the ladder inside each tab.

const IN = 914_400;

async function open(): Promise<PptxEditorInstance> {
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

let barWidth = 100_000;
const width = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth');
beforeAll(() => {
  Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
    configurable: true,
    get: () => barWidth,
  });
});
afterAll(() => {
  if (width) Object.defineProperty(HTMLElement.prototype, 'clientWidth', width);
});
afterEach(() => {
  cleanup();
  barWidth = 100_000;
});

const tabNames = () => screen.getAllByRole('tab').map((t) => t.textContent);
const panelGroups = () =>
  [...document.querySelectorAll('[role="tabpanel"] .pptx-toolbar > [data-group]')].map((g) =>
    g.getAttribute('data-group')
  );

describe('PptxEditor.Ribbon', () => {
  test('the tabs follow the reference PDF, with the Quick Access Toolbar beside them', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Ribbon />
      </PptxEditorContext.Provider>
    );
    expect(tabNames()).toEqual(['Home', 'Insert', 'Design', 'Slide Show', 'View']);
    expect(screen.getByRole('tab', { name: 'Home' }).getAttribute('aria-selected')).toBe('true');
    expect(panelGroups().slice(0, 4)).toEqual(['edit', 'format', 'slide', 'font']);
    expect(document.querySelector('.pptx-ribbon__qat [data-slot="history.undo"]') !== null).toBe(
      true
    );
    // Home draws the clipboard: Paste first, then Cut and Copy.
    const panel = document.querySelector('[role="tabpanel"]')!;
    expect(panel.querySelector('[data-slot="edit.paste"]') !== null).toBe(true);
    expect(panel.querySelector('[data-slot="edit.cut"]') !== null).toBe(true);
  });

  test('a click and the arrow keys select a tab', async () => {
    const editor = await open();
    const changes: string[] = [];
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Ribbon onTabChange={(t) => changes.push(t)} />
      </PptxEditorContext.Provider>
    );
    act(() => void fireEvent.click(screen.getByRole('tab', { name: 'Design' })));
    expect(panelGroups()).toEqual(['file', 'slide']);
    act(
      () =>
        void fireEvent.keyDown(screen.getByRole('tab', { name: 'Design' }), { key: 'ArrowRight' })
    );
    expect(panelGroups()).toEqual(['view', 'slide']);
    expect(changes).toEqual(['design', 'slideShow']);
  });

  test('Shape Format shows while shapes are selected, and a lost tab gives way to Home', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Ribbon />
      </PptxEditorContext.Provider>
    );
    expect(tabNames()).not.toContain('Shape Format');
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] }));
    const format = screen.getByRole('tab', { name: 'Shape Format' });
    expect(format.hasAttribute('data-contextual')).toBe(true);
    act(() => void fireEvent.click(format));
    expect(panelGroups()).toEqual(['insert', 'shape', 'arrange']);
    act(() => editor.setSelection({ kind: 'none' }));
    expect(screen.getByRole('tab', { name: 'Home' }).getAttribute('aria-selected')).toBe('true');
  });

  test('Ribbon.Tab changes a tab with toolbar children, hides one, and adds a host tab', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Ribbon>
          <PptxEditor.Ribbon.Tab id="home" className="acme-home">
            <Toolbar.Strike hidden />
            <Toolbar.Bold className="acme-bold" />
          </PptxEditor.Ribbon.Tab>
          <PptxEditor.Ribbon.Tab id="view" hidden />
          <PptxEditor.Ribbon.Tab id="legal" label="Legal" after="home">
            <Toolbar.Group id="legal" label="Legal">
              <Toolbar.Action label="Ask legal" />
            </Toolbar.Group>
          </PptxEditor.Ribbon.Tab>
        </PptxEditor.Ribbon>
      </PptxEditorContext.Provider>
    );
    expect(tabNames()).toEqual(['Home', 'Legal', 'Insert', 'Design', 'Slide Show']);
    expect(document.querySelector('.pptx-ribbon__panel.acme-home') !== null).toBe(true);
    expect(document.querySelector('[data-slot="text.strike"]') === null).toBe(true);
    expect(document.querySelector('[data-slot="text.bold"]')!.classList.contains('acme-bold')).toBe(
      true
    );
    act(() => void fireEvent.click(screen.getByRole('tab', { name: 'Legal' })));
    expect(screen.getByRole('button', { name: 'Ask legal' })).toBeTruthy();
  });

  test('a controlled tab follows the host state', async () => {
    const editor = await open();
    function Host() {
      const [tab, setTab] = useState('view');
      return (
        <PptxEditorContext.Provider value={editor}>
          <PptxEditor.Ribbon tab={tab} onTabChange={setTab} />
          <span data-testid="tab">{tab}</span>
        </PptxEditorContext.Provider>
      );
    }
    render(<Host />);
    expect(panelGroups()).toEqual(['view']);
    act(() => void fireEvent.click(screen.getByRole('tab', { name: 'Insert' })));
    expect(screen.getByTestId('tab').textContent).toBe('insert');
    expect(panelGroups()).toEqual(['slide', 'table', 'insert']);
  });

  test('preset={false} shows only the host tabs', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Ribbon preset={false}>
          <PptxEditor.Ribbon.Tab id="mine" label="Mine">
            <Toolbar.Group id="mine" label="Mine">
              <Toolbar.Action label="Go" />
            </Toolbar.Group>
          </PptxEditor.Ribbon.Tab>
        </PptxEditor.Ribbon>
      </PptxEditorContext.Provider>
    );
    expect(tabNames()).toEqual(['Mine']);
    expect(document.querySelector('.pptx-ribbon__qat') === null).toBe(true);
  });

  test('a gone contextual tab gives way to Home for good, tells the host, and keeps the focus', async () => {
    const editor = await open();
    const changes: string[] = [];
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Ribbon onTabChange={(t) => changes.push(t)} />
      </PptxEditorContext.Provider>
    );
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] }));
    const format = screen.getByRole('tab', { name: 'Shape Format' });
    act(() => void fireEvent.click(format));
    act(() => format.focus());
    act(() => editor.setSelection({ kind: 'none' }));
    const home = screen.getByRole('tab', { name: 'Home' });
    expect(home.getAttribute('aria-selected')).toBe('true');
    expect(changes).toEqual(['shapeFormat', 'home']);
    expect(document.activeElement).toBe(home);
    // The shape comes back into the selection: Home stays, as no one chose Shape Format again.
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] }));
    expect(screen.getByRole('tab', { name: 'Home' }).getAttribute('aria-selected')).toBe('true');
  });

  test('a change to a contextual tab while its context is off adds no stray tab', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Ribbon>
          <PptxEditor.Ribbon.Tab id="shapeFormat">
            <Toolbar.Slot slotId="arrange.order" hidden />
          </PptxEditor.Ribbon.Tab>
        </PptxEditor.Ribbon>
      </PptxEditorContext.Provider>
    );
    expect(tabNames()).toEqual(['Home', 'Insert', 'Design', 'Slide Show', 'View']);
  });

  test('preset={false} keeps the title and groups of a built-in id', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Ribbon preset={false}>
          <PptxEditor.Ribbon.Tab id="home" />
        </PptxEditor.Ribbon>
      </PptxEditorContext.Provider>
    );
    expect(tabNames()).toEqual(['Home']);
    expect(panelGroups().slice(0, 2)).toEqual(['edit', 'format']);
  });

  test('misuse warns: an untitled unknown id, an unknown after, after on a built-in tab, a stray child', async () => {
    const editor = await open();
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      render(
        <PptxEditorContext.Provider value={editor}>
          <PptxEditor.Ribbon>
            <PptxEditor.Ribbon.Tab id="Home" />
            <PptxEditor.Ribbon.Tab id="mine" label="Mine" after="nowhere" />
            <PptxEditor.Ribbon.Tab id="view" after="home" />
            <PptxEditor.Ribbon.Tab id="gone" hidden />
            <Toolbar.Bold hidden />
          </PptxEditor.Ribbon>
        </PptxEditorContext.Provider>
      );
      const text = warn.mock.calls.map((c) => String(c[0])).join('\n');
      expect(text).toContain('Ribbon.Tab id="Home" is not a built-in tab');
      expect(text).toContain('after="nowhere", but no tab has that id');
      expect(text).toContain('id="view" is a built-in tab, so its after is ignored');
      expect(text).toContain('id="gone" is hidden, but no tab has that id');
      expect(text).toContain('draws only Ribbon.Tab children');
    } finally {
      warn.mockRestore();
    }
  });

  test('an insert opens the Format tab of the new object; a plain selection does not', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Ribbon />
      </PptxEditorContext.Provider>
    );
    const selected = () =>
      document.querySelector('[role="tab"][aria-selected="true"]')!.textContent;
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] }));
    expect(selected()).toBe('Home');
    act(() => void editor.exec({ type: 'insertTextBox' }));
    expect(selected()).toBe('Shape Format');
    act(() => editor.setSelection({ kind: 'none' }));
    act(() => void editor.exec({ type: 'insertTable', rows: 2, cols: 2 }));
    expect(selected()).toBe('Table Design');
  });

  test('a controlled ribbon hears the insert; undo, duplicate, and a plain selection do not switch', async () => {
    const editor = await open();
    const changes: string[] = [];
    function Host() {
      const [tab, setTab] = useState('home');
      return (
        <PptxEditorContext.Provider value={editor}>
          <PptxEditor.Ribbon
            tab={tab}
            onTabChange={(t) => {
              changes.push(t);
              setTab(t);
            }}
          />
        </PptxEditorContext.Provider>
      );
    }
    render(<Host />);
    act(() => void editor.exec({ type: 'insertTextBox' }));
    expect(changes).toEqual(['shapeFormat']);
    act(() => void fireEvent.click(screen.getByRole('tab', { name: 'Home' })));
    act(() => void editor.exec({ type: 'duplicate' }));
    act(() => void editor.exec({ type: 'undo' }));
    act(() => void editor.exec({ type: 'undo' }));
    expect(changes).toEqual(['shapeFormat', 'home']);
  });

  // Regression: a selector that builds a new array on each call looped ("Maximum update depth
  // exceeded") before `useEditorState` kept the last slice while the snapshot was unchanged.
  test('useEditorState with a selector that builds a new array renders, as ribbon.mdx shows', async () => {
    const editor = await open();
    let renders = 0;
    function HostRibbon() {
      const tabs = useEditorState((s) => ribbonTabs(s));
      renders += 1;
      return <p data-testid="tabs">{tabs.map((t) => t.tab.id).join(',')}</p>;
    }
    render(
      <PptxEditorContext.Provider value={editor}>
        <HostRibbon />
      </PptxEditorContext.Provider>
    );
    expect(screen.getByTestId('tabs').textContent?.includes('home')).toBe(true);
    const first = renders;
    expect(first < 4).toBe(true);
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] }));
    expect(screen.getByTestId('tabs').textContent?.includes('shapeFormat')).toBe(true);
    expect(renders - first < 4).toBe(true);
  });
});
