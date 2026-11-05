import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen, within } from '@testing-library/react';
import type { ReactElement } from 'react';
import {
  createPptxEditor,
  resetDevWarnings,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../context.ts';
import { PptxEditorContent, PptxEditorRoot } from '../PptxEditorRoot.tsx';
import { requestLinkPopup } from '../chrome-actions.tsx';
import {
  PptxEditor,
  Toolbar,
  useContextMenuRequest,
  useContextMenuTarget,
  useEditorCommand,
} from '../index.ts';
import { tooltipText } from './tooltip-text.ts';

// `PptxEditor.Menu` and `PptxEditor.ContextMenu`, as DOCX Editor's menus
// (docs/site/content/customization/menus.mdx), and the popups registry entries for the link
// popup, the find bar, and the canvas menu.

const IN = 914_400;
const PX = 9_525;

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false, zoom: 1 });
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

afterEach(() => {
  cleanup();
  document.body.replaceChildren();
});

const triggers = () =>
  [...document.querySelectorAll('.pptx-menubar__trigger')].map((b) => b.textContent);

function openMenu(name: string): HTMLElement {
  const trigger = screen.getByRole('menuitem', { name });
  act(() => void fireEvent.click(trigger));
  return document.querySelector<HTMLElement>('.pptx-menubar__panel')!;
}

const rows = (panel: HTMLElement) =>
  [...panel.querySelectorAll('[data-slot]')].map((r) => r.getAttribute('data-slot'));

describe('PptxEditor.Menu', () => {
  test('the preset bar follows the default menu order, and rows carry the registry state', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Menu />
      </PptxEditorContext.Provider>
    );
    expect(triggers()).toEqual([
      'File',
      'Edit',
      'View',
      'Insert',
      'Format',
      'Arrange',
      'Slide Show',
    ]);
    const edit = openMenu('Edit');
    expect(rows(edit).slice(0, 2)).toEqual(['history.undo', 'history.redo']);
    const undo = edit.querySelector<HTMLElement>('[data-slot="history.undo"]')!;
    expect(undo.getAttribute('aria-disabled')).toBe('true');
    expect(undo.textContent).toContain('Undo');
  });

  test('a row runs its slot and closes the menu', async () => {
    const editor = await open();
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] }));
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Menu />
      </PptxEditorContext.Provider>
    );
    const edit = openMenu('Edit');
    act(() => void fireEvent.click(edit.querySelector('[data-slot="more.duplicate"]')!));
    expect(editor.session!.layout(0)!.items.length).toBe(2);
    expect(document.querySelector('.pptx-menubar__panel') === null).toBe(true);
  });

  test('the pointer moves the focused row, so one row looks active', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Menu />
      </PptxEditorContext.Provider>
    );
    screen.getByRole('menuitem', { name: 'Insert' }).focus();
    const insert = openMenu('Insert');
    const items = [...insert.querySelectorAll<HTMLElement>('[role^="menuitem"]')];
    expect(document.activeElement === items[0]).toBe(true);
    const label = items[3]!.querySelector('.pptx-menu__label')!;
    // The first move only records the pointer, so a menu that opens under a resting pointer
    // keeps its focus; a real move focuses the row under it.
    act(() => void fireEvent.mouseMove(label, { clientX: 10, clientY: 10 }));
    expect(document.activeElement === items[0]).toBe(true);
    act(() => void fireEvent.mouseMove(label, { clientX: 10, clientY: 10 }));
    expect(document.activeElement === items[0]).toBe(true);
    act(() => void fireEvent.mouseMove(label, { clientX: 14, clientY: 12 }));
    expect(document.activeElement === items[3]).toBe(true);
  });

  test('named menus change rows in place, hide menus, and add host rows', async () => {
    const editor = await open();
    let opened = 0;
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Menu className="acme-bar">
          <PptxEditor.Menu.File>
            <PptxEditor.Menu.Row label="Open…" icon="folder_open" onSelect={() => (opened += 1)} />
          </PptxEditor.Menu.File>
          <PptxEditor.Menu.Edit>
            <PptxEditor.Menu.Item slotId="history.redo" hidden />
            <PptxEditor.Menu.Item slotId="history.undo" className="acme-undo" />
          </PptxEditor.Menu.Edit>
          <PptxEditor.Menu.Arrange hidden />
          <PptxEditor.Menu.Entry id="help" label="Help">
            <PptxEditor.Menu.Group label="Support">
              <PptxEditor.Menu.Row label="Contact" onSelect={() => {}} />
            </PptxEditor.Menu.Group>
          </PptxEditor.Menu.Entry>
        </PptxEditor.Menu>
      </PptxEditorContext.Provider>
    );
    expect(document.querySelector('.pptx-menubar.acme-bar') !== null).toBe(true);
    expect(triggers()).not.toContain('Arrange');
    expect(triggers()[triggers().length - 1]).toBe('Help');
    const edit = openMenu('Edit');
    expect(rows(edit)).not.toContain('history.redo');
    expect(edit.querySelector('[data-slot="history.undo"]')!.classList.contains('acme-undo')).toBe(
      true
    );
    act(() => void fireEvent.keyDown(document, { key: 'Escape' }));
    const file = openMenu('File');
    expect(rows(file)).toEqual(['file.pageSetup']);
    act(() => void fireEvent.click(screen.getByRole('menuitem', { name: 'Open…' })));
    expect(opened).toBe(1);
    const help = openMenu('Help');
    expect(help.querySelector('[role="group"][aria-label="Support"]') !== null).toBe(true);
  });

  test('preset={false} renders only the host menus', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Menu preset={false}>
          <PptxEditor.Menu.Entry id="mine" label="Mine">
            <PptxEditor.Menu.Item slotId="text.bold" />
          </PptxEditor.Menu.Entry>
        </PptxEditor.Menu>
      </PptxEditorContext.Provider>
    );
    expect(triggers()).toEqual(['Mine']);
  });

  test('Menu.Item for slide.layout opens the layout gallery and changes the layout', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Menu preset={false}>
          <PptxEditor.Menu.Entry id="slides" label="Slides">
            <PptxEditor.Menu.Item slotId="slide.layout" />
          </PptxEditor.Menu.Entry>
        </PptxEditor.Menu>
      </PptxEditorContext.Provider>
    );
    const panel = openMenu('Slides');
    const sub = panel.querySelector<HTMLElement>('[aria-haspopup="menu"]')!;
    sub.focus();
    act(() => void fireEvent.keyDown(sub, { key: 'ArrowRight' }));
    const gallery = document.querySelector<HTMLElement>('.pptx-menu--submenu')!;
    // The layout gallery, not the shape gallery: each tile is a layout of the deck.
    const tiles = [...gallery.querySelectorAll<HTMLElement>('.pptx-layout-gallery__item')];
    expect(tiles.length).toBe(editor.snapshot().layouts!.length);
    const before = editor.snapshot().slides[0]!.layoutId;
    const other = tiles.find((t) => t.getAttribute('aria-checked') !== 'true')!;
    act(() => void fireEvent.click(other));
    expect(editor.snapshot().slides[0]!.layoutId).not.toBe(before);
  });

  test('Insert > New slide adds a slide at once; a host row with direct={false} opens the gallery', async () => {
    const editor = await open();
    const view = render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Menu />
      </PptxEditorContext.Provider>
    );
    let panel = openMenu('Insert');
    const row = panel.querySelector<HTMLElement>('[data-slot="slide.add"]')!;
    expect(rows(panel)[0]).toBe('slide.add');
    expect(row.getAttribute('aria-haspopup')).toBeNull();
    act(() => void fireEvent.click(row));
    expect(editor.snapshot().slides.length).toBe(2);
    view.unmount();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Menu>
          <PptxEditor.Menu.Insert>
            <PptxEditor.Menu.Item slotId="slide.add" direct={false} />
          </PptxEditor.Menu.Insert>
        </PptxEditor.Menu>
      </PptxEditorContext.Provider>
    );
    panel = openMenu('Insert');
    // The gallery form replaces the built-in row in place and drops its Ctrl+M: the key adds at
    // once. A submenu row carries no `data-slot`.
    const gallery = panel.querySelector<HTMLElement>('[role="menuitem"]')!;
    expect(gallery.getAttribute('aria-haspopup')).toBe('menu');
    expect(gallery.textContent).toBe('New slide');
    expect(panel.querySelector('[data-slot="slide.add"]') === null).toBe(true);
  });

  test('Slide Show > Hide slide is a check row that hides the current slide', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Menu />
      </PptxEditorContext.Provider>
    );
    const panel = openMenu('Slide Show');
    const row = panel.querySelector<HTMLElement>('[data-slot="slide.hide"]')!;
    expect(row.getAttribute('role')).toBe('menuitemcheckbox');
    act(() => void fireEvent.click(row));
    expect(editor.snapshot().slides[0]!.hidden).toBe(true);
  });

  test('File > Page setup opens the registry dialog', async () => {
    let editor: PptxEditorInstance | null = null;
    render(
      <PptxEditorRoot onReady={(e) => (editor = e)}>
        <PptxEditor.Menu />
      </PptxEditorRoot>
    );
    await act(
      async () =>
        void (await editor!.load(buildPptx({ slides: [{ layout: 'blank', shapes: '' }] })))
    );
    const file = openMenu('File');
    act(() => void fireEvent.click(file.querySelector('[data-slot="file.pageSetup"]')!));
    expect(document.querySelector('[data-pptx-dialog="pageSetup"]') !== null).toBe(true);
  });
});

async function canvas(popups?: Parameters<typeof PptxEditorRoot>[0]['popups']) {
  let editor: PptxEditorInstance | null = null;
  render(
    <PptxEditorRoot onReady={(e) => (editor = e)} {...(popups ? { popups } : {})}>
      <PptxEditorContent />
    </PptxEditorRoot>
  );
  await act(
    async () =>
      void (await editor!.load(
        buildPptx({ slides: [{ shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN }) }] })
      ))
  );
  const stage = document.querySelector('.pptx-surface__stage') as HTMLElement;
  stage.getBoundingClientRect = () =>
    ({
      left: 0,
      top: 0,
      right: 1280,
      bottom: 720,
      width: 1280,
      height: 720,
      x: 0,
      y: 0,
    }) as DOMRect;
  return { editor: editor!, stage };
}

const rightClick = (stage: HTMLElement, x: number, y: number) =>
  act(() => {
    fireEvent.contextMenu(stage, { clientX: (x * IN) / PX, clientY: (y * IN) / PX });
  });

describe('PptxEditor.ContextMenu and popups.contextMenu', () => {
  test('a renderer changes rows in place and adds host rows', async () => {
    let asked = '';
    function AskRow() {
      const target = useContextMenuRequest();
      const element = useContextMenuTarget();
      return (
        <PptxEditor.ContextMenu.Item
          label="Ask legal"
          onSelect={() => (asked = element === target?.owner ? (target?.kind ?? '') : 'wrong')}
        />
      );
    }
    const { stage } = await canvas({
      contextMenu: (props) => (
        <PptxEditor.ContextMenu {...props} className="acme-menu">
          <PptxEditor.ContextMenu.Slot slotId="format.painter" hidden />
          <AskRow />
        </PptxEditor.ContextMenu>
      ),
    });
    rightClick(stage, 1.5, 1.5);
    const menu = document.querySelector<HTMLElement>('.pptx-contextmenu.acme-menu')!;
    expect(menu !== null).toBe(true);
    expect(menu.querySelector('[data-slot="format.painter"]') === null).toBe(true);
    expect(menu.querySelector('[data-slot="arrange.order.front"]') !== null).toBe(true);
    act(() => void fireEvent.click(screen.getByRole('menuitem', { name: 'Ask legal' })));
    expect(asked).toBe('shapes');
    expect(document.querySelector('.pptx-contextmenu') === null).toBe(true);
  });

  test('false shows no menu', async () => {
    const { stage } = await canvas({ contextMenu: false });
    rightClick(stage, 1.5, 1.5);
    expect(document.querySelector('.pptx-contextmenu') === null).toBe(true);
  });

  test('empty slide area offers Paste, Format Background, and New slide', async () => {
    const { stage } = await canvas();
    rightClick(stage, 10, 6);
    const slots = [...document.querySelectorAll('.pptx-contextmenu [data-slot]')].map((r) =>
      r.getAttribute('data-slot')
    );
    expect(slots).toEqual(['edit.paste', 'slide.background', 'slide.add']);
  });

  test('New slide is a registry row: ContextMenu.Slot hides it, and no second row shows', async () => {
    const { stage } = await canvas({
      contextMenu: (props) => (
        <PptxEditor.ContextMenu {...props}>
          <PptxEditor.ContextMenu.Slot slotId="slide.add" hidden />
        </PptxEditor.ContextMenu>
      ),
    });
    rightClick(stage, 10, 6);
    const slots = [...document.querySelectorAll('.pptx-contextmenu [data-slot]')].map((r) =>
      r.getAttribute('data-slot')
    );
    expect(slots).toEqual(['edit.paste', 'slide.background']);
  });
});

describe('popups.hyperlink and popups.find', () => {
  test('a hyperlink renderer replaces the link popup', async () => {
    let editor: PptxEditorInstance | null = null;
    render(
      <PptxEditorRoot
        onReady={(e) => (editor = e)}
        popups={{
          hyperlink: ({ onClose }) => (
            <div data-testid="my-link">
              <button type="button" onClick={onClose}>
                Done
              </button>
            </div>
          ),
        }}
      >
        <Toolbar preset={false}>
          <Toolbar.Link />
        </Toolbar>
      </PptxEditorRoot>
    );
    await act(
      async () =>
        void (await editor!.load(
          buildPptx({
            slides: [
              {
                layout: 'blank',
                shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run('Hi')) }),
              },
            ],
          })
        ))
    );
    const nodeId = editor!.session!.layout(0)!.items[0]!.element.nodeId;
    act(() =>
      editor!.setSelection({
        kind: 'text',
        nodeId,
        anchor: { paragraph: 0, offset: 0 },
        focus: { paragraph: 0, offset: 2 },
      })
    );
    act(() => void fireEvent.click(document.querySelector('[data-slot="text.link"]')!));
    expect(screen.getByTestId('my-link')).toBeTruthy();
    act(() => void fireEvent.click(screen.getByText('Done')));
    expect(screen.queryByTestId('my-link') === null).toBe(true);
  });

  test('hyperlink false disables the link control with a reason', async () => {
    const editor = await open();
    render(
      <PptxEditorRoot popups={{ hyperlink: false }}>
        <PptxEditorContext.Provider value={editor}>
          <Toolbar preset={false}>
            <Toolbar.Link />
          </Toolbar>
        </PptxEditorContext.Provider>
      </PptxEditorRoot>
    );
    const link = document.querySelector<HTMLButtonElement>('[data-slot="text.link"]')!;
    expect(link.getAttribute('aria-disabled')).toBe('true');
    expect(tooltipText(link)).toContain('Links are turned off');
  });

  test('find false turns the find bar and its controls off', async () => {
    const editor = await open();
    const { container } = render(
      <PptxEditorRoot popups={{ find: false }}>
        <PptxEditorContext.Provider value={editor}>
          <div className="pptx-editor">
            <Toolbar preset={false}>
              <Toolbar.Find />
            </Toolbar>
            <PptxEditor.FindBar />
          </div>
        </PptxEditorContext.Provider>
      </PptxEditorRoot>
    );
    expect(container.querySelector('.pptx-find') === null).toBe(true);
    const find = container.querySelector<HTMLButtonElement>('[data-slot="find.find"]')!;
    expect(find.getAttribute('aria-disabled')).toBe('true');
  });
});

describe('Menu.Row href', () => {
  test('a web link opens in a new tab, and a script link is disabled', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Menu preset={false}>
          <PptxEditor.Menu.Entry id="help" label="Help">
            <PptxEditor.Menu.Row label="Docs" href="https://example.com/docs" />
            <PptxEditor.Menu.Row label="Bad" href="javascript:alert(1)" />
          </PptxEditor.Menu.Entry>
        </PptxEditor.Menu>
      </PptxEditorContext.Provider>
    );
    openMenu('Help');
    const docs = screen.getByRole('menuitem', { name: 'Docs' });
    expect(docs.localName).toBe('a');
    expect(docs.getAttribute('href')).toBe('https://example.com/docs');
    expect(docs.getAttribute('rel')).toBe('noopener noreferrer');
    const bad = screen.getByRole('menuitem', { name: 'Bad' });
    expect(bad.localName).toBe('button');
    expect(bad.getAttribute('aria-disabled')).toBe('true');
  });
});

describe('menu review fixes, round 1', () => {
  async function loaded(ui: (onReady: (e: PptxEditorInstance) => void) => ReactElement) {
    let editor: PptxEditorInstance | null = null;
    render(ui((e) => (editor = e)));
    await act(
      async () =>
        void (await editor!.load(
          buildPptx({
            slides: [
              {
                layout: 'blank',
                shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run('Hi')) }),
              },
            ],
          })
        ))
    );
    return editor!;
  }

  test('Insert > Link opens the link popup with no toolbar on the page', async () => {
    const editor = await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady}>
        <PptxEditor.Menu />
      </PptxEditorRoot>
    ));
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    act(() =>
      editor.setSelection({
        kind: 'text',
        nodeId,
        anchor: { paragraph: 0, offset: 0 },
        focus: { paragraph: 0, offset: 2 },
      })
    );
    const insert = openMenu('Insert');
    act(() => void fireEvent.click(insert.querySelector('[data-slot="text.link"]')!));
    expect(document.querySelector('.pptx-link') !== null).toBe(true);
  });

  test('a popups entry set to false disables its menu row with the reason', async () => {
    await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady} popups={{ pageSetup: false }}>
        <PptxEditor.Menu />
      </PptxEditorRoot>
    ));
    const row = openMenu('File').querySelector<HTMLElement>('[data-slot="file.pageSetup"]')!;
    expect(row.getAttribute('aria-disabled')).toBe('true');
    expect(tooltipText(row)).toContain('This dialog is turned off');
  });

  test('useEditorCommand runs a dialog slot through the root', async () => {
    function Opener() {
      const cmd = useEditorCommand('file.pageSetup');
      return (
        <button type="button" data-testid="opener" onClick={() => cmd.execute()}>
          Page setup
        </button>
      );
    }
    await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady}>
        <Opener />
      </PptxEditorRoot>
    ));
    act(() => void fireEvent.click(screen.getByTestId('opener')));
    expect(document.querySelector('[data-pptx-dialog="pageSetup"]') !== null).toBe(true);
  });

  test('the menu bar is one Tab stop, and the arrows and ArrowDown work as in DOCX Editor', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Menu />
      </PptxEditorContext.Provider>
    );
    const all = [...document.querySelectorAll<HTMLElement>('.pptx-menubar__trigger')];
    expect(all.filter((b) => b.tabIndex === 0).length).toBe(1);
    all[0]!.focus();
    act(() => void fireEvent.keyDown(all[0]!, { key: 'ArrowRight' }));
    expect(document.activeElement).toBe(all[1]);
    expect(all[1]!.tabIndex).toBe(0);
    act(() => void fireEvent.keyDown(all[1]!, { key: 'ArrowDown' }));
    const panel = document.querySelector<HTMLElement>('.pptx-menubar__panel')!;
    expect(panel.getAttribute('data-menu')).toBe('edit');
    expect(panel.contains(document.activeElement)).toBe(true);
  });

  test('a click after hover switched the menu keeps it open', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Menu />
      </PptxEditorContext.Provider>
    );
    openMenu('File');
    const edit = screen.getByRole('menuitem', { name: 'Edit' });
    act(() => void fireEvent.mouseEnter(edit));
    act(() => void fireEvent.click(edit));
    expect(document.querySelector('.pptx-menubar__panel')!.getAttribute('data-menu')).toBe('edit');
  });

  test('Insert > Table opens the size grid, and a cell inserts a table of that size', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Menu />
      </PptxEditorContext.Provider>
    );
    const insert = openMenu('Insert');
    const table = within(insert).getByRole('menuitem', { name: 'Table' });
    expect(table.getAttribute('aria-haspopup')).toBe('menu');
    table.focus();
    act(() => void fireEvent.keyDown(table, { key: 'ArrowRight' }));
    const grid = document.querySelector<HTMLElement>('.pptx-menu--submenu [role="grid"]')!;
    expect(grid.getAttribute('data-slot')).toBe('table.insert');
    const cell = within(grid).getByRole('gridcell', { name: '3 × 2 table' });
    act(() => void fireEvent.click(cell));
    const items = editor.session!.layout(0)!.items;
    expect(items).toHaveLength(2);
    expect(editor.snapshot().selection).toMatchObject({
      nodeId: items[1]!.element.nodeId,
      cell: { row: 0, column: 0 },
    });
    // The insert closes the menu.
    expect(document.querySelector('.pptx-menubar__panel') === null).toBe(true);
  });

  test('Escape in a submenu closes only the submenu', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Menu />
      </PptxEditorContext.Provider>
    );
    const insert = openMenu('Insert');
    const sub = insert.querySelector<HTMLElement>('[aria-haspopup="menu"]')!;
    sub.focus();
    act(() => void fireEvent.keyDown(sub, { key: 'ArrowRight' }));
    const subPanel = document.querySelector<HTMLElement>('.pptx-menu--submenu')!;
    const item = subPanel.querySelector<HTMLElement>('button, [role="menuitem"]') ?? subPanel;
    act(() => void fireEvent.keyDown(item, { key: 'Escape' }));
    expect(document.querySelector('.pptx-menu--submenu') === null).toBe(true);
    expect(document.querySelector('.pptx-menubar__panel') !== null).toBe(true);
  });

  test('Edit > Cut and Paste move a shape; a caret in text names why Cut is off', async () => {
    const editor = await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady}>
        <PptxEditor.Menu />
      </PptxEditorRoot>
    ));
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] }));
    const cutRow = openMenu('Edit').querySelector('[data-slot="edit.cut"]')!;
    act(() => void fireEvent.click(cutRow));
    expect(editor.session!.layout(0)!.items.length).toBe(0);
    const pasteRow = openMenu('Edit').querySelector('[data-slot="edit.paste"]')!;
    act(() => void fireEvent.click(pasteRow));
    expect(editor.session!.layout(0)!.items.length).toBe(1);
    const pasted = editor.session!.layout(0)!.items[0]!.element.nodeId;
    act(() =>
      editor.setSelection({
        kind: 'text',
        nodeId: pasted,
        anchor: { paragraph: 0, offset: 1 },
        focus: { paragraph: 0, offset: 1 },
      })
    );
    const cut = openMenu('Edit').querySelector<HTMLElement>('[data-slot="edit.cut"]')!;
    expect(cut.getAttribute('aria-disabled')).toBe('true');
    expect(tooltipText(cut)).toContain('Select some text first');
  });

  test('the rail menu takes a renderer, and false shows none', async () => {
    for (const entry of [() => <div data-testid="my-rail-menu" />, false as const]) {
      await loaded((onReady) => (
        <PptxEditorRoot onReady={onReady} popups={{ railContextMenu: entry }}>
          <PptxEditor.SlideRail />
        </PptxEditorRoot>
      ));
      const item = document.querySelector<HTMLElement>('.pptx-rail__item')!;
      act(() => void fireEvent.contextMenu(item, { clientX: 20, clientY: 20 }));
      expect(document.querySelector('[data-testid="my-rail-menu"]') !== null).toBe(entry !== false);
      expect(document.querySelector('.pptx-menu') === null).toBe(true);
      cleanup();
      document.body.replaceChildren();
    }
  });

  test('a find renderer replaces the bar, and Find is off while no bar is mounted', async () => {
    const editor = await loaded((onReady) => (
      <PptxEditorRoot
        onReady={onReady}
        popups={{ find: ({ replace }) => <div data-testid="my-find">{String(replace)}</div> }}
      >
        <PptxEditor.Menu />
        <PptxEditor.FindBar />
      </PptxEditorRoot>
    ));
    const findMenu = within(openMenu('Edit')).getByRole('menuitem', { name: 'Find' });
    act(() => void fireEvent.click(findMenu));
    const find = document.querySelector<HTMLElement>('[data-slot="find.replace"]')!;
    act(() => void fireEvent.click(find));
    expect(screen.getByTestId('my-find').textContent).toBe('true');
    cleanup();
    document.body.replaceChildren();
    await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady}>
        <PptxEditor.Menu />
      </PptxEditorRoot>
    ));
    const sub = within(openMenu('Edit')).getByRole('menuitem', { name: 'Find' });
    act(() => void fireEvent.click(sub));
    const row = document.querySelector<HTMLElement>('[data-slot="find.find"]')!;
    expect(row.getAttribute('aria-disabled')).toBe('true');
    expect(tooltipText(row)).toContain('Add PptxEditor.FindBar');
    expect(editor).toBeTruthy();
  });

  test('in the slide sorter, Find is off with a reason for the end user', async () => {
    const editor = await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady}>
        <PptxEditor.Menu />
        <PptxEditor.FindBar />
      </PptxEditorRoot>
    ));
    act(() => void editor.exec({ type: 'setPresentationView', view: 'slideSorter' }));
    const sub = within(openMenu('Edit')).getByRole('menuitem', { name: 'Find' });
    act(() => void fireEvent.click(sub));
    const row = document.querySelector<HTMLElement>('[data-slot="find.find"]')!;
    expect(row.getAttribute('aria-disabled')).toBe('true');
    expect(tooltipText(row)).toContain('Open the Normal view to find text');
  });

  test('misused menu parts warn', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      await loaded((onReady) => (
        <>
          <PptxEditor.ContextMenu />
          <PptxEditorRoot onReady={onReady}>
            <PptxEditor.Menu>
              <PptxEditor.Menu.Edit>
                <PptxEditor.Menu.Item slotId="text.bold" hidden />
                <PptxEditor.Menu.Item slotId={'nope.nope' as never} />
                <PptxEditor.Menu.Item slotId="text.color" />
                <PptxEditor.Menu.Item slotId="shape.styles" />
              </PptxEditor.Menu.Edit>
              <PptxEditor.Menu.Entry id="acme" label="Acme" after="nope" />
            </PptxEditor.Menu>
          </PptxEditorRoot>
        </>
      ));
      openMenu('Edit');
      const text = warn.mock.calls.map((c: unknown[]) => String(c[0])).join('\n');
      expect(text).toContain('Menu.Edit > Menu.Item slotId="text.bold" hidden matches no row');
      expect(text).toContain('slotId="nope.nope" is not a registered control');
      expect(text).toContain('slotId="text.color" picks a value');
      expect(text).toContain('Menu.Entry has after="nope"');
      expect(text).toContain('PptxEditor.ContextMenu is outside PptxEditor.Root');
      const color = document.querySelector<HTMLElement>('[data-slot="text.color"]')!;
      expect(color.getAttribute('aria-disabled')).toBe('true');
      // Quick Styles is a gallery, but not the Insert Shape grid: its row is disabled too.
      expect(text).toContain('slotId="shape.styles" picks a value');
      const styles = document.querySelector<HTMLElement>('[data-slot="shape.styles"]')!;
      expect(styles.getAttribute('aria-disabled')).toBe('true');
      expect(styles.getAttribute('aria-haspopup')).toBeNull();
    } finally {
      warn.mockRestore();
    }
  });

  test('after a pointer-opened menu runs a row, focus goes back to the slide', async () => {
    const editor = await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady}>
        <PptxEditor.Menu />
        <PptxEditorContent />
      </PptxEditorRoot>
    ));
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    act(() =>
      editor.setSelection({
        kind: 'text',
        nodeId,
        anchor: { paragraph: 0, offset: 0 },
        focus: { paragraph: 0, offset: 2 },
      })
    );
    const input = document.querySelector<HTMLElement>('.pptx-surface__input')!;
    input.focus();
    const bold = openMenu('Format').querySelector<HTMLElement>('[data-slot="text.bold"]')!;
    act(() => void fireEvent.click(bold));
    expect(document.activeElement).toBe(input);
  });

  test('Escape in a pointer-opened menu focuses its menubar item, not the slide', async () => {
    await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady}>
        <PptxEditor.Menu />
        <PptxEditorContent />
      </PptxEditorRoot>
    ));
    // The slide had focus before the pointer open: a row would give it back, Escape does not.
    const before = document.querySelector<HTMLElement>('.pptx-surface__input')!;
    before.focus();
    expect(document.activeElement).toBe(before);
    openMenu('Format');
    expect(!!document.activeElement?.closest('.pptx-menu')).toBe(true);
    act(() => void fireEvent.keyDown(document.activeElement!, { key: 'Escape' }));
    expect(document.querySelectorAll('.pptx-menubar__panel').length).toBe(0);
    expect(document.activeElement).toBe(screen.getByRole('menuitem', { name: 'Format' }));
  });

  test('menu Copy sends the copy event from the surface text input', async () => {
    const editor = await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady}>
        <PptxEditor.Menu />
        <PptxEditorContent />
      </PptxEditorRoot>
    ));
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] }));
    const original = document.execCommand;
    let from = '';
    document.execCommand = ((command: string) => {
      const target = document.activeElement as HTMLElement;
      from = `${command}:${target.className}`;
      target.dispatchEvent(new Event(command, { bubbles: true }));
      return true;
    }) as typeof document.execCommand;
    try {
      const copy = openMenu('Edit').querySelector<HTMLElement>('[data-slot="edit.copy"]')!;
      act(() => void fireEvent.click(copy));
    } finally {
      document.execCommand = original;
    }
    expect(from).toContain('copy:');
    expect(from).toContain('pptx-surface__input');
  });

  test('Enter in the link field applies once, and focus goes back to the slide text', async () => {
    const editor = await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady}>
        <PptxEditorContent />
      </PptxEditorRoot>
    ));
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    act(() =>
      editor.setSelection({
        kind: 'text',
        nodeId,
        anchor: { paragraph: 0, offset: 0 },
        focus: { paragraph: 0, offset: 2 },
      })
    );
    act(() => requestLinkPopup(editor));
    const field = screen.getByRole('textbox', { name: 'Link' });
    expect(document.activeElement).toBe(field);
    act(() => void fireEvent.change(field, { target: { value: 'https://example.com/' } }));
    let notPrevented = true;
    act(() => void (notPrevented = fireEvent.keyDown(field, { key: 'Enter' })));
    expect(notPrevented).toBe(false);
    expect(document.querySelector('.pptx-link-popup') === null).toBe(true);
    expect(document.activeElement).toBe(document.querySelector('.pptx-surface__input'));
  });

  test('useEditorCommand: off outside a root with the reason; a URL sets the link without the popup', async () => {
    const editor = await open();
    const seen: { enabled?: boolean; reason?: string | null } = {};
    function Probe() {
      const cmd = useEditorCommand('file.pageSetup');
      seen.enabled = cmd.isEnabled;
      seen.reason = cmd.disabledReason;
      return null;
    }
    render(
      <PptxEditorContext.Provider value={editor}>
        <Probe />
      </PptxEditorContext.Provider>
    );
    expect(seen).toEqual({ enabled: false, reason: 'disabled.noRoot' });
    cleanup();
    let link: ReturnType<typeof useEditorCommand> | null = null;
    function Linker() {
      link = useEditorCommand('text.link');
      return null;
    }
    const ed = await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady} popups={{ hyperlink: false }}>
        <Linker />
      </PptxEditorRoot>
    ));
    const nodeId = ed.session!.layout(0)!.items[0]!.element.nodeId;
    act(() =>
      ed.setSelection({
        kind: 'text',
        nodeId,
        anchor: { paragraph: 0, offset: 0 },
        focus: { paragraph: 0, offset: 2 },
      })
    );
    act(() => link!.execute('https://example.com/'));
    expect(link!.value).toBe('https://example.com/');
    act(() => link!.execute(null));
    expect(link!.value ?? '').toBe('');
  });

  test('Escape and Remove in the link popup give focus back to the slide text', async () => {
    const editor = await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady}>
        <PptxEditorContent />
      </PptxEditorRoot>
    ));
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    const selectText = () =>
      act(() =>
        editor.setSelection({
          kind: 'text',
          nodeId,
          anchor: { paragraph: 0, offset: 0 },
          focus: { paragraph: 0, offset: 2 },
        })
      );
    const input = () => document.querySelector('.pptx-surface__input');
    selectText();
    act(() => requestLinkPopup(editor));
    act(
      () => void fireEvent.keyDown(screen.getByRole('textbox', { name: 'Link' }), { key: 'Escape' })
    );
    expect(document.querySelector('.pptx-link-popup') === null).toBe(true);
    expect(document.activeElement).toBe(input());
    act(() => void editor.exec({ type: 'setLink', url: 'https://example.com/' }));
    selectText();
    act(() => requestLinkPopup(editor));
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Remove link' })));
    expect(document.activeElement).toBe(input());
  });

  test('the link popup opens at the selected text, and Cancel closes it with no change', async () => {
    // The Vue twin is packages/vue/test/overlays.test.ts ('the link popup opens at the selected text').
    const editor = await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady}>
        <PptxEditorContent />
      </PptxEditorRoot>
    ));
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    act(() =>
      editor.setSelection({
        kind: 'text',
        nodeId,
        anchor: { paragraph: 0, offset: 0 },
        focus: { paragraph: 0, offset: 2 },
      })
    );
    const marks = [...document.querySelectorAll('.pptx-surface__text-selection')];
    expect(marks.length > 0).toBe(true);
    for (const mark of marks)
      mark.getBoundingClientRect = () =>
        ({ left: 100, top: 200, right: 180, bottom: 220, width: 80, height: 20 }) as DOMRect;
    act(() => requestLinkPopup(editor));
    const popup = document.querySelector<HTMLElement>('.pptx-link-popup')!;
    // Under the selected words, with the popover's 6px gap.
    expect(popup.style.left).toBe('100px');
    expect(popup.style.top).toBe('226px');
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Cancel' })));
    expect(document.querySelector('.pptx-link-popup') === null).toBe(true);
    expect(editor.snapshot().format?.link ?? '').toBe('');
    expect(document.activeElement).toBe(document.querySelector('.pptx-surface__input'));
  });

  test('with no text selected, the link popup opens under the control that asked', async () => {
    const editor = await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady}>
        <PptxEditorContent />
      </PptxEditorRoot>
    ));
    const button = document.createElement('button');
    document.body.append(button);
    button.getBoundingClientRect = () =>
      ({ left: 40, top: 10, right: 70, bottom: 30, width: 30, height: 20 }) as DOMRect;
    button.getClientRects = () => [button.getBoundingClientRect()] as unknown as DOMRectList;
    act(() => editor.selectAll());
    act(() => requestLinkPopup(editor, button));
    const popup = document.querySelector<HTMLElement>('.pptx-link-popup')!;
    expect(popup.style.left).toBe('40px');
    expect(popup.style.top).toBe('36px');
    expect(screen.getByRole('button', { name: 'Cancel' })).toBeDefined();
    expect(screen.queryByRole('button', { name: 'Remove link' })).toBeNull();
    button.remove();
  });

  test('execute() for text.link warns when no popup can open', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    let link: ReturnType<typeof useEditorCommand> | null = null;
    function Linker() {
      link = useEditorCommand('text.link');
      return null;
    }
    try {
      const ed = await loaded((onReady) => (
        <PptxEditorRoot onReady={onReady} popups={{ hyperlink: false }}>
          <Linker />
        </PptxEditorRoot>
      ));
      const nodeId = ed.session!.layout(0)!.items[0]!.element.nodeId;
      act(() =>
        ed.setSelection({
          kind: 'text',
          nodeId,
          anchor: { paragraph: 0, offset: 0 },
          focus: { paragraph: 0, offset: 2 },
        })
      );
      act(() => link!.execute());
      expect(warn.mock.calls.map((c: unknown[]) => String(c[0])).join('')).toContain(
        'opens the link popup, but none can open here'
      );
    } finally {
      warn.mockRestore();
    }
  });
});
