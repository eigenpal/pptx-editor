import { afterAll, afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import {
  chromeSlotIds,
  createPptxEditor,
  toolbarContextOf,
  toolbarGroupsFor,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { serializeXmlPart, xmlPartOf } from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape, solid } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContext, useEditorCommand, useEditorState } from '../context.ts';
import { Toolbar, ToolbarSlot } from '../Toolbar.tsx';

// The React adapter (audit Q3): the hooks, the toolbar derived from `CHROME_GROUPS`, the
// color menus (theme colors written as `a:schemeClr`), menu keyboard access, and the
// package's dependency shape.

const IN = 914_400;

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const r = await editor.load(
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 2 * IN,
            cy: IN,
            fill: solid('FFC000'),
            text: para(run('Hi')),
          }),
        },
      ],
    })
  );
  expect(r.ok).toBe(true);
  return editor;
}

const selectShape = (editor: PptxEditorInstance): void => {
  const id = editor.session!.layout(0)!.items[0]!.element.nodeId;
  act(() => editor.setSelection({ kind: 'shapes', nodeIds: [id] }));
};

const slideXml = (e: PptxEditorInstance): string =>
  serializeXmlPart(xmlPartOf(e.session!.store.package, e.session!.info.slides[0]!.partName)!);

// happy-dom has no layout: give the toolbar room so no group moves into the More menu.
const width = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth');
beforeAll(() => {
  Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
    configurable: true,
    get: () => 100_000,
  });
});
afterAll(() => {
  if (width) Object.defineProperty(HTMLElement.prototype, 'clientWidth', width);
});
afterEach(cleanup);

describe('hooks', () => {
  test('useEditorState re-renders when its slice changes, and only then', async () => {
    const editor = await open();
    let renders = 0;
    function Probe() {
      renders += 1;
      const count = useEditorState((s) => s.slides.length);
      return <span data-testid="count">{count}</span>;
    }
    render(
      <PptxEditorContext.Provider value={editor}>
        <Probe />
      </PptxEditorContext.Provider>
    );
    expect(screen.getByTestId('count').textContent).toBe('1');
    const before = renders;
    act(() => void editor.exec({ type: 'addSlide' }));
    expect(screen.getByTestId('count').textContent).toBe('2');
    const after = renders;
    expect(after).toBeGreaterThan(before);
    // A change outside the slice does not render again.
    act(() => editor.setZoom(2));
    expect(renders).toBe(after);
  });

  test('useEditorCommand gives enabled state with a reason, and runs the command', async () => {
    const editor = await open();
    let handle: ReturnType<typeof useEditorCommand> | null = null;
    function Probe() {
      handle = useEditorCommand('shape.fill');
      return null;
    }
    render(
      <PptxEditorContext.Provider value={editor}>
        <Probe />
      </PptxEditorContext.Provider>
    );
    expect(handle!.isEnabled).toBe(false);
    expect(handle!.disabledReason).toBeString();
    selectShape(editor);
    expect(handle!.isEnabled).toBe(true);
    act(() => handle!.execute({ theme: 'accent2' }));
    expect(slideXml(editor)).toContain('<a:schemeClr val="accent2"/>');
  });
});

describe('toolbar', () => {
  test('renders the groups for the selection in order, with registry slots only', async () => {
    const editor = await open();
    selectShape(editor);
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar />
      </PptxEditorContext.Provider>
    );
    const groups = [...container.querySelectorAll<HTMLElement>('[data-group-index]')];
    const bar = toolbarGroupsFor(toolbarContextOf(editor.snapshot()));
    expect(bar.map((g) => g.id)).toContain('shape');
    expect(groups.map((g) => Number(g.dataset.groupIndex))).toEqual(bar.map((_, i) => i));
    const known = new Set<string>(chromeSlotIds());
    const slots = [...container.querySelectorAll<HTMLElement>('[data-slot]')].map(
      (b) => b.dataset.slot!
    );
    expect(slots.length).toBeGreaterThan(10);
    for (const slot of slots) expect(known.has(slot)).toBe(true);
    // Each slot sits in the group that owns it.
    for (const g of groups) {
      const id = bar[Number(g.dataset.groupIndex)]!.id;
      for (const b of g.querySelectorAll<HTMLElement>('[data-slot]'))
        expect(b.dataset.slot!.startsWith(`${id}.`)).toBe(true);
    }
  });

  test('chrome mousedown keeps the focus where it is', async () => {
    const editor = await open();
    selectShape(editor);
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar />
      </PptxEditorContext.Provider>
    );
    const button = container.querySelector<HTMLElement>('[data-slot="text.bold"]')!;
    expect(fireEvent.mouseDown(button)).toBe(false);
  });

  test('the format painter button shows whether the painter is on', async () => {
    const editor = await open();
    // A hand-built bar keeps the button while nothing is selected, so it shows the off state.
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar preset={false}>
          <ToolbarSlot slotId="format.painter" />
        </Toolbar>
      </PptxEditorContext.Provider>
    );
    const button = container.querySelector<HTMLButtonElement>('[data-slot="format.painter"]')!;
    expect(button.getAttribute('aria-pressed')).toBe('false');
    expect(button.getAttribute('aria-disabled')).toBe('true');
    selectShape(editor);
    expect(button.getAttribute('aria-disabled')).toBeNull();
    act(() => void fireEvent.click(button));
    expect(editor.snapshot().formatPainter).toBe('once');
    expect(button.getAttribute('aria-pressed')).toBe('true');
    // The second click of a double-click keeps it on.
    act(() => void fireEvent.click(button));
    expect(editor.snapshot().formatPainter).toBe('sticky');
    expect(button.getAttribute('aria-pressed')).toBe('true');
    act(() => editor.setMode('view'));
    expect(button.getAttribute('aria-pressed')).toBe('false');
  });
});

/** Open the Fill color menu from the keyboard: focus on its trigger, then activate it. */
async function fillMenu(): Promise<{ editor: PptxEditorInstance; trigger: HTMLElement }> {
  const editor = await open();
  render(
    <div className="pptx-editor">
      <PptxEditorContext.Provider value={editor}>
        <Toolbar />
      </PptxEditorContext.Provider>
    </div>
  );
  selectShape(editor);
  const trigger = screen
    .getAllByRole('button', { name: 'Fill color' })
    .find((b) => b.getAttribute('aria-haspopup') === 'menu')!;
  act(() => trigger.focus());
  fireEvent.click(trigger);
  return { editor, trigger };
}

describe('color menus', () => {
  test('show the theme colors with their variants, by name', async () => {
    await fillMenu();
    const menu = screen.getByRole('menu', { name: 'Fill color' });
    expect(menu).toBeTruthy();
    expect(screen.getByRole('group', { name: 'Theme Colors' })).toBeTruthy();
    expect(screen.getByRole('menuitem', { name: 'Accent 1, Lighter 60%' })).toBeTruthy();
    expect(screen.getByRole('menuitem', { name: 'Text 1, Lighter 50%' })).toBeTruthy();
    expect(screen.getByRole('menuitem', { name: 'Background 1, Darker 5%' })).toBeTruthy();
    expect(screen.getByRole('menuitem', { name: 'Dark Red' })).toBeTruthy();
  });

  test('a theme swatch writes a:schemeClr; a standard swatch writes a:srgbClr', async () => {
    const { editor } = await fillMenu();
    fireEvent.click(screen.getByRole('menuitem', { name: 'Accent 1, Lighter 60%' }));
    expect(slideXml(editor)).toContain(
      '<a:schemeClr val="accent1"><a:lumMod val="40000"/><a:lumOff val="60000"/></a:schemeClr>'
    );
    expect(screen.queryByRole('menu') === null).toBe(true);
    // The split button repeats the last pick, as a theme color.
    act(() => void editor.exec({ type: 'setShapeFill', color: '#000000' }));
    fireEvent.click(document.querySelector<HTMLElement>('[data-slot="shape.fill"]')!);
    expect(slideXml(editor)).toContain('<a:lumOff val="60000"/>');
  });
});

describe('menu keyboard access', () => {
  test('a keyboard open focuses the first row; arrows move by swatch and by grid row', async () => {
    await fillMenu();
    const active = (): string | null => document.activeElement?.getAttribute('aria-label') ?? null;
    expect(active()).toBe('Background 1');
    fireEvent.keyDown(document.activeElement!, { key: 'ArrowRight' });
    expect(active()).toBe('Text 1');
    fireEvent.keyDown(document.activeElement!, { key: 'ArrowDown' });
    expect(active()).toBe('Text 1, Lighter 50%');
    fireEvent.keyDown(document.activeElement!, { key: 'ArrowUp' });
    fireEvent.keyDown(document.activeElement!, { key: 'ArrowLeft' });
    expect(active()).toBe('Background 1');
    // Home and End go to the first and last rows.
    fireEvent.keyDown(document.activeElement!, { key: 'End' });
    expect(document.activeElement?.textContent).toContain('Custom gradient');
    fireEvent.keyDown(document.activeElement!, { key: 'Home' });
    expect(active()).toBe('Background 1');
  });

  test('Escape closes the menu and returns focus to its trigger', async () => {
    const { trigger } = await fillMenu();
    expect(screen.getByRole('menu')).toBeTruthy();
    fireEvent.keyDown(document.activeElement!, { key: 'Escape' });
    expect(screen.queryByRole('menu') === null).toBe(true);
    expect(document.activeElement).toBe(trigger);
  });

  test('picking a row returns focus to the trigger, not to the page', async () => {
    const { trigger } = await fillMenu();
    fireEvent.click(screen.getByRole('menuitem', { name: 'Dark Red' }));
    expect(document.activeElement).toBe(trigger);
  });

  test('a pointer open moves the keys into the menu; Escape returns to the trigger', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <input aria-label="elsewhere" />
        <Toolbar />
      </PptxEditorContext.Provider>
    );
    selectShape(editor);
    const field = screen.getByRole('textbox', { name: 'elsewhere' });
    const trigger = screen
      .getAllByRole('button', { name: 'Fill color' })
      .find((b) => b.getAttribute('aria-haspopup') === 'menu')!;
    const pointerOpen = (): HTMLElement => {
      act(() => field.focus());
      fireEvent.mouseDown(trigger);
      fireEvent.click(trigger);
      return screen.getByRole('menu');
    };
    // The panel takes focus, so a key no longer reaches the field; it shows no ring until a key.
    let menu = pointerOpen();
    expect(document.activeElement).toBe(menu);
    expect(menu.hasAttribute('data-pointer-open')).toBe(true);
    fireEvent.keyDown(menu, { key: 'ArrowDown' });
    expect(menu.hasAttribute('data-pointer-open')).toBe(false);
    expect(document.activeElement?.getAttribute('aria-label')).toBe('Background 1');
    fireEvent.keyDown(document.activeElement!, { key: 'Escape' });
    expect(screen.queryByRole('menu') === null).toBe(true);
    expect(document.activeElement).toBe(trigger);
    // A pick gives focus back to where it was before the pointer open.
    menu = pointerOpen();
    fireEvent.click(screen.getByRole('menuitem', { name: 'Dark Red' }));
    expect(document.activeElement).toBe(field);
  });

  test('action rows are menuitem; choice rows are menuitemradio with aria-checked', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar />
      </PptxEditorContext.Provider>
    );
    selectShape(editor);
    fireEvent.click(screen.getByRole('button', { name: 'Vertical align' }));
    const rows = screen.getAllByRole('menuitemradio');
    expect(rows.length).toBe(3);
    expect(rows.every((r) => r.hasAttribute('aria-checked'))).toBe(true);
    expect(rows.every((r) => r.tabIndex === -1)).toBe(true);
  });
});

describe('package shape', () => {
  const root = join(import.meta.dir, '..', '..');
  const pkg = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8')) as {
    dependencies?: Record<string, string>;
    peerDependencies: Record<string, string>;
  };

  test('the engine and React are peers, so one engine copy resolves; i18n is a dependency', () => {
    // As in DOCX Editor: the adapter installs the shared strings, and the app installs core.
    expect(Object.keys(pkg.dependencies ?? {})).toEqual(['@pptx-editor.dev/i18n']);
    expect(Object.keys(pkg.peerDependencies).sort()).toEqual([
      '@pptx-editor.dev/core',
      'react',
      'react-dom',
    ]);
  });

  test('sources import only the peers, i18n, and their own files', () => {
    const src = join(root, 'src');
    const allowed = /^(react|react-dom|react\/.*|@pptx-editor\.dev\/(core|i18n)(\/.*)?|\.\.?\/.*)$/;
    for (const file of readdirSync(src).filter((f) => /\.tsx?$/.test(f))) {
      const text = readFileSync(join(src, file), 'utf8');
      for (const m of text.matchAll(/^(?:import|export)[^'"]*from '([^']+)'/gm))
        expect([file, allowed.test(m[1]!)]).toEqual([file, true]);
    }
  });
});
