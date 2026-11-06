import { afterAll, afterEach, beforeAll, describe, expect, spyOn, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import {
  createPptxEditor,
  panelItems,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { installDrawingRenderers, type DrawingEffectsRenderer } from '@pptx-editor.dev/core/layout';
import { buildPptx, shape, solid } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../context.ts';
import { PptxEditor, Toolbar, useEditorValueCommand } from '../index.ts';
import { tooltipText } from './tooltip-text.ts';

// G4 Shape Format controls on the flat toolbar: Quick styles, Shape effects, Rotate, and the
// task pane slots, with the customization ladder on the new slots.

const IN = 914_400;

const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          shape({
            id: 2,
            name: 'Blue box',
            x: IN,
            y: IN,
            cx: 2 * IN,
            cy: IN,
            fill: solid('4472C4'),
          }) +
          shape({
            id: 3,
            name: 'Orange box',
            x: 4 * IN,
            y: IN,
            cx: IN,
            cy: IN,
            fill: solid('ED7D31'),
          }),
      },
    ],
  });

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(deck())).ok).toBe(true);
  return editor;
}

const firstId = (e: PptxEditorInstance): string => e.session!.layout(0)!.items[0]!.element.nodeId;

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
afterEach(() => cleanup());

/** An effects renderer that paints nothing: the chrome reads only that one is installed. */
const EFFECTS: DrawingEffectsRenderer = {
  body: () => null,
  glow: () => null,
  reflection: () => null,
};
// The Shape Effects tests below use glow, so they run with the drawing module's renderer
// installed; `Pro effect controls` turns it off. A late install logs a warning once, which these
// tests do not need to see.
let warn: ReturnType<typeof spyOn>;
beforeAll(() => {
  warn = spyOn(console, 'warn').mockImplementation(() => {});
  installDrawingRenderers({ effects: EFFECTS });
});
afterAll(() => {
  installDrawingRenderers({ effects: undefined });
  warn.mockRestore();
});

describe('Shape Format controls', () => {
  test('the new slots compose with preset={false}', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Toolbar preset={false}>
          <Toolbar.Slot slotId="shape.styles" />
          <Toolbar.Slot slotId="shape.effects" />
          <Toolbar.Slot slotId="shape.properties" />
          <Toolbar.Slot slotId="arrange.rotate" />
          <Toolbar.Slot slotId="arrange.selectionPane" />
        </PptxEditor.Toolbar>
      </PptxEditorContext.Provider>
    );
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [firstId(editor)] }));
    for (const name of [
      'Quick styles',
      'Shape effects',
      'Format shape',
      'Rotate',
      'Selection pane',
    ])
      expect(screen.getAllByRole('button', { name })).toHaveLength(1);

    // Shape effects: one submenu per effect, the current choice checked.
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Shape effects' })));
    act(() => void fireEvent.click(screen.getByRole('menuitem', { name: 'Shadow' })));
    const none = screen.getAllByRole('menuitemradio', { name: 'No shadow' })[0]!;
    expect(none.getAttribute('aria-checked')).toBe('true');
    act(
      () =>
        void fireEvent.click(screen.getByRole('menuitemradio', { name: 'Offset: bottom right' }))
    );
    expect(editor.snapshot().geometry?.effects?.shadow).toBe('outerBottomRight');
    expect(screen.queryByRole('menuitemradio', { name: 'Offset: bottom right' }) === null).toBe(
      true
    );
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Shape effects' })));
    act(() => void fireEvent.click(screen.getByRole('menuitem', { name: 'Glow' })));
    // Glow Variations: a row of the six accents per size, Up and Down move by a row.
    const grid = screen.getByRole('group', { name: 'Glow variations' });
    expect(grid.getAttribute('data-columns')).toBe('6');
    act(
      () => void fireEvent.click(screen.getByRole('menuitemradio', { name: '8 pt glow, Accent 2' }))
    );
    expect(editor.snapshot().geometry?.effects).toMatchObject({ glow: 8, glowAccent: 'accent2' });
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Shape effects' })));
    act(() => void fireEvent.click(screen.getByRole('menuitem', { name: 'Glow' })));
    expect(
      screen
        .getByRole('menuitemradio', { name: '8 pt glow, Accent 2' })
        .getAttribute('aria-checked')
    ).toBe('true');
    act(
      () => void fireEvent.click(screen.getByRole('menuitemradio', { name: '8 pt glow, Accent 2' }))
    );

    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Quick styles' })));
    act(
      () =>
        void fireEvent.click(
          screen.getByRole('menuitemradio', { name: 'Intense effect: Accent 2' })
        )
    );
    expect(editor.snapshot().geometry?.shapeStyle).toEqual({
      row: 'intenseEffect',
      color: 'accent2',
    });
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Quick styles' })));
    expect(
      screen
        .getByRole('menuitemradio', { name: 'Intense effect: Accent 2' })
        .getAttribute('aria-checked')
    ).toBe('true');
  });

  test('the default bar leaves them out; the ladder works on them', async () => {
    const editor = await open();
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [firstId(editor)] }));
    const bar = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar />
      </PptxEditorContext.Provider>
    );
    // The bar is full at common widths, so the new shape slots stay out of it.
    for (const slot of ['shape.styles', 'shape.effects', 'shape.properties', 'arrange.rotate'])
      expect(document.querySelector(`[data-slot="${slot}"]`) === null).toBe(true);
    bar.unmount();
    render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar>
          <Toolbar.Group id="acme" label="Acme">
            <Toolbar.Slot slotId="shape.effects" className="acme-effects" />
          </Toolbar.Group>
          <Toolbar.Slot slotId="shape.styles" hidden />
        </Toolbar>
        <Toolbar preset={false}>
          <Toolbar.Slot slotId="arrange.rotate" className="acme-rotate" />
          <Toolbar.Slot slotId="shape.properties" className="acme-properties" />
        </Toolbar>
      </PptxEditorContext.Provider>
    );
    const effects = document.querySelector('[data-slot="shape.effects"]')!;
    expect(effects.classList.contains('acme-effects')).toBe(true);
    expect(document.querySelector('[data-slot="shape.styles"]') === null).toBe(true);
    expect(document.querySelectorAll('.acme-properties').length).toBe(1);
    const rotate = document.querySelectorAll('.acme-rotate');
    expect(rotate.length).toBe(1);
    act(() => void fireEvent.click(rotate[0]!));
    act(() => void fireEvent.click(screen.getByRole('menuitem', { name: 'Flip horizontal' })));
    expect(editor.snapshot().geometry?.flipH).toBe(true);
  });
});

describe('Align menu', () => {
  test('Distribute with two objects shows as disabled with the reason', async () => {
    const editor = await open();
    const ids = editor.session!.layout(0)!.items.map((i) => i.element.nodeId);
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: ids }));
    render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar preset={false}>
          <Toolbar.Slot slotId="arrange.align" />
        </Toolbar>
      </PptxEditorContext.Provider>
    );
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Align objects' })));
    const row = screen.getByRole('menuitem', { name: 'Distribute horizontally' });
    // Disabled but focusable, so the arrow keys reach it and the reason can be read.
    expect(row.getAttribute('aria-disabled')).toBe('true');
    expect((row as HTMLButtonElement).disabled).toBe(false);
    expect(tooltipText(row)).toBe('Select three or more objects to distribute them');
    expect(panelItems(row.closest('[role="menu"]') as HTMLElement)).toContain(row);
    const before = editor.snapshot().geometry;
    act(() => void fireEvent.click(row));
    expect(editor.snapshot().geometry).toEqual(before);
    expect(
      screen.getByRole('menuitem', { name: 'Align left' }).getAttribute('aria-disabled')
    ).toBeNull();
  });
});

describe('task pane slots', () => {
  test('Format shape and Selection pane are composable and say the pane is not built', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar preset={false}>
          <Toolbar.Slot slotId="shape.properties" />
          <Toolbar.Slot slotId="arrange.selectionPane" />
        </Toolbar>
      </PptxEditorContext.Provider>
    );
    for (const name of ['Format shape', 'Selection pane']) {
      const button = screen.getByRole('button', { name });
      // aria-disabled, not disabled: the control stays reachable and gives its reason.
      expect(button.getAttribute('aria-disabled')).toBe('true');
      const reason = document.getElementById(button.getAttribute('aria-describedby')!);
      expect(reason?.textContent).toBe('This pane is not available in this version');
    }
  });

  test('Quick styles and Shape effects stay reachable and give a reason with no selection', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar preset={false}>
          <Toolbar.Slot slotId="shape.styles" />
          <Toolbar.Slot slotId="shape.effects" />
        </Toolbar>
      </PptxEditorContext.Provider>
    );
    for (const name of ['Quick styles', 'Shape effects']) {
      const button = screen.getByRole('button', { name });
      expect(button.hasAttribute('disabled')).toBe(false);
      expect(button.getAttribute('aria-disabled')).toBe('true');
      const reason = document.getElementById(button.getAttribute('aria-describedby')!);
      expect(reason?.textContent?.length).toBeGreaterThan(0);
      // A click on a disabled picker opens nothing.
      act(() => void fireEvent.click(button));
      expect(button.getAttribute('aria-expanded')).toBe('false');
    }
  });
});

describe('useEditorValueCommand', () => {
  test('runs a typed value and shows the slot state', async () => {
    const editor = await open();
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [firstId(editor)] }));
    let handle: ReturnType<typeof useEditorValueCommand<'shape.effects'>> | null = null;
    function Probe() {
      handle = useEditorValueCommand('shape.effects');
      // @ts-expect-error A typo in a typed slot value does not compile.
      if (false as boolean) handle.execute('shadow.outerTpo');
      return null;
    }
    render(
      <PptxEditorContext.Provider value={editor}>
        <Probe />
      </PptxEditorContext.Provider>
    );
    expect(handle!.isEnabled).toBe(true);
    let result: { ok: boolean; reason?: string } | null = null;
    act(() => {
      result = handle!.execute('glow.8.accent2');
    });
    expect(result!.ok).toBe(true);
    expect(editor.snapshot().geometry?.effects).toMatchObject({ glow: 8, glowAccent: 'accent2' });
    expect(handle!.value).toContain('glow.8.accent2');
    expect(handle!.options).toContain('shadow.none');
    // Removing an effect the shape does not have changes nothing: the result says why.
    act(() => {
      result = handle!.execute('reflection.none');
    });
    expect(result!).toMatchObject({ ok: false, reason: 'disabled.noEffectChange' });
  });

  test('a numeric slot gives numbers, as its type says', async () => {
    const editor = await open();
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [firstId(editor)] }));
    let width: ReturnType<typeof useEditorValueCommand<'shape.lineWidth'>> | null = null;
    function Probe() {
      width = useEditorValueCommand('shape.lineWidth');
      return null;
    }
    render(
      <PptxEditorContext.Provider value={editor}>
        <Probe />
      </PptxEditorContext.Provider>
    );
    act(() => void width!.execute(2.25));
    expect(width!.value).toBe(2.25);
    expect(width!.options).toContain(0.75);
    expect(width!.options.every((o) => typeof o === 'number')).toBe(true);
  });
});

describe('Shape effects in a menu', () => {
  test('a menu row opens one submenu for each effect kind, with the choice checked', async () => {
    const editor = await open();
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [firstId(editor)] }));
    act(
      () =>
        void editor.exec({
          type: 'setShapeEffect',
          effect: { kind: 'shadow', preset: 'outerTop' },
        })
    );
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Menu>
          <PptxEditor.Menu.Entry id="acme" label="Acme">
            <PptxEditor.Menu.Item slotId="shape.effects" />
          </PptxEditor.Menu.Entry>
        </PptxEditor.Menu>
      </PptxEditorContext.Provider>
    );
    act(() => void fireEvent.click(screen.getByRole('menuitem', { name: 'Acme' })));
    act(() => void fireEvent.click(screen.getByRole('menuitem', { name: 'Shape effects' })));
    const kinds = ['Shadow', 'Reflection', 'Glow', 'Soft edges'];
    for (const name of kinds) expect(screen.getByRole('menuitem', { name })).toBeTruthy();
    act(() => void fireEvent.click(screen.getByRole('menuitem', { name: 'Shadow' })));
    const checked = [...document.querySelectorAll('[aria-checked="true"]')].map(
      (r) => r.textContent
    );
    expect(checked).toHaveLength(1);
    expect(checked[0]).toContain('Offset: top');
    // One choice of several: radio rows, as the toolbar menu has.
    expect(document.querySelector('[aria-checked="true"]')!.getAttribute('role')).toBe(
      'menuitemradio'
    );
  });
});

const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
  ),
  (c) => c.charCodeAt(0)
);
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

async function openPicture(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        rels: [['rIdImg', IMAGE_REL, '../media/image1.png']],
        shapes:
          '<p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture 1"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
          '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
          `<p:spPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${IN}" cy="${IN}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`,
      },
    ],
    media: [{ name: 'image1.png', bytes: PNG, contentType: 'image/png' }],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

/** The labels of the rows the open menus show. */
const rows = (): string[] =>
  [...document.querySelectorAll('[role="menuitem"], [role="menuitemradio"]')].map(
    (r) => r.textContent?.trim() ?? ''
  );

describe('Pro effect controls', () => {
  afterEach(() => installDrawingRenderers({ effects: EFFECTS }));

  test('without the drawing module, Shape effects offers only the outer shadows', async () => {
    installDrawingRenderers({ effects: undefined });
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Toolbar preset={false}>
          <Toolbar.Slot slotId="shape.effects" />
        </PptxEditor.Toolbar>
      </PptxEditorContext.Provider>
    );
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [firstId(editor)] }));
    const button = screen.getByRole('button', { name: 'Shape effects' });
    act(() => void fireEvent.click(button));
    for (const name of ['Reflection', 'Glow', 'Soft edges'])
      expect(screen.queryByRole('menuitem', { name })).toBeNull();
    act(() => void fireEvent.click(screen.getByRole('menuitem', { name: 'Shadow' })));
    expect(rows()).toContain('Offset: top');
    expect(rows().some((r) => r.startsWith('Inside:'))).toBe(false);
    act(() => void fireEvent.click(screen.getByRole('menuitemradio', { name: 'Offset: top' })));
    expect(editor.snapshot().geometry?.effects?.shadow).toBe('outerTop');

    // A module installed later shows its effects the next time the menu opens.
    installDrawingRenderers({ effects: EFFECTS });
    act(() => void fireEvent.click(button));
    for (const name of ['Shadow', 'Reflection', 'Glow', 'Soft edges'])
      expect(screen.getByRole('menuitem', { name })).toBeTruthy();
  });

  test('without the drawing module, a Shape effects menu row has only the Shadow submenu', async () => {
    installDrawingRenderers({ effects: undefined });
    const editor = await open();
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [firstId(editor)] }));
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Menu>
          <PptxEditor.Menu.Entry id="acme" label="Acme">
            <PptxEditor.Menu.Item slotId="shape.effects" />
          </PptxEditor.Menu.Entry>
        </PptxEditor.Menu>
      </PptxEditorContext.Provider>
    );
    act(() => void fireEvent.click(screen.getByRole('menuitem', { name: 'Acme' })));
    act(() => void fireEvent.click(screen.getByRole('menuitem', { name: 'Shape effects' })));
    expect(screen.getByRole('menuitem', { name: 'Shadow' })).toBeTruthy();
    for (const name of ['Reflection', 'Glow', 'Soft edges'])
      expect(screen.queryByRole('menuitem', { name })).toBeNull();
    act(() => void fireEvent.click(screen.getByRole('menuitem', { name: 'Shadow' })));
    expect(rows().some((r) => r.startsWith('Inside:'))).toBe(false);
  });

  test('Picture styles leaves out the reflection and soft edge styles without the module', async () => {
    const styles = async (): Promise<string[]> => {
      const editor = await openPicture();
      render(
        <PptxEditorContext.Provider value={editor}>
          <PptxEditor.Toolbar preset={false}>
            <Toolbar.Slot slotId="image.styles" />
          </PptxEditor.Toolbar>
        </PptxEditorContext.Provider>
      );
      act(() => editor.setSelection({ kind: 'shapes', nodeIds: [firstId(editor)] }));
      act(() => void fireEvent.click(screen.getByRole('button', { name: 'Picture styles' })));
      const shown = rows();
      cleanup();
      return shown;
    };
    installDrawingRenderers({ effects: undefined });
    expect(await styles()).toEqual(['Simple frame, white', 'Simple frame, black']);
    installDrawingRenderers({ effects: EFFECTS });
    expect(await styles()).toEqual([
      'Simple frame, white',
      'Reflected rounded rectangle',
      'Soft edge rectangle',
      'Simple frame, black',
      'Soft edge oval',
    ]);
  });

  test('useEditorValueCommand options refresh after a late install', async () => {
    installDrawingRenderers({ effects: undefined });
    const editor = await open();
    let options: readonly string[] = [];
    function Probe() {
      options = useEditorValueCommand('shape.effects').options;
      return <output>{options.length}</output>;
    }
    render(
      <PptxEditorContext.Provider value={editor}>
        <Probe />
      </PptxEditorContext.Provider>
    );
    expect(options.some((v) => v.startsWith('glow.'))).toBe(false);
    expect(options.some((v) => v.startsWith('shadow.inner'))).toBe(false);
    act(() => installDrawingRenderers({ effects: EFFECTS }));
    expect(options).toContain('glow.8.accent2');
    expect(options).toContain('shadow.innerTop');
    expect(options).toContain('reflection.none');
    expect(options).toContain('softEdge.none');
  });
});
