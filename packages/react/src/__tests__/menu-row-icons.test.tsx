import { afterEach, beforeEach, describe, expect, spyOn, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import {
  ICON_PATHS,
  createPptxEditor,
  resetDevWarnings,
  resolveChromeMenuRow,
  type ChromeSlotId,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../context.ts';
import { PptxEditorContent, PptxEditorRoot } from '../PptxEditorRoot.tsx';
import { PptxEditor } from '../index.ts';

// Customization ladder step 2 on menus: the `icon` of `Menu.Item` and `ContextMenu.Slot` shows
// on every row kind, as DOCX Editor's `Menu.Item icon` does (docs/site/content/customization/
// menus.mdx). The Insert > Shape gallery row once dropped it.

const IN = 914_400;
const PX = 9_525;

/** One row of each kind: `[slot, value, direct]`. */
const ROWS: readonly (readonly [ChromeSlotId, string?, boolean?])[] = [
  ['more.duplicate'], // plain command row
  ['text.bold'], // toggle (checkbox) row
  ['paragraph.align', 'center'], // choice (radio) row
  ['shape.effects', 'shadow.none'], // grouped value row
  ['insert.line', 'arrow'], // a line kind, which draws a line by default
  ['arrange.order'], // submenu of options
  ['shape.effects'], // submenu of option groups
  ['insert.line'], // submenu of line kinds
  ['table.insert'], // table size grid
  ['slide.layout'], // layout gallery
  ['slide.add', undefined, false], // layout gallery, from a row that runs at once by default
  ['insert.shape'], // shape gallery
  ['text.color'], // a picker with no menu form: a disabled row
];

const rowKey = ([slot, value]: readonly [ChromeSlotId, string?, boolean?]) =>
  value === undefined ? slot : `${slot}.${value}`;

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

function hostMenu(editor: PptxEditorInstance, icon: (key: string) => React.ReactNode) {
  render(
    <PptxEditorContext.Provider value={editor}>
      <PptxEditor.Menu preset={false}>
        <PptxEditor.Menu.Entry id="host" label="Host">
          {ROWS.map((row) => (
            <PptxEditor.Menu.Item
              key={rowKey(row)}
              slotId={row[0]}
              {...(row[1] !== undefined ? { value: row[1] } : {})}
              {...(row[2] !== undefined ? { direct: row[2] } : {})}
              icon={icon(rowKey(row))}
            />
          ))}
        </PptxEditor.Menu.Entry>
      </PptxEditor.Menu>
    </PptxEditorContext.Provider>
  );
  act(() => void fireEvent.click(screen.getByRole('menuitem', { name: 'Host' })));
  return document.querySelector<HTMLElement>('.pptx-menubar__panel')!;
}

/** The icon cell of each top-level row of a panel. */
const iconCells = (panel: HTMLElement) => [
  ...panel.querySelectorAll<HTMLElement>(':scope > .pptx-menu__row > .pptx-menu__icon'),
];

let warn: ReturnType<typeof spyOn>;
beforeEach(() => {
  resetDevWarnings();
  warn = spyOn(console, 'warn').mockImplementation(() => {});
});
afterEach(() => {
  warn.mockRestore();
  cleanup();
  document.body.replaceChildren();
});

describe('Menu.Item icon on every row kind', () => {
  test('icon content shows in place of the registry icon on each row kind', async () => {
    const editor = await open();
    const panel = hostMenu(editor, (key) => <span data-testid={`icon-${key}`} />);
    const cells = iconCells(panel);
    expect(cells.length).toBe(ROWS.length);
    ROWS.forEach((row, i) => {
      const key = rowKey(row);
      const icon = screen.getByTestId(`icon-${key}`);
      // The host icon is the whole icon cell: no registry icon or drawn line beside it.
      expect(icon.parentElement).toBe(cells[i]!);
      expect(cells[i]!.children.length).toBe(1);
    });
  });

  test('a Material Symbol name draws that symbol on each row kind', async () => {
    const editor = await open();
    const panel = hostMenu(editor, () => 'grid_view');
    const cells = iconCells(panel);
    expect(cells.length).toBe(ROWS.length);
    const want = ICON_PATHS['grid_view']![0];
    for (const cell of cells) {
      const paths = [...cell.querySelectorAll('svg path')].map((p) => p.getAttribute('d'));
      expect(paths[0]).toBe(want);
    }
  });

  test('an unknown symbol name draws an empty icon and warns once, on each row kind', async () => {
    const editor = await open();
    const panel = hostMenu(editor, () => 'not_a_symbol');
    for (const cell of iconCells(panel)) {
      expect(cell.querySelector('svg')).not.toBeNull();
      expect(cell.querySelectorAll('svg path')).toHaveLength(0);
    }
    const said: string[] = warn.mock.calls.map((c: unknown[]) => String(c[0]));
    const unknown = said.filter((m) => m.includes('"not_a_symbol"'));
    expect(unknown).toEqual([
      '[pptx-editor] The icon "not_a_symbol" is not a packaged Material Symbol, so it draws ' +
        'nothing. Use a name from ICON_PATHS, or pass icon content (an element).',
    ]);
  });

  test('icon={null} removes the icon on each row kind', async () => {
    const editor = await open();
    const panel = hostMenu(editor, () => null);
    for (const cell of iconCells(panel)) expect(cell.childElementCount).toBe(0);
  });

  test('The Insert > Shape gallery row keeps its icon and still opens the gallery', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Menu preset={false}>
          <PptxEditor.Menu.Entry id="host" label="Host">
            <PptxEditor.Menu.Item slotId="insert.shape" icon="shapes" />
          </PptxEditor.Menu.Entry>
        </PptxEditor.Menu>
      </PptxEditorContext.Provider>
    );
    act(() => void fireEvent.click(screen.getByRole('menuitem', { name: 'Host' })));
    const panel = document.querySelector<HTMLElement>('.pptx-menubar__panel')!;
    const row = panel.querySelector<HTMLElement>(':scope > .pptx-menu__row')!;
    expect(row.querySelector('.pptx-menu__icon svg path')!.getAttribute('d')).toBe(
      ICON_PATHS['shapes']![0]!
    );
    act(() => void fireEvent.click(row));
    expect(row.getAttribute('aria-expanded')).toBe('true');
    // One cascading row per gallery category, each led by a drawn shape.
    const categories = document.querySelectorAll('.pptx-menu--submenu > .pptx-menu__row');
    expect(categories.length).toBeGreaterThan(0);
  });

  test('without an icon, each row kind keeps its registry icon', async () => {
    const editor = await open();
    const panel = hostMenu(editor, () => undefined);
    const cells = iconCells(panel);
    ROWS.forEach(([slot, value], i) => {
      const drawn = [...cells[i]!.querySelectorAll('svg path')].map((p) => p.getAttribute('d'));
      // Insert > Line rows draw a line; every other row draws its registry paths, or nothing.
      if (slot === 'insert.line') {
        expect(cells[i]!.querySelector('svg')).not.toBeNull();
        return;
      }
      const row = resolveChromeMenuRow({ kind: 'item', slot, ...(value ? { value } : {}) })!;
      expect(drawn).toEqual(row.paths ? [...row.paths] : []);
    });
  });
});

describe('ContextMenu.Slot icon', () => {
  async function canvas(contextMenu: Parameters<typeof PptxEditorRoot>[0]['popups']) {
    let editor: PptxEditorInstance | null = null;
    render(
      <PptxEditorRoot onReady={(e) => (editor = e)} popups={contextMenu!}>
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
    return stage;
  }
  const rightClick = (stage: HTMLElement, x: number, y: number) =>
    act(() => {
      fireEvent.contextMenu(stage, { clientX: (x * IN) / PX, clientY: (y * IN) / PX });
    });

  test('a shape row and the New slide row show the host icon', async () => {
    const stage = await canvas({
      contextMenu: (props) => (
        <PptxEditor.ContextMenu {...props}>
          <PptxEditor.ContextMenu.Slot
            slotId="arrange.order"
            value="front"
            icon={<span data-testid="icon-front" />}
          />
          <PptxEditor.ContextMenu.Slot
            slotId="slide.add"
            icon={<span data-testid="icon-new-slide" />}
          />
        </PptxEditor.ContextMenu>
      ),
    });
    rightClick(stage, 1.5, 1.5);
    expect(
      screen.getByTestId('icon-front').closest('[data-slot="arrange.order.front"]')
    ).not.toBeNull();
    act(() => void fireEvent.keyDown(document, { key: 'Escape' }));
    rightClick(stage, 10, 6);
    expect(screen.getByTestId('icon-new-slide').closest('[data-slot="slide.add"]')).not.toBeNull();
  });
});
