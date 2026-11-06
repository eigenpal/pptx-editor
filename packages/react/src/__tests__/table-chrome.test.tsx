import { afterAll, afterEach, beforeAll, describe, expect, spyOn, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { attr, children, find, xmlPartOf, type Element } from '@pptx-editor.dev/core/store';
import { createTranslator } from '@pptx-editor.dev/i18n';
import { buildPptx, cell, shape, table } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContext, useEditorCommand } from '../context.ts';
import { Toolbar, ToolbarGroup, ToolbarSlot } from '../Toolbar.tsx';
import { MenuTableGrid } from '../TableControls.tsx';
import { Toolbar as Parts } from '../index.ts';
import { tooltipText } from './tooltip-text.ts';

// Tables in the React adapter (editor-ui-parity G6): a click in a cell places the caret,
// a drag across cells selects a range, a grid line drag resizes columns, and the Table Design
// and Table Layout groups join the toolbar while the selection is in a table.

const IN = 914_400;
const PX = 9_525;
const t = createTranslator();
const STYLE = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';

/** `behind` puts a large shape under the table, as a slide's decoration sits; `front`, a box over it. */
const deck = (extra?: 'behind' | 'front'): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          (extra === 'behind'
            ? shape({ id: 3, x: 0, y: 0, cx: 8 * IN, cy: 5 * IN, fill: 'FF0000' })
            : '') +
          table({
            id: 2,
            x: IN,
            y: IN,
            cx: 3 * IN,
            cy: 2 * IN,
            columns: [IN, IN, IN],
            tblPr: `<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>${STYLE}</a:tableStyleId></a:tblPr>`,
            rows: ['ABC', 'DEF']
              .map((r) => `<a:tr h="${IN}">${[...r].map((c) => cell(c)).join('')}</a:tr>`)
              .join(''),
          }) +
          (extra === 'front'
            ? shape({ id: 4, x: 1.5 * IN, y: 1.25 * IN, cx: IN, cy: IN / 2 })
            : ''),
      },
    ],
  });

async function open(extra?: 'behind' | 'front'): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({
    fixedMeasurer: true,
    substitutes: false,
    zoom: 1,
    translate: (key, params) => t(key as never, params),
  });
  expect((await editor.load(deck(extra))).ok).toBe(true);
  return editor;
}

const part = (e: PptxEditorInstance) => e.session!.info.slides[0].partName;
const frames = (e: PptxEditorInstance): Element[] =>
  find(xmlPartOf(e.session!.store.package, part(e))!.root, 'p:spTree')!.children.filter(
    (c): c is Element => c.type === 'element' && c.local === 'graphicFrame'
  );
const rowsOf = (frame: Element) => children(find(frame, 'a:tbl'), 'a:tr');

/** Attach the surface; happy-dom lays nothing out, so one CSS pixel is 9 525 EMU at zoom 1. */
function surface(editor: PptxEditorInstance): HTMLDivElement {
  const host = document.createElement('div');
  document.body.append(host);
  editor.attach(host);
  const stage = host.querySelector('.pptx-surface__stage') as HTMLElement;
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
  return host;
}

const at = (emuX: number, emuY: number) => ({ clientX: emuX / PX, clientY: emuY / PX });
const stageOf = (host: HTMLElement) => host.querySelector('.pptx-surface__stage') as HTMLElement;

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
afterEach(() => {
  cleanup();
  document.body.replaceChildren();
});

describe('table surface', () => {
  test('a click in a cell places the caret there; a drag across cells selects a range', async () => {
    const editor = await open();
    const host = surface(editor);
    const stage = stageOf(host);
    stage.dispatchEvent(new PointerEvent('pointerdown', { ...at(2.5 * IN, 2.5 * IN), button: 0 }));
    expect(editor.snapshot().selection).toMatchObject({
      kind: 'text',
      cell: { row: 1, column: 1 },
    });
    document.dispatchEvent(new PointerEvent('pointermove', at(3.5 * IN, 2.5 * IN)));
    document.dispatchEvent(new PointerEvent('pointerup', at(3.5 * IN, 2.5 * IN)));
    expect(editor.snapshot().selection).toMatchObject({
      kind: 'cells',
      anchor: { row: 1, column: 1 },
      focus: { row: 1, column: 2 },
    });
    expect(host.querySelectorAll('.pptx-surface__cell-selection')).toHaveLength(2);
    expect(editor.snapshot().table).toMatchObject({ selectedColumns: { from: 1, to: 2 } });
    editor.destroy();
  });

  test('a screen reader hears the cell of the caret', async () => {
    const editor = await open();
    const host = surface(editor);
    const id = frames(editor)[0].id;
    const at0 = { paragraph: 0, offset: 0 };
    act(() =>
      editor.setSelection({
        kind: 'text',
        nodeId: id,
        cell: { row: 1, column: 2 },
        anchor: at0,
        focus: at0,
      })
    );
    const live = host.querySelector('[aria-live="polite"]');
    expect(live?.textContent).toBe('Row 2, column 3');
    // The selected table is named, as any selected shape is (Tab moves through them).
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [id] }));
    expect(live?.textContent).toBe('Table 1 selected');
    act(() => editor.setSelection({ kind: 'none' }));
    expect(live?.textContent).toBe('');
    editor.destroy();
  });

  test('a handle drag scales the table from its start, whatever path the pointer took', async () => {
    const editor = await open();
    const host = surface(editor);
    const id = frames(editor)[0].id;
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [id] }));
    const sizes = () => {
      const tbl = find(frames(editor)[0], 'a:tbl')!;
      return {
        w: children(find(tbl, 'a:tblGrid'), 'a:gridCol').map((g) => attr(g, 'w')),
        h: rowsOf(frames(editor)[0]).map((r) => attr(r, 'h')),
      };
    };
    const before = sizes();
    // The right and bottom edges are grid lines, so a table has no east or south side handle.
    expect(host.querySelector('[data-handle="s"], [data-handle="e"]')).toBeNull();
    const handle = host.querySelector('[data-handle="n"]') as SVGCircleElement;
    const hx = Number(handle.getAttribute('x')) + 4 * PX;
    const hy = Number(handle.getAttribute('y')) + 4 * PX;
    const p = (dy: number) => ({ clientX: hx / PX, clientY: (hy - dy) / PX, bubbles: true });
    handle.dispatchEvent(new PointerEvent('pointerdown', { ...p(0), button: 0 }));
    for (const dy of [IN, 2 * IN, -IN / 2, -IN, IN / 2])
      document.dispatchEvent(new PointerEvent('pointermove', p(dy)));
    document.dispatchEvent(new PointerEvent('pointerup', p(IN / 2)));
    // The box ends half an inch taller than at the start (2 in): each row is 1.25 in, as one
    // straight drag gives, whatever path the pointer took.
    expect(sizes()).toEqual({ w: before.w, h: ['1143000', '1143000'] });
    act(() => void editor.exec({ type: 'undo' }));
    expect(sizes()).toEqual(before);
    // A side handle changes the width only: the rows stay.
    const side = host.querySelector('[data-handle="w"]') as SVGCircleElement;
    const sx = Number(side.getAttribute('x')) + 4 * PX;
    const sy = Number(side.getAttribute('y')) + 4 * PX;
    const q = (dx: number) => ({ clientX: (sx - dx) / PX, clientY: sy / PX, bubbles: true });
    side.dispatchEvent(new PointerEvent('pointerdown', { ...q(0), button: 0 }));
    document.dispatchEvent(new PointerEvent('pointermove', q(IN / 2)));
    document.dispatchEvent(new PointerEvent('pointerup', q(IN / 2)));
    expect(sizes().h).toEqual(before.h);
    expect(sizes().w).not.toEqual(before.w);
    editor.destroy();
  });

  test('with a row that text holds, a drag gives the same table as a straight drag', async () => {
    const drag = async (path: number[]) => {
      const editor = await open();
      const host = surface(editor);
      const id = frames(editor)[0].id;
      const at0 = { paragraph: 0, offset: 1 };
      act(() =>
        editor.setSelection({
          kind: 'text',
          nodeId: id,
          cell: { row: 0, column: 0 },
          anchor: at0,
          focus: at0,
        })
      );
      for (let i = 0; i < 8; i += 1) act(() => void editor.input('paragraph'));
      act(() => editor.setSelection({ kind: 'shapes', nodeIds: [id] }));
      const handle = host.querySelector('[data-handle="n"]') as SVGCircleElement;
      const hx = Number(handle.getAttribute('x')) + 4 * PX;
      const hy = Number(handle.getAttribute('y')) + 4 * PX;
      const p = (dy: number) => ({ clientX: hx / PX, clientY: (hy - dy) / PX, bubbles: true });
      handle.dispatchEvent(new PointerEvent('pointerdown', { ...p(0), button: 0 }));
      for (const dy of path) document.dispatchEvent(new PointerEvent('pointermove', p(dy)));
      document.dispatchEvent(new PointerEvent('pointerup', p(path[path.length - 1])));
      const h = rowsOf(frames(editor)[0]).map((r) => attr(r, 'h'));
      editor.destroy();
      return h;
    };
    const straight = await drag([-IN / 4]);
    expect(await drag([IN, -2 * IN, -3 * IN, IN / 2, -IN / 4])).toEqual(straight);
  });

  test('a drag after a re-attach starts from the table as it is now', async () => {
    const editor = await open();
    const id = frames(editor)[0].id;
    const drag = (host: HTMLElement, dx: number) => {
      act(() => editor.setSelection({ kind: 'shapes', nodeIds: [id] }));
      const side = host.querySelector('[data-handle="w"]') as SVGCircleElement;
      const sx = Number(side.getAttribute('x')) + 4 * PX;
      const sy = Number(side.getAttribute('y')) + 4 * PX;
      const q = (d: number) => ({ clientX: (sx - d) / PX, clientY: sy / PX, bubbles: true });
      side.dispatchEvent(new PointerEvent('pointerdown', { ...q(0), button: 0 }));
      document.dispatchEvent(new PointerEvent('pointermove', q(dx)));
      document.dispatchEvent(new PointerEvent('pointerup', q(dx)));
    };
    drag(surface(editor), IN / 2);
    const cols = [2 * IN, IN, IN];
    expect(editor.exec({ type: 'setTableGrid', columns: cols }).ok).toBe(true);
    // A new surface numbers its gestures from 1 again.
    drag(surface(editor), IN / 2);
    const widths = children(find(frames(editor)[0], 'a:tblGrid'), 'a:gridCol').map((g) =>
      Number(attr(g, 'w'))
    );
    expect(widths.map((w) => Math.round((w / widths[1]) * 100) / 100)).toEqual([2, 1, 1]);
    editor.destroy();
  });

  test('deleting the table moves the focus to the slide', async () => {
    const editor = await open();
    const host = surface(editor);
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [frames(editor)[0].id] }));
    const outside = document.createElement('button');
    document.body.append(outside);
    outside.focus();
    expect(editor.exec({ type: 'deleteTable' }).ok).toBe(true);
    expect(host.contains(document.activeElement)).toBe(true);
    editor.destroy();
  });

  test('Shift+click on a table with shapes selected toggles it in the shape selection', async () => {
    const editor = await open();
    const host = surface(editor);
    const stage = stageOf(host);
    const id = frames(editor)[0].id;
    editor.setSelection({ kind: 'shapes', nodeIds: [id] });
    const press = { ...at(2.5 * IN, 2.5 * IN), button: 0, shiftKey: true };
    stage.dispatchEvent(new PointerEvent('pointerdown', press));
    document.dispatchEvent(new PointerEvent('pointerup', at(2.5 * IN, 2.5 * IN)));
    // As with any shape, Shift+click takes the selected table out of the shape selection.
    expect(editor.snapshot().selection.kind).not.toBe('text');
    expect(editor.snapshot().selection).not.toMatchObject({ nodeIds: [id] });
    editor.destroy();
  });

  test('a drag on an inner grid line resizes the two columns as one undo step', async () => {
    const editor = await open();
    const host = surface(editor);
    const stage = stageOf(host);
    const frame = frames(editor)[0];
    editor.setSelection({ kind: 'shapes', nodeIds: [frame.id] });
    // The line between columns 1 and 2 is at x = 2 in.
    stage.dispatchEvent(new PointerEvent('pointerdown', { ...at(2 * IN, 1.5 * IN), button: 0 }));
    document.dispatchEvent(new PointerEvent('pointermove', at(2 * IN + 10 * PX, 1.5 * IN)));
    document.dispatchEvent(new PointerEvent('pointermove', at(2 * IN + 20 * PX, 1.5 * IN)));
    document.dispatchEvent(new PointerEvent('pointerup', at(2 * IN + 20 * PX, 1.5 * IN)));
    const widths = () =>
      children(find(frames(editor)[0], 'a:tblGrid'), 'a:gridCol').map((g) => Number(attr(g, 'w')));
    expect(widths()).toEqual([IN + 20 * PX, IN - 20 * PX, IN]);
    act(() => void editor.exec({ type: 'undo' }));
    expect(widths()).toEqual([IN, IN, IN]);
    editor.destroy();
  });
});

describe('focus after a grid drag', () => {
  test('a caret that stays in the table takes the keys again after a line drag', async () => {
    const editor = await open();
    const host = surface(editor);
    const stage = stageOf(host);
    const frame = frames(editor)[0];
    const at0 = { paragraph: 0, offset: 1 };
    const cell = { row: 0, column: 0 };
    act(() =>
      editor.setSelection({ kind: 'text', nodeId: frame.id, cell, anchor: at0, focus: at0 })
    );
    stage.dispatchEvent(new PointerEvent('pointerdown', { ...at(2 * IN, 1.5 * IN), button: 0 }));
    document.dispatchEvent(new PointerEvent('pointermove', at(2 * IN + 20 * PX, 1.5 * IN)));
    document.dispatchEvent(new PointerEvent('pointerup', at(2 * IN + 20 * PX, 1.5 * IN)));
    expect(editor.snapshot().selection.kind).toBe('text');
    expect(document.activeElement?.className).toBe('pptx-surface__input');
    editor.destroy();
  });
});

describe('the border band of a selected table', () => {
  test('a press just inside the bottom edge places a caret; it never moves the table', async () => {
    const editor = await open();
    const host = surface(editor);
    const stage = stageOf(host);
    const frame = frames(editor)[0];
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [frame.id] }));
    // The bottom edge is at y = 3 in; the press is 6 px above it, in the second row's cell.
    const y = 3 * IN - 6 * PX;
    stage.dispatchEvent(new PointerEvent('pointerdown', { ...at(1.5 * IN, y), button: 0 }));
    document.dispatchEvent(new PointerEvent('pointermove', at(1.5 * IN, y + 25 * PX)));
    document.dispatchEvent(new PointerEvent('pointerup', at(1.5 * IN, y + 25 * PX)));
    const sel = editor.snapshot().selection;
    expect(sel.kind === 'text' || sel.kind === 'cells').toBe(true);
    const off = find(frames(editor)[0], 'p:xfrm')!;
    expect(attr(find(off, 'a:off')!, 'y')).toBe(String(IN));
    editor.destroy();
  });
});

describe('table edge over a shape', () => {
  test('a press just outside the edge resizes the selected table, not the shape behind', async () => {
    const editor = await open('behind');
    const host = surface(editor);
    const stage = stageOf(host);
    const frame = frames(editor)[0];
    const at0 = { paragraph: 0, offset: 0 };
    act(() =>
      editor.setSelection({
        kind: 'text',
        nodeId: frame.id,
        cell: { row: 0, column: 0 },
        anchor: at0,
        focus: at0,
      })
    );
    // The right edge is at x = 4 in; the press is 2 px outside it, over the shape.
    const x = 4 * IN + 2 * PX;
    stage.dispatchEvent(new PointerEvent('pointerdown', { ...at(x, 1.5 * IN), button: 0 }));
    document.dispatchEvent(new PointerEvent('pointermove', at(x + 20 * PX, 1.5 * IN)));
    document.dispatchEvent(new PointerEvent('pointerup', at(x + 20 * PX, 1.5 * IN)));
    const widths = children(find(frames(editor)[0], 'a:tblGrid'), 'a:gridCol').map((g) =>
      Number(attr(g, 'w'))
    );
    expect(widths).toEqual([IN, IN, IN + 20 * PX]);
    const root = xmlPartOf(editor.session!.store.package, part(editor))!.root;
    const sp = find(root, 'p:sp')!;
    expect(attr(find(sp, 'a:off')!, 'x')).toBe('0');
    editor.destroy();
  });
});

describe('table grid line under a shape', () => {
  test('a press on a shape in front of the table takes the shape, not the grid line', async () => {
    const editor = await open('front');
    const host = surface(editor);
    const stage = stageOf(host);
    const frame = frames(editor)[0];
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [frame.id] }));
    // The column line at x = 2 in runs under the box (1.5 in to 2.5 in).
    stage.dispatchEvent(new PointerEvent('pointerdown', { ...at(2 * IN, 1.5 * IN), button: 0 }));
    document.dispatchEvent(new PointerEvent('pointerup', at(2 * IN, 1.5 * IN)));
    const root = xmlPartOf(editor.session!.store.package, part(editor))!.root;
    const box = find(root, 'p:sp')!;
    const sel = editor.snapshot().selection;
    expect(sel.kind === 'shapes' ? sel.nodeIds : sel.kind).toEqual([box.id]);
    // The grid stays as it was: no column drag started.
    const widths = children(find(frames(editor)[0], 'a:tblGrid'), 'a:gridCol').map((g) =>
      Number(attr(g, 'w'))
    );
    expect(widths).toEqual([IN, IN, IN]);
    editor.destroy();
  });
});

describe('table toolbar', () => {
  const groups = (el: HTMLElement): string[] =>
    [...el.querySelectorAll<HTMLElement>('[role="toolbar"] > [role="group"]')].map(
      (g) => g.dataset.group ?? ''
    );

  test('table parts take an icon, the host element (asChild), or a class, and a raw hook inserts', async () => {
    const editor = await open();
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [frames(editor)[0].id] }));
    render(
      <PptxEditorContext.Provider value={editor}>
        <Parts preset={false}>
          <Parts.TableInsertRowBelow icon={<b data-testid="below-icon" />} />
          <Parts.TableMergeCells asChild>
            <button type="button" data-testid="host-merge" />
          </Parts.TableMergeCells>
          <Parts.TableInsert icon={<b data-testid="insert-icon" />} />
        </Parts>
      </PptxEditorContext.Provider>
    );
    expect(screen.getByTestId('below-icon').closest('[data-slot]')?.getAttribute('data-slot')).toBe(
      'table.insertRowBelow'
    );
    expect(screen.getByTestId('host-merge').getAttribute('data-slot')).toBe('table.mergeCells');
    expect(
      screen.getByTestId('insert-icon').closest('[data-slot]')?.getAttribute('data-slot')
    ).toBe('table.insert');
    act(() => void fireEvent.click(screen.getByTestId('below-icon')));
    expect(rowsOf(frames(editor)[0])).toHaveLength(3);
    function Raw() {
      const insert = useEditorCommand('table.insert');
      return (
        <button
          type="button"
          data-testid="raw"
          onClick={() => insert.execute({ rows: 2, cols: 2 })}
        />
      );
    }
    render(
      <PptxEditorContext.Provider value={editor}>
        <Raw />
      </PptxEditorContext.Provider>
    );
    act(() => void fireEvent.click(screen.getByTestId('raw')));
    expect(frames(editor)).toHaveLength(2);
  });

  test('asChild on a table dropdown picker warns in development', async () => {
    const editor = await open();
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [frames(editor)[0].id] }));
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      render(
        <PptxEditorContext.Provider value={editor}>
          <Toolbar>
            <ToolbarSlot slotId="table.style" asChild>
              <button type="button" />
            </ToolbarSlot>
          </Toolbar>
        </PptxEditorContext.Provider>
      );
      expect(warn.mock.calls.map((c: unknown[]) => String(c[0])).join('\n')).toContain(
        'table.style'
      );
    } finally {
      warn.mockRestore();
    }
  });

  test('a host hides or restyles the table group and its slots by id', async () => {
    const editor = await open();
    const hidden = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar>
          <ToolbarGroup id="table" hidden />
        </Toolbar>
      </PptxEditorContext.Provider>
    );
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar>
          <ToolbarSlot slotId="table.mergeCells" hidden />
          <ToolbarSlot slotId="table.insertRowBelow" className="host-below" />
        </Toolbar>
      </PptxEditorContext.Provider>
    );
    act(() =>
      editor.setSelection({
        kind: 'text',
        nodeId: frames(editor)[0].id,
        cell: { row: 0, column: 0 },
        anchor: { paragraph: 0, offset: 0 },
        focus: { paragraph: 0, offset: 0 },
      })
    );
    expect(groups(hidden.container)).not.toContain('table');
    expect(groups(container)).toContain('table');
    expect(container.querySelector('[data-slot="table.mergeCells"]') === null).toBe(true);
    const below = container.querySelector<HTMLButtonElement>('[data-slot="table.insertRowBelow"]')!;
    expect(below.classList.contains('host-below')).toBe(true);
    // Hiding a control changes the interface only: the command still runs.
    act(() =>
      editor.setSelection({
        kind: 'cells',
        nodeId: frames(editor)[0].id,
        anchor: { row: 0, column: 0 },
        focus: { row: 0, column: 1 },
      })
    );
    expect(editor.exec({ type: 'mergeCells' }).ok).toBe(true);
  });

  test('the table group shows in a table, after the bar; its controls run the commands', async () => {
    const editor = await open();
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar />
      </PptxEditorContext.Provider>
    );
    expect(groups(container)).not.toContain('table');
    const frame = frames(editor)[0];
    act(() =>
      editor.setSelection({
        kind: 'text',
        nodeId: frame.id,
        cell: { row: 0, column: 0 },
        anchor: { paragraph: 0, offset: 0 },
        focus: { paragraph: 0, offset: 0 },
      })
    );
    expect(groups(container).at(-1)).toBe('table');
    // The bar's short table group: the style options stay on the ribbon and in Toolbar.Slot.
    expect(container.querySelector('[data-slot="table.headerRow"]')).toBeNull();
    expect(container.querySelector('[data-slot="table.style"]')).not.toBeNull();
    act(() => void fireEvent.click(container.querySelector('[data-slot="table.insertRowBelow"]')!));
    expect(rowsOf(frames(editor)[0])).toHaveLength(3);
    // One cell: Merge Cells is disabled with the engine's reason.
    const merge = container.querySelector<HTMLButtonElement>('[data-slot="table.mergeCells"]')!;
    expect(merge.getAttribute('aria-disabled')).toBe('true');
    expect(tooltipText(merge)).toContain('Select two or more cells to merge');
  });

  test('Insert > Table picks the size from the grid; the formatting bar leaves it to menus', async () => {
    const editor = await open();
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar />
        <Toolbar preset={false}>
          <ToolbarSlot slotId="table.insert" />
        </Toolbar>
      </PptxEditorContext.Provider>
    );
    // The preset bar has no Insert > Table: a menu or the ribbon's Insert tab places it.
    expect(container.querySelectorAll('[data-slot="table.insert"]')).toHaveLength(1);
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Table' })));
    const cell = screen.getByRole('gridcell', { name: '3 × 2 table' });
    act(() => void fireEvent.mouseEnter(cell));
    expect(screen.getByText('3 × 2 table', { selector: 'p' })).toBeTruthy();
    act(() => void fireEvent.click(cell));
    const added = frames(editor)[1];
    expect(rowsOf(added)).toHaveLength(2);
    expect(children(rowsOf(added)[0], 'a:tc')).toHaveLength(3);
    expect(editor.snapshot().selection).toMatchObject({
      nodeId: added.id,
      cell: { row: 0, column: 0 },
    });
  });

  test('the size grid: the trigger names its popup, focus enters the grid, keys move it', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar preset={false}>
          <ToolbarSlot slotId="table.insert" />
        </Toolbar>
      </PptxEditorContext.Provider>
    );
    const trigger = screen.getByRole('button', { name: 'Table' });
    expect(trigger.getAttribute('aria-haspopup')).toBe('dialog');
    expect(trigger.getAttribute('aria-expanded')).toBe('false');
    act(() => void fireEvent.click(trigger));
    expect(trigger.getAttribute('aria-expanded')).toBe('true');
    const grid = screen.getByRole('grid', { name: 'Table' });
    // One tab stop, and it has the focus once the popup shows (the next frame).
    await act(() => new Promise<void>((r) => requestAnimationFrame(() => r())));
    expect(grid.querySelectorAll('[tabindex="0"]')).toHaveLength(1);
    expect(document.activeElement?.getAttribute('data-cell')).toBe('1x1');
    const key = async (k: string) => {
      act(() => void fireEvent.keyDown(document.activeElement!, { key: k }));
      await act(() => Promise.resolve());
    };
    await key('ArrowRight');
    await key('ArrowRight');
    await key('ArrowDown');
    expect(document.activeElement?.getAttribute('data-cell')).toBe('2x3');
    await key('End');
    expect(document.activeElement?.getAttribute('data-cell')).toBe('2x10');
    await key('Home');
    expect(document.activeElement?.getAttribute('data-cell')).toBe('2x1');
    // The caption follows the keys; it is not a live region.
    const caption = grid.querySelector('.pptx-table-grid__label')!;
    expect(caption.textContent).toBe('1 × 2 table');
    expect(caption.getAttribute('aria-hidden')).toBe('true');
    expect(grid.querySelector('[aria-live]') === null).toBe(true);
    // Escape closes the grid and gives the focus back to its trigger.
    await key('Escape');
    expect(screen.queryByRole('grid') === null).toBe(true);
    expect(document.activeElement).toBe(trigger);
    // Tab out of the grid (the focus leaves it) closes it too.
    act(() => void fireEvent.click(trigger));
    await act(() => new Promise<void>((r) => requestAnimationFrame(() => r())));
    act(() => void fireEvent.blur(document.activeElement!, { relatedTarget: document.body }));
    expect(screen.queryByRole('grid') === null).toBe(true);
  });

  test('Split Cells disables the presets the engine refuses, with the reason', async () => {
    const editor = await open();
    const nodeId = frames(editor)[0].id;
    act(() =>
      editor.setSelection({
        kind: 'cells',
        nodeId,
        anchor: { row: 0, column: 0 },
        focus: { row: 0, column: 1 },
      })
    );
    expect(editor.exec({ type: 'mergeCells' }).ok).toBe(true);
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar />
      </PptxEditorContext.Provider>
    );
    act(() =>
      editor.setSelection({
        kind: 'text',
        nodeId,
        cell: { row: 0, column: 0 },
        anchor: { paragraph: 0, offset: 0 },
        focus: { paragraph: 0, offset: 0 },
      })
    );
    act(() => void fireEvent.click(container.querySelector('[data-slot="table.splitCell"]')!));
    // A cell two columns wide splits into 2 columns, not into 3.
    const three = screen.getByRole('menuitem', { name: /3 columns/ });
    // Disabled but focusable, with the reason as its description.
    expect(three.getAttribute('aria-disabled')).toBe('true');
    expect(tooltipText(three)).toBe('A merged cell splits only into parts that divide it evenly');
    const two = screen.getByRole('menuitem', { name: /^2 columns$/ });
    expect(two.hasAttribute('disabled')).toBe(false);
  });

  test('a host grid with onPick gets the size and inserts nothing itself', async () => {
    const editor = await open();
    const picks: [number, number][] = [];
    render(
      <PptxEditorContext.Provider value={editor}>
        <MenuTableGrid className="host-grid" onPick={(rows, cols) => picks.push([rows, cols])} />
      </PptxEditorContext.Provider>
    );
    const grid = screen.getByRole('grid');
    expect(grid.classList.contains('host-grid')).toBe(true);
    act(() => void fireEvent.click(screen.getByRole('gridcell', { name: '4 × 3 table' })));
    expect(picks).toEqual([[3, 4]]);
    expect(frames(editor)).toHaveLength(1);
  });

  test('the style list names the deck styles and checks the table style', async () => {
    const editor = await open();
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar />
      </PptxEditorContext.Provider>
    );
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [frames(editor)[0].id] }));
    act(() => void fireEvent.click(container.querySelector('[data-slot="table.style"]')!));
    const item = screen.getByRole('menuitemradio', { name: 'Medium Style 2 – Accent 1' });
    expect(item.getAttribute('aria-checked')).toBe('true');
  });

  test('a disabled style list stays focusable, says why, and does not open', async () => {
    const editor = await open();
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar />
      </PptxEditorContext.Provider>
    );
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [frames(editor)[0].id] }));
    act(() => editor.setMode('view'));
    const trigger = container.querySelector<HTMLButtonElement>('[data-slot="table.style"]');
    expect(trigger !== null).toBe(true);
    if (!trigger) return;
    expect(trigger.disabled).toBe(false);
    expect(trigger.getAttribute('aria-disabled')).toBe('true');
    const note = document.getElementById(trigger.getAttribute('aria-describedby')!);
    expect(note?.textContent).toBeTruthy();
    act(() => void fireEvent.click(trigger));
    expect(screen.queryByRole('menuitemradio') === null).toBe(true);
  });
});
