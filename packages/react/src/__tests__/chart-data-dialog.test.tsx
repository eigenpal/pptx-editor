import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  canvasMenuKind,
  createPptxEditor,
  toolbarContextOf,
  toolbarCommandState,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { PptxEditorContext } from '../context.ts';
import { ChartDataDialog } from '../index.ts';

// The Edit Data dialog: the grid of the selected chart's data, its keyboard moves, the row and
// column edits, OK as one undo step, and a chart whose data cannot be edited.

const FIXTURE = join(import.meta.dir, '../../../../fixtures/chart-data.pptx');

async function open(bytes = new Uint8Array(readFileSync(FIXTURE))): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(bytes)).ok).toBe(true);
  const chart = editor.snapshot().shapes!.find((s) => s.name === 'Chart 1')!;
  act(() => editor.setSelection({ kind: 'shapes', nodeIds: [chart.nodeId] }));
  return editor;
}

function show(editor: PptxEditorInstance, onClose: () => void = () => {}) {
  return render(
    <PptxEditorContext.Provider value={editor}>
      <ChartDataDialog open onClose={onClose} />
    </PptxEditorContext.Provider>
  );
}

const cell = (name: string): HTMLInputElement => screen.getByLabelText(name) as HTMLInputElement;
const button = (name: string): HTMLButtonElement =>
  screen.getByRole('button', { name }) as HTMLButtonElement;
const values = (editor: PptxEditorInstance) => {
  const r = editor.chartData();
  if (!r.ok) throw new Error(r.reason);
  return r.data;
};

afterEach(cleanup);

describe('chart chrome', () => {
  test('a selected chart shows the chart group and the chart menu', async () => {
    const editor = await open();
    const snap = editor.snapshot();
    expect(toolbarContextOf(snap)).toBe('chart');
    expect(canvasMenuKind(snap)).toBe('chart');
    expect(toolbarCommandState(editor, 'chart.editData')).toMatchObject({ enabled: true });
    act(() => editor.setSelection({ kind: 'none' }));
    expect(toolbarCommandState(editor, 'chart.editData')).toMatchObject({
      enabled: false,
      disabledReason: 'disabled.noChart',
    });
  });
});

describe('Edit Data dialog', () => {
  test('shows the data as a grid; OK applies the edits as one undo step', async () => {
    const editor = await open();
    let closed = 0;
    show(editor, () => (closed += 1));
    const grid = screen.getByRole('grid');
    expect(grid.getAttribute('aria-rowcount')).toBe('5');
    expect(grid.getAttribute('aria-colcount')).toBe('4');
    expect(cell('Name of series 1').value).toBe('Plan');
    expect(cell('Category 2').value).toBe('South');
    expect(cell('Actual, East').value).toBe('47');
    act(() => void fireEvent.change(cell('Plan, North'), { target: { value: '50' } }));
    act(() => void fireEvent.change(cell('Name of series 3'), { target: { value: 'Outlook' } }));
    // Remove South: focus a cell of its row, then Remove category.
    act(() => void fireEvent.focus(cell('Plan, South')));
    act(() => void fireEvent.click(button('Remove category')));
    // A new series goes after the column of the focused cell.
    act(() => void fireEvent.click(button('Add series')));
    expect(screen.getByRole('grid').getAttribute('aria-colcount')).toBe('5');
    expect(values(editor).series[0]!.values[0]).toBe(42);
    act(() => void fireEvent.click(document.querySelector('[data-pptx-part="apply"]')!));
    expect(closed).toBe(1);
    const data = values(editor);
    expect(data.categories.map((c) => c.label)).toEqual(['North', 'East', 'West']);
    expect(data.series.map((s) => s.name)).toEqual(['Plan', 'Series 4', 'Actual', 'Outlook']);
    expect(data.series[0]!.values).toEqual([50, 51, 28]);
    expect(data.series[1]!.values).toEqual([null, null, null]);
    act(() => void editor.exec({ type: 'undo' }));
    expect(values(editor).series.map((s) => s.name)).toEqual(['Plan', 'Actual', 'Forecast']);
  });

  test('the arrow keys and Return move between cells; one cell is in the tab order', async () => {
    const editor = await open();
    show(editor);
    const first = cell('Plan, North');
    expect(first.tabIndex).toBe(0);
    expect(cell('Plan, South').tabIndex).toBe(-1);
    act(() => first.focus());
    act(() => void fireEvent.keyDown(first, { key: 'ArrowDown' }));
    expect(document.activeElement).toBe(cell('Plan, South'));
    act(() => void fireEvent.keyDown(document.activeElement!, { key: 'Enter' }));
    expect(document.activeElement).toBe(cell('Plan, East'));
    const east = cell('Plan, East');
    east.setSelectionRange(east.value.length, east.value.length);
    act(() => void fireEvent.keyDown(east, { key: 'ArrowRight' }));
    expect(document.activeElement).toBe(cell('Actual, East'));
    expect(cell('Actual, East').tabIndex).toBe(0);
    act(() => void fireEvent.keyDown(document.activeElement!, { key: 'ArrowUp', ctrlKey: true }));
    expect(document.activeElement).toBe(cell('Actual, North'));
  });

  test('a value that is not a number blocks OK with a named, focused error', async () => {
    const editor = await open();
    let closed = 0;
    show(editor, () => (closed += 1));
    act(() => void fireEvent.change(cell('Actual, West'), { target: { value: 'many' } }));
    act(() => void fireEvent.click(document.querySelector('[data-pptx-part="apply"]')!));
    expect(closed).toBe(0);
    const bad = cell('Actual, West');
    expect(bad.getAttribute('aria-invalid')).toBe('true');
    const error = document.querySelector('[data-pptx-part="error"]')!;
    expect(error.textContent).toBe('Actual, West: Enter a number, or leave the cell empty.');
    expect(bad.getAttribute('aria-describedby')).toBe(error.id);
    expect(document.activeElement).toBe(bad);
    expect(values(editor).series[1]!.values[3]).toBe(33);
  });

  test('a chart whose data cannot be edited shows why', async () => {
    const files = unzipSync(new Uint8Array(readFileSync(FIXTURE)));
    const xml = strFromU8(files['ppt/charts/chart1.xml']!);
    files['ppt/charts/chart1.xml'] = strToU8(xml.replaceAll('c:barChart>', 'c:stockChart>'));
    const editor = await open(zipSync(files));
    show(editor);
    expect(screen.queryByRole('grid')).toBeNull();
    expect(screen.getByRole('status').textContent).toBe(
      'You cannot edit the data of this chart. The editor does not edit the data of this chart type.'
    );
    expect((document.querySelector('[data-pptx-part="apply"]') as HTMLButtonElement).disabled).toBe(
      true
    );
  });
});
