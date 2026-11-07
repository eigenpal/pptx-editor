import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { createApp, defineComponent, h, nextTick, shallowRef, type App, type Component } from 'vue';
import {
  canvasMenuKind,
  createPptxEditor,
  toolbarContextOf,
  toolbarCommandState,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { providePptxEditor } from '../src/composables.ts';
import { ChartDataDialog } from '../src/index.ts';

// The Vue twin of packages/react/src/__tests__/chart-data-dialog.test.tsx: the grid of the
// selected chart's data, its keyboard moves, the row and column edits, OK as one undo step, and
// a chart whose data cannot be edited.

const FIXTURE = join(import.meta.dir, '../../../fixtures/chart-data.pptx');

let app: App | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
});

async function open(bytes = new Uint8Array(readFileSync(FIXTURE))): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(bytes)).ok).toBe(true);
  const chart = editor.snapshot().shapes!.find((s) => s.name === 'Chart 1')!;
  editor.setSelection({ kind: 'shapes', nodeIds: [chart.nodeId] });
  return editor;
}

async function show(editor: PptxEditorInstance): Promise<{ closed: number }> {
  const state = { closed: 0 };
  host = document.createElement('div');
  document.body.append(host);
  app = createApp(
    defineComponent({
      setup() {
        providePptxEditor(shallowRef(editor));
        return () =>
          h(ChartDataDialog as Component, { open: true, onClose: () => (state.closed += 1) });
      },
    }) as Component
  );
  app.mount(host);
  await nextTick();
  return state;
}

const cell = (name: string): HTMLInputElement =>
  document.querySelector<HTMLInputElement>(`input[aria-label="${name}"]`)!;
const button = (name: string): HTMLButtonElement =>
  [...document.querySelectorAll<HTMLButtonElement>('button')].find((b) => b.textContent === name)!;
const grid = (): HTMLElement => document.querySelector<HTMLElement>('[role="grid"]')!;
const apply = (): HTMLButtonElement =>
  document.querySelector<HTMLButtonElement>('[data-pptx-part="apply"]')!;

async function type(el: HTMLInputElement, value: string): Promise<void> {
  el.value = value;
  el.dispatchEvent(new Event('input', { bubbles: true }));
  await nextTick();
}

async function key(el: Element, init: KeyboardEventInit): Promise<void> {
  el.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init }));
  await nextTick();
}

const values = (editor: PptxEditorInstance) => {
  const r = editor.chartData();
  if (!r.ok) throw new Error(r.reason);
  return r.data;
};

describe('chart chrome (Vue)', () => {
  test('a selected chart shows the chart group and the chart menu', async () => {
    const editor = await open();
    const snap = editor.snapshot();
    expect(toolbarContextOf(snap)).toBe('chart');
    expect(canvasMenuKind(snap)).toBe('chart');
    expect(toolbarCommandState(editor, 'chart.editData')).toMatchObject({ enabled: true });
    editor.setSelection({ kind: 'none' });
    expect(toolbarCommandState(editor, 'chart.editData')).toMatchObject({
      enabled: false,
      disabledReason: 'disabled.noChart',
    });
  });
});

describe('ChartDataDialog (Vue)', () => {
  test('shows the data as a grid; OK applies the edits as one undo step', async () => {
    const editor = await open();
    const state = await show(editor);
    expect(grid().getAttribute('aria-rowcount')).toBe('5');
    expect(grid().getAttribute('aria-colcount')).toBe('4');
    expect(cell('Name of series 1').value).toBe('Plan');
    expect(cell('Category 2').value).toBe('South');
    expect(cell('Actual, East').value).toBe('47');
    await type(cell('Plan, North'), '50');
    await type(cell('Name of series 3'), 'Outlook');
    // Remove South: focus a cell of its row, then Remove category.
    cell('Plan, South').dispatchEvent(new FocusEvent('focus'));
    await nextTick();
    button('Remove category').click();
    await nextTick();
    // A new series goes after the column of the focused cell.
    button('Add series').click();
    await nextTick();
    expect(grid().getAttribute('aria-colcount')).toBe('5');
    expect(values(editor).series[0]!.values[0]).toBe(42);
    apply().click();
    await nextTick();
    expect(state.closed).toBe(1);
    const data = values(editor);
    expect(data.categories.map((c) => c.label)).toEqual(['North', 'East', 'West']);
    expect(data.series.map((s) => s.name)).toEqual(['Plan', 'Series 4', 'Actual', 'Outlook']);
    expect(data.series[0]!.values).toEqual([50, 51, 28]);
    expect(data.series[1]!.values).toEqual([null, null, null]);
    editor.exec({ type: 'undo' });
    expect(values(editor).series.map((s) => s.name)).toEqual(['Plan', 'Actual', 'Forecast']);
  });

  test('the arrow keys and Return move between cells; one cell is in the tab order', async () => {
    const editor = await open();
    await show(editor);
    const first = cell('Plan, North');
    expect(first.tabIndex).toBe(0);
    expect(cell('Plan, South').tabIndex).toBe(-1);
    first.focus();
    await nextTick();
    await key(first, { key: 'ArrowDown' });
    expect(document.activeElement).toBe(cell('Plan, South'));
    await key(document.activeElement!, { key: 'Enter' });
    expect(document.activeElement).toBe(cell('Plan, East'));
    const east = cell('Plan, East');
    east.setSelectionRange(east.value.length, east.value.length);
    await key(east, { key: 'ArrowRight' });
    expect(document.activeElement).toBe(cell('Actual, East'));
    expect(cell('Actual, East').tabIndex).toBe(0);
    await key(document.activeElement!, { key: 'ArrowUp', ctrlKey: true });
    expect(document.activeElement).toBe(cell('Actual, North'));
  });

  test('a value that is not a number blocks OK with a named, focused error', async () => {
    const editor = await open();
    const state = await show(editor);
    await type(cell('Actual, West'), 'many');
    apply().click();
    await nextTick();
    expect(state.closed).toBe(0);
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
    await show(editor);
    expect(grid()).toBeNull();
    expect(document.querySelector('[role="status"]')!.textContent).toBe(
      'You cannot edit the data of this chart. The editor does not edit the data of this chart type.'
    );
    expect(apply().disabled).toBe(true);
  });
});
