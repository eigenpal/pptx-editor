// The Chart Data tab of the side panel in the React adapter: the empty state, the labeled fields
// that apply on Enter or blur as one undo step each, the refusal of text that is not a value, add
// and remove, and the chart slot that turns the side panel to the tab.
// `packages/vue/test/chart-data-pane.test.ts` checks that the Vue twin renders the same markup.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { PptxEditorContent, PptxEditorRoot } from '../PptxEditorRoot.tsx';
import { SidePanel, Toolbar } from '../index.ts';

const deck = new Uint8Array(
  readFileSync(join(import.meta.dir, '../../../../fixtures/chart-data.pptx'))
);

afterEach(() => {
  cleanup();
  document.body.replaceChildren();
});

async function mount(ui: React.ReactNode): Promise<PptxEditorInstance> {
  let editor: PptxEditorInstance | null = null;
  render(
    <PptxEditorRoot onReady={(e) => (editor = e)}>
      <PptxEditorContent />
      {ui}
    </PptxEditorRoot>
  );
  await act(async () => void (await editor!.load(deck)));
  return editor!;
}

function selectChart(editor: PptxEditorInstance): void {
  const chart = editor.snapshot().shapes!.find((s) => s.name.startsWith('Chart'))!;
  act(() => editor.setSelection({ kind: 'shapes', nodeIds: [chart.nodeId] }));
}

const values = (editor: PptxEditorInstance) => {
  const r = editor.chartData();
  if (!r.ok) throw new Error(r.reason);
  return r.data;
};

function type(field: HTMLElement, text: string, key = 'Enter'): void {
  act(() => void fireEvent.change(field, { target: { value: text } }));
  act(() => void fireEvent.keyDown(field, { key }));
}

describe('Chart Data tab (React)', () => {
  test('the tab shows only for a chart, and the chart slot opens it', async () => {
    const editor = await mount(
      <>
        <Toolbar preset={false}>
          <Toolbar.Slot slotId="chart.editData" />
        </Toolbar>
        <SidePanel />
      </>
    );
    expect(screen.queryByRole('tab', { name: 'Chart' })).toBe(null);
    selectChart(editor);
    const tab = screen.getByRole('tab', { name: 'Chart' });
    expect(tab.getAttribute('aria-selected')).toBe('false');
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Edit chart…' })));
    expect(tab.getAttribute('aria-selected')).toBe('true');
    expect(document.activeElement).toBe(tab);
    // With the chart deselected, the open tab stays and says to select a chart.
    act(() => editor.setSelection({ kind: 'none' }));
    expect(screen.getByText('No chart selected')).toBeTruthy();
  });

  test('a field applies on Enter as one undo step, labeled by series and category', async () => {
    const editor = await mount(<SidePanel tab="chart" />);
    selectChart(editor);
    const before = values(editor).series[0]!.values[1];
    const field = screen.getByRole('textbox', { name: 'Plan, South' });
    type(field, '77');
    expect(values(editor).series[0]!.values[1]).toBe(77);
    expect((field as HTMLInputElement).value).toBe('77');
    act(() => void editor.exec({ type: 'undo' }));
    expect(values(editor).series[0]!.values[1]).toBe(before);
    expect((field as HTMLInputElement).value).toBe(String(before));
  });

  test('text that is not a value stays in its field with an error, and Escape restores it', async () => {
    const editor = await mount(<SidePanel tab="chart" />);
    selectChart(editor);
    const field = screen.getByRole('textbox', { name: 'Plan, North' });
    act(() => void fireEvent.change(field, { target: { value: 'many' } }));
    act(() => void fireEvent.blur(field));
    expect(field.getAttribute('aria-invalid')).toBe('true');
    const error = document.getElementById(field.getAttribute('aria-describedby')!);
    expect(error?.textContent).toBe('Enter a number, or leave the cell empty.');
    expect(values(editor).series[0]!.values[0]).toBe(42);
    act(() => void fireEvent.keyDown(field, { key: 'Escape' }));
    expect((field as HTMLInputElement).value).toBe('42');
    expect(field.getAttribute('aria-invalid')).toBe('false');
  });

  test('adds and removes categories and series, and renames a series', async () => {
    const editor = await mount(<SidePanel tab="chart" />);
    selectChart(editor);
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Add category' })));
    expect(values(editor).categories.map((c) => c.label)).toEqual([
      'North',
      'South',
      'East',
      'West',
      'Category 5',
    ]);
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Remove category South' })));
    expect(values(editor).categories.map((c) => c.label)).not.toContain('South');
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Add series' })));
    expect(values(editor).series.map((s) => s.name)).toEqual([
      'Plan',
      'Actual',
      'Forecast',
      'Series 4',
    ]);
    type(screen.getByRole('textbox', { name: 'Name of series 4' }), 'Stretch');
    expect(values(editor).series[3]!.name).toBe('Stretch');
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Remove series Actual' })));
    expect(values(editor).series.map((s) => s.name)).toEqual(['Plan', 'Forecast', 'Stretch']);
  });

  test('an add focuses the new row, and a remove focuses the row that takes its place', async () => {
    const editor = await mount(<SidePanel tab="chart" />);
    selectChart(editor);
    const frame = () => act(() => new Promise<void>((r) => requestAnimationFrame(() => r())));
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Add category' })));
    await frame();
    expect(document.activeElement).toBe(screen.getByRole('textbox', { name: 'Category 5' }));
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Remove category South' })));
    await frame();
    expect(document.activeElement).toBe(screen.getByRole('textbox', { name: 'Category 2' }));
    expect((document.activeElement as HTMLInputElement).value).toBe('East');
  });

  test('removing a category clears rejected drafts before indexes change', async () => {
    const editor = await mount(<SidePanel tab="chart" />);
    selectChart(editor);
    type(screen.getByRole('textbox', { name: 'Plan, South' }), 'many');
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Remove category North' })));
    const east = screen.getByRole('textbox', { name: 'Plan, East' }) as HTMLInputElement;
    expect(east.value).toBe('51');
    expect(east.getAttribute('aria-invalid')).toBe('false');
    expect(values(editor).series[0]!.values).toEqual([35, 51, 28]);
    act(() => void fireEvent.blur(east));
    expect(values(editor).series[0]!.values).toEqual([35, 51, 28]);
  });

  test('selecting a chart on another slide clears rejected drafts', async () => {
    const editor = await mount(<SidePanel tab="chart" />);
    selectChart(editor);
    type(screen.getByRole('textbox', { name: 'Plan, South' }), 'many');
    act(() => editor.goToSlide(1));
    selectChart(editor);
    expect(document.querySelector('input[aria-invalid="true"]')).toBe(null);
    expect(
      screen.getAllByRole('textbox').some((field) => (field as HTMLInputElement).value === 'many')
    ).toBe(false);
  });

  test('removing a series clears a rejected name from the next series', async () => {
    const editor = await mount(<SidePanel tab="chart" />);
    selectChart(editor);
    type(screen.getByRole('textbox', { name: 'Name of series 2' }), 'Plan');
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Remove series Plan' })));
    const name = screen.getByRole('textbox', { name: 'Name of series 2' }) as HTMLInputElement;
    expect(name.value).toBe('Forecast');
    expect(name.getAttribute('aria-invalid')).toBe('false');
  });

  test('undo and redo keep focus in an existing clean field', async () => {
    const editor = await mount(<SidePanel tab="chart" />);
    selectChart(editor);
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Add category' })));
    const field = screen.getByRole('textbox', { name: 'Category 1' });
    field.focus();
    act(() => void fireEvent.keyDown(field, { key: 'z', ctrlKey: true }));
    expect(values(editor).categories.length).toBe(4);
    expect(document.activeElement).toBe(field);
    act(() => void fireEvent.keyDown(field, { key: 'y', ctrlKey: true }));
    expect(values(editor).categories.length).toBe(5);
    expect(document.activeElement).toBe(field);
  });

  test('appends and other value edits preserve pending typing and focus', async () => {
    const editor = await mount(<SidePanel tab="chart" />);
    selectChart(editor);
    const field = screen.getByRole('textbox', { name: 'Plan, East' }) as HTMLInputElement;
    field.focus();
    act(() => void fireEvent.change(field, { target: { value: '77' } }));
    const before = values(editor);
    act(
      () =>
        void editor.exec({
          type: 'setChartData',
          data: {
            ...before,
            categories: [...before.categories, { label: 'Added' }],
            series: [
              ...before.series.map((series) => ({ ...series, values: [...series.values, null] })),
              { name: 'Added series', values: [1, 2, 3, 4, 5] },
            ],
          },
        })
    );
    expect(field.value).toBe('77');
    expect(document.activeElement).toBe(field);
    expect(values(editor).series[0]!.values[2]).toBe(51);
    const appended = values(editor);
    act(
      () =>
        void editor.exec({
          type: 'setChartData',
          data: {
            ...appended,
            series: appended.series.map((series, index) =>
              index === 1 ? { ...series, values: series.values.map(() => 12) } : series
            ),
          },
        })
    );
    expect(field.value).toBe('77');
    expect(document.activeElement).toBe(field);
    act(() => void fireEvent.keyDown(field, { key: 'Enter' }));
    expect(values(editor).series[0]!.values[2]).toBe(77);
  });

  test('switching between equal-size charts clears drafts without replacing fields', async () => {
    const editor = await mount(<SidePanel tab="chart" />);
    selectChart(editor);
    const original = editor.snapshot().chart!.nodeId;
    act(() => void editor.exec({ type: 'duplicate' }));
    expect(editor.snapshot().chart!.nodeId).not.toBe(original);
    const field = screen.getByRole('textbox', { name: 'Plan, East' }) as HTMLInputElement;
    type(field, 'many');
    field.focus();
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [original] }));
    expect(screen.getByRole('textbox', { name: 'Plan, East' })).toBe(field);
    expect(field.value).toBe('51');
    expect(field.getAttribute('aria-invalid')).toBe('false');
    expect(document.activeElement).toBe(field);
  });

  test('a name that the workbook refuses stays in its field with the reason', async () => {
    const editor = await mount(<SidePanel tab="chart" />);
    selectChart(editor);
    // The embedded workbook's table cannot hold two series with the same name.
    type(screen.getByRole('textbox', { name: 'Name of series 2' }), 'Plan');
    const field = screen.getByRole('textbox', { name: 'Name of series 2' });
    expect(field.getAttribute('aria-invalid')).toBe('true');
    expect(values(editor).series[1]!.name).toBe('Actual');
  });
});
