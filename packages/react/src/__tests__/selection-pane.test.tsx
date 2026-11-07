import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import { useState } from 'react';
import type { PptxEditorInstance, SidePanelTab } from '@pptx-editor.dev/core/editor';
import { buildPptx, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContent, PptxEditorRoot } from '../PptxEditorRoot.tsx';
import { SelectionPane, SidePanel } from '../index.ts';

const IN = 914_400;
const group =
  '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Cards"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/><a:chOff x="0" y="0"/>' +
  '<a:chExt cx="1828800" cy="914400"/></a:xfrm></p:grpSpPr>' +
  shape({ id: 11, name: 'Card A', x: 0, y: 0, cx: IN, cy: IN }) +
  '</p:grpSp>';

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
  const deck = buildPptx({
    slides: [
      {
        shapes:
          shape({ id: 2, name: 'One', x: IN, y: IN, cx: IN, cy: IN }) +
          shape({ id: 3, name: 'Two', x: 2 * IN, y: IN, cx: IN, cy: IN }) +
          group,
      },
    ],
  });
  await act(async () => void (await editor!.load(deck)));
  return editor!;
}

const row = (name: string) => screen.getByRole('treeitem', { name: new RegExp(`^${name}`) });

describe('PptxEditor.SelectionPane', () => {
  test('lists the slide top first, selects, hides, renames, and shows all', async () => {
    const editor = await mount(<SelectionPane />);
    expect(screen.getAllByRole('treeitem').map((r) => r.getAttribute('aria-label'))).toEqual([
      'Cards',
      'Two',
      'One',
    ]);
    // A press on a row selects its element and keeps focus in the pane.
    act(() => void fireEvent.mouseDown(row('Two')));
    const two = editor.snapshot().shapes!.find((s) => s.name === 'Two')!.nodeId;
    expect(editor.snapshot().selection).toEqual({ kind: 'shapes', nodeIds: [two] });
    expect(row('Two').getAttribute('aria-selected')).toBe('true');
    expect(document.activeElement).toBe(row('Two'));
    // The eye hides the element; the row says so.
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Hide Two' })));
    expect(row('Two').getAttribute('aria-label')).toBe('Two (hidden)');
    expect(editor.snapshot().shapes!.find((s) => s.name === 'Two')!.hidden).toBe(true);
    // F2 renames; Enter applies and gives focus back to the row.
    act(() => void fireEvent.keyDown(row('One'), { key: 'F2' }));
    const field = screen.getByRole('textbox', { name: 'Name' });
    expect(document.activeElement).toBe(field);
    act(() => void fireEvent.change(field, { target: { value: 'Headline' } }));
    act(() => void fireEvent.keyDown(field, { key: 'Enter' }));
    expect(row('Headline')).toBeDefined();
    expect(document.activeElement).toBe(row('Headline'));
    // Escape keeps the old name.
    act(() => void fireEvent.keyDown(row('Headline'), { key: 'F2' }));
    const again = screen.getByRole('textbox', { name: 'Name' });
    act(() => void fireEvent.change(again, { target: { value: 'Other' } }));
    act(() => void fireEvent.keyDown(again, { key: 'Escape' }));
    expect(row('Headline')).toBeDefined();
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Show all' })));
    expect(editor.snapshot().shapes!.every((s) => !s.hidden)).toBe(true);
  });

  test('a group opens with the arrow keys and for a selected child', async () => {
    const editor = await mount(<SelectionPane />);
    expect(row('Cards').getAttribute('aria-expanded')).toBe('false');
    act(() => void fireEvent.keyDown(row('Cards'), { key: 'ArrowRight' }));
    expect(row('Cards').getAttribute('aria-expanded')).toBe('true');
    expect(row('Card A').getAttribute('aria-level')).toBe('2');
    act(() => void fireEvent.keyDown(row('Cards'), { key: 'ArrowLeft' }));
    expect(screen.queryByRole('treeitem', { name: 'Card A' })).toBe(null);
    const cards = editor.snapshot().shapes!.find((s) => s.name === 'Cards')!;
    act(() =>
      editor.setSelection({
        kind: 'shapes',
        nodeIds: [cards.children![0]!.nodeId],
        group: cards.nodeId,
      })
    );
    expect(row('Card A').getAttribute('aria-selected')).toBe('true');
  });

  test('the side panel shows it on its Selection tab, as a controlled tab', async () => {
    const picked: SidePanelTab[] = [];
    function Host() {
      const [tab, setTab] = useState<SidePanelTab>('selection');
      return <SidePanel tab={tab} onTabChange={(t) => (picked.push(t), setTab(t))} />;
    }
    await mount(<Host />);
    expect(screen.getByRole('tab', { name: 'Selection' }).getAttribute('aria-selected')).toBe(
      'true'
    );
    expect(screen.getByRole('tree')).toBeDefined();
    act(() => void fireEvent.click(screen.getByRole('tab', { name: 'Format' })));
    expect(picked).toEqual(['format']);
    expect(screen.queryByRole('tree')).toBe(null);
  });
});
