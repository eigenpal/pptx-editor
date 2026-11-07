import { afterEach, describe, expect, test } from 'bun:test';
import { createApp, defineComponent, h, nextTick, shallowRef, type App, type Component } from 'vue';
import {
  runToolbarCommand,
  type PptxEditorInstance,
  type SidePanelTab,
} from '@pptx-editor.dev/core/editor';
import { buildPptx, shape } from '../../../scripts/fixtures/builder.ts';
import {
  listenPaneRequests,
  useEditorCommand,
  PptxEditorContent,
  PptxEditorRoot,
  PptxEditorSelectionPane,
  PptxEditorSidePanel,
  PptxEditorViewport,
} from '../src/index.ts';

const IN = 914_400;
const group =
  '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Cards"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/><a:chOff x="0" y="0"/>' +
  '<a:chExt cx="1828800" cy="914400"/></a:xfrm></p:grpSpPr>' +
  shape({ id: 11, name: 'Card A', x: 0, y: 0, cx: IN, cy: IN }) +
  '</p:grpSp>';

let app: App | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
  document.body.replaceChildren();
});

async function mount(inner: () => unknown): Promise<PptxEditorInstance> {
  host = document.createElement('div');
  document.body.append(host);
  let ready: PptxEditorInstance | null = null;
  app = createApp(
    defineComponent({
      setup: () => () =>
        h(PptxEditorRoot, { onReady: (e: PptxEditorInstance) => (ready = e) }, () =>
          h(PptxEditorViewport, null, () => [h(PptxEditorContent), inner()])
        ),
    }) as Component
  );
  app.mount(host);
  await nextTick();
  const deck = buildPptx({
    slides: [
      {
        shapes:
          shape({ id: 2, name: 'One', x: IN, y: IN, cx: IN, cy: IN }) +
          shape({ id: 3, name: 'Two', x: 2 * IN, y: IN, cx: IN, cy: IN }) +
          group,
      },
      { shapes: shape({ id: 2, name: 'Other', x: IN, y: IN, cx: IN, cy: IN }) },
    ],
  });
  const editor = ready as PptxEditorInstance | null;
  expect((await editor!.load(deck)).ok).toBe(true);
  await nextTick();
  return editor!;
}

const rows = () => [...document.querySelectorAll<HTMLElement>('[role="treeitem"]')];
const row = (name: string) => rows().find((r) => r.getAttribute('aria-label')?.startsWith(name))!;
const key = async (el: Element, k: string, init: KeyboardEventInit = {}): Promise<void> => {
  el.dispatchEvent(
    new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true, ...init })
  );
  await nextTick();
  await nextTick();
};

describe('PptxEditorSelectionPane', () => {
  test('lists the slide top first, selects, hides, renames, and shows all', async () => {
    const editor = await mount(() => h(PptxEditorSelectionPane));
    expect(rows().map((r) => r.getAttribute('aria-label'))).toEqual(['Cards', 'Two', 'One']);
    row('Two').dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
    await nextTick();
    const two = editor.snapshot().shapes!.find((s) => s.name === 'Two')!.nodeId;
    expect(editor.snapshot().selection).toEqual({ kind: 'shapes', nodeIds: [two] });
    expect(row('Two').getAttribute('aria-selected')).toBe('true');
    expect(document.activeElement).toBe(row('Two'));
    row('Two').querySelector<HTMLButtonElement>('[aria-label="Hide Two"]')!.click();
    await nextTick();
    expect(row('Two').getAttribute('aria-label')).toBe('Two (hidden)');
    await key(row('One'), 'F2');
    const field = document.querySelector<HTMLInputElement>('input[aria-label="Name"]')!;
    expect(document.activeElement).toBe(field);
    field.value = 'Headline';
    await key(field, 'Enter');
    expect(row('Headline')).toBeDefined();
    expect(document.activeElement).toBe(row('Headline'));
    await key(row('Headline'), 'F2');
    const again = document.querySelector<HTMLInputElement>('input[aria-label="Name"]')!;
    again.value = 'Other';
    await key(again, 'Escape');
    expect(row('Headline')).toBeDefined();
    document.querySelector<HTMLButtonElement>('.pptx-selection-pane__bar .pptx-button')!.click();
    await nextTick();
    expect(editor.snapshot().shapes!.every((s) => !s.hidden)).toBe(true);
  });

  test('a group opens with the arrow keys and for a selected child', async () => {
    const editor = await mount(() => h(PptxEditorSelectionPane));
    expect(row('Cards').getAttribute('aria-expanded')).toBe('false');
    await key(row('Cards'), 'ArrowRight');
    expect(row('Cards').getAttribute('aria-expanded')).toBe('true');
    expect(row('Card A').getAttribute('aria-level')).toBe('2');
    await key(row('Cards'), 'ArrowLeft');
    expect(row('Card A')).toBeUndefined();
    const cards = editor.snapshot().shapes!.find((s) => s.name === 'Cards')!;
    editor.setSelection({
      kind: 'shapes',
      nodeIds: [cards.children![0]!.nodeId],
      group: cards.nodeId,
    });
    await nextTick();
    await nextTick();
    expect(row('Card A').getAttribute('aria-selected')).toBe('true');
  });

  test('the side panel shows it on its Selection tab, as a controlled tab', async () => {
    const tab = shallowRef<SidePanelTab>('selection');
    const picked: SidePanelTab[] = [];
    await mount(() =>
      h(PptxEditorSidePanel, {
        tab: tab.value,
        'onUpdate:tab': (t: SidePanelTab) => (picked.push(t), (tab.value = t)),
      })
    );
    const tabOf = (name: string) =>
      [...document.querySelectorAll<HTMLElement>('[role="tab"]')].find(
        (t) => t.textContent === name
      )!;
    expect(tabOf('Selection').getAttribute('aria-selected')).toBe('true');
    expect(document.querySelector('[role="tree"]') !== null).toBe(true);
    tabOf('Format').click();
    await nextTick();
    expect(picked).toEqual(['format']);
    expect(document.querySelector('[role="tree"]')).toBe(null);
  });
  test('H shows or hides the focused row; Ctrl and Cmd keys pass by', async () => {
    const editor = await mount(() => h(PptxEditorSelectionPane));
    const hidden = () => editor.snapshot().shapes!.find((s) => s.name === 'Two')!.hidden;
    expect(row('Two').getAttribute('aria-keyshortcuts')).toBe('H F2');
    // The row's description names the key hint.
    const hint = document.getElementById(row('Two').getAttribute('aria-describedby')!);
    expect(hint?.textContent).toContain('Press H');
    await key(row('Two'), 'h', { ctrlKey: true });
    expect(hidden()).toBeFalsy();
    await key(row('Two'), 'h');
    expect(hidden()).toBe(true);
    await key(row('Two'), 'H');
    expect(hidden()).toBeFalsy();
  });

  test('Show all reaches a group child; a new slide closes the rename field', async () => {
    const editor = await mount(() => h(PptxEditorSelectionPane));
    const button = (name: string) =>
      [...document.querySelectorAll<HTMLButtonElement>('button')].find(
        (b) => b.textContent === name
      )!;
    const cards = () => editor.snapshot().shapes!.find((s) => s.name === 'Cards')!;
    const kid = cards().children![0]!.nodeId;
    editor.exec({ type: 'setShapesVisible', nodeIds: [kid], visible: false });
    await nextTick();
    button('Show all').click();
    await nextTick();
    expect(cards().children![0]!.hidden).toBeFalsy();
    // Hide all hides the top level only: showing the group again shows its children.
    button('Hide all').click();
    await nextTick();
    expect(cards().hidden).toBe(true);
    expect(cards().children![0]!.hidden).toBeFalsy();
    await key(row('One'), 'F2');
    const field = document.querySelector<HTMLInputElement>('input[aria-label="Name"]')!;
    expect(field.value).toBe('One');
    // Node ids repeat across slides: the field must not rename an element of the next slide.
    editor.goToSlide(1);
    await nextTick();
    await nextTick();
    expect(document.querySelector('input[aria-label="Name"]')).toBe(null);
    expect(rows().map((r) => r.getAttribute('aria-label'))).toEqual(['Other']);
    expect(editor.snapshot().shapes![0]!.name).toBe('Other');
  });

  test('the toolbar Selection pane slot turns the side panel to its tab', async () => {
    const emitted: SidePanelTab[] = [];
    const editor = await mount(() =>
      h(PptxEditorSidePanel, { 'onUpdate:tab': (t: SidePanelTab) => emitted.push(t) })
    );
    const selected = () =>
      document.querySelector('[role="tab"][aria-selected="true"]')?.textContent ?? null;
    expect(selected()).toBe('Format');
    expect(runToolbarCommand(editor, 'arrange.selectionPane')).toEqual({ ok: true });
    await nextTick();
    await nextTick();
    expect(selected()).toBe('Selection');
    // The focus goes to the tab, so a keyboard user lands in the panel.
    expect(document.activeElement?.textContent).toBe('Selection');
    expect(emitted).toEqual(['selection']);
    app!.unmount();
    app = null;
    // With no side panel, the slot says so.
    expect(runToolbarCommand(editor, 'arrange.selectionPane')).toEqual({
      ok: false,
      reason: 'disabled.noSidePanel',
    });
  });
  test('without a side panel, the slot is off and a host listener turns it on', async () => {
    const Probe = defineComponent({
      setup() {
        const cmd = useEditorCommand('arrange.selectionPane');
        return () =>
          h('output', { id: 'probe' }, `${cmd.isEnabled.value} ${cmd.disabledReason.value}`);
      },
    });
    const editor = await mount(() => h(Probe));
    const probe = () => document.getElementById('probe')!.textContent;
    expect(probe()).toBe('false disabled.noSidePanel');
    const stop = listenPaneRequests(editor, () => {});
    await nextTick();
    expect(probe()).toBe('true null');
    stop();
    await nextTick();
    expect(probe()).toBe('false disabled.noSidePanel');
  });
});
