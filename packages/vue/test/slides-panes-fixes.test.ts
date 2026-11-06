// The Vue twin of packages/react/src/__tests__/slides-panes-fixes.test.tsx: the open-failure
// notice, the section menu's keys and the focus after Move section up, the rail keys in the
// viewing mode and Home and End without sections, notes that store the formats they show, the
// find bar on a new document and with matches but no current one, Cmd/Ctrl+F in the slide
// sorter, and the focus after the sorter goes away.

import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { createApp, defineComponent, h, nextTick, shallowRef, type App, type Component } from 'vue';
import {
  createPptxEditor,
  resetDevWarnings,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import { buildSlideLayoutsDeck } from '../../../scripts/fixtures/slide-layouts.ts';
import {
  PptxEditor,
  PptxEditorEditNotice,
  PptxEditorFindBar,
  PptxEditorNotesPane,
  PptxEditorOpenFailure,
  PptxEditorRoot,
  PptxEditorSlideRail,
  PptxEditorSlideSorter,
  PptxEditorToolbar,
  PptxEditorViewport,
  requestFindBar,
} from '../src/index.ts';
import { providePptxEditor, useEditorState } from '../src/composables.ts';

let app: App | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
  document.body.replaceChildren();
});

const IN = 914_400;
const tick = async (): Promise<void> => {
  await nextTick();
  await nextTick();
};

const flatDeck = (count = 3, notes?: string) =>
  buildPptx({
    slides: Array.from({ length: count }, (_, i) => ({
      layout: 'blank' as const,
      shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run(`S${i + 1}`)) }),
      ...(notes && i === 0 ? { notes } : {}),
    })),
  });

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

async function mount(render: () => unknown, editor?: PptxEditorInstance): Promise<HTMLElement> {
  host = document.createElement('div');
  host.className = 'pptx-editor';
  document.body.append(host);
  app = createApp(
    defineComponent({
      setup() {
        if (editor) {
          providePptxEditor(shallowRef(editor));
        }
        return render;
      },
    }) as Component
  );
  app.mount(host);
  await tick();
  return host;
}

const byLabel = (role: string, name: RegExp | string): HTMLElement => {
  const all = [...document.querySelectorAll<HTMLElement>(`[role="${role}"]`)];
  // A slide item takes its name from hidden text (`.pptx-sr-only`), not `aria-label`.
  const label = (el: HTMLElement): string | null =>
    el.getAttribute('aria-label') ??
    el.querySelector(':scope > .pptx-sr-only')?.textContent ??
    null;
  const found = all.find((el) =>
    typeof name === 'string' ? label(el) === name : name.test(label(el) ?? el.textContent ?? '')
  );
  if (!found) throw new Error(`no ${role} ${String(name)}`);
  return found;
};
const key = (el: Element, init: KeyboardEventInit): void =>
  void el.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init }));
const notice = () => document.querySelector('.pptx-edit-notice')?.textContent ?? '';

describe('open failure', () => {
  test('a refused file shows why, keeps the open deck, and shows again for the next refusal', async () => {
    const editor = await open(flatDeck());
    let dismissed = 0;
    await mount(() => h(PptxEditorOpenFailure, { onDismiss: () => (dismissed += 1) }), editor);
    expect(document.querySelector('.pptx-open-failure') === null).toBe(true);
    await editor.load(new Uint8Array(17).fill(7));
    await tick();
    expect(editor.snapshot().status).toBe('ready');
    expect(editor.snapshot().slides.length).toBe(3);
    const el = document.querySelector<HTMLElement>('[role="alert"]')!;
    expect(el.textContent).toContain('The file did not open');
    expect(el.textContent).toContain('The presentation that was open stays open.');
    expect(el.textContent).toContain('Error code: zip');
    document.querySelector<HTMLButtonElement>('.pptx-open-failure__action')!.click();
    await tick();
    expect(dismissed).toBe(1);
    expect(document.querySelector('.pptx-open-failure') === null).toBe(true);
    await editor.load(new Uint8Array(17).fill(7));
    await tick();
    expect(document.querySelector('.pptx-open-failure') !== null).toBe(true);
    await editor.load(flatDeck(2));
    await tick();
    expect(document.querySelector('.pptx-open-failure') === null).toBe(true);
  });

  test('the default editor mounts the notice with the openFailureActions slot', async () => {
    let editor: PptxEditorInstance | null = null;
    await mount(() =>
      h(
        PptxEditor,
        { onReady: (e: PptxEditorInstance) => (editor = e) },
        {
          openFailureActions: ({ failure }: { failure: { code: string } }) =>
            h('a', { class: 'pptx-open-failure__action', href: `#help-${failure.code}` }, 'Help'),
        }
      )
    );
    await editor!.load(new Uint8Array([1, 2, 3]));
    await tick();
    expect(document.querySelector('a.pptx-open-failure__action')?.getAttribute('href')).toBe(
      '#help-zip'
    );
  });
});

describe('open failure props', () => {
  test('the message and actions props render, and a slot wins over its prop', async () => {
    const editor = await open(flatDeck());
    await mount(
      () =>
        h(
          PptxEditorOpenFailure,
          {
            message: (failure: { code: string }) => `Host text ${failure.code}`,
            actions: () => h('a', { class: 'from-prop' }, 'Prop'),
          },
          { actions: () => h('a', { class: 'from-slot' }, 'Slot') }
        ),
      editor
    );
    await editor.load(new Uint8Array(17).fill(7));
    await tick();
    expect(document.querySelector('.pptx-open-failure__message')?.textContent).toBe(
      'Host text zip'
    );
    expect(document.querySelectorAll('.from-slot').length).toBe(1);
    expect(document.querySelectorAll('.from-prop').length).toBe(0);
  });

  test('PptxEditor forwards the openFailure object, and warns when the slot shadows it', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => undefined);
    try {
      let editor: PptxEditorInstance | null = null;
      await mount(() =>
        h(
          PptxEditor,
          {
            onReady: (e: PptxEditorInstance) => (editor = e),
            openFailure: { message: 'Host message', actions: () => h('a', { class: 'from-prop' }) },
          },
          { openFailureActions: () => h('a', { class: 'from-slot' }) }
        )
      );
      await editor!.load(new Uint8Array([1, 2, 3]));
      await tick();
      expect(document.querySelector('.pptx-open-failure__message')?.textContent).toBe(
        'Host message'
      );
      expect(document.querySelectorAll('.from-slot').length).toBe(1);
      expect(document.querySelectorAll('.from-prop').length).toBe(0);
      const text = warn.mock.calls.map((c) => String(c[0])).join('\n');
      expect(text).toContain('The #openFailureActions slot replaces openFailure.actions.');
    } finally {
      warn.mockRestore();
    }
  });

  test('the openFailureActions slot with :open-failure="false" warns', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => undefined);
    try {
      await mount(() =>
        h(PptxEditor, { openFailure: false }, { openFailureActions: () => h('a') })
      );
      const text = warn.mock.calls.map((c) => String(c[0])).join('\n');
      expect(text).toContain(':open-failure="false"');
    } finally {
      warn.mockRestore();
    }
  });
});

describe('section menu', () => {
  test('Enter on a row adds no slide; Move section up keeps focus on the header', async () => {
    const editor = await open(buildSlideLayoutsDeck());
    const before = editor.snapshot().slides.length;
    await mount(() => h(PptxEditorSlideRail), editor);
    const header = byLabel('treeitem', /^Section Details,/);
    header.focus();
    key(header, { key: 'F10', shiftKey: true });
    await tick();
    const row = byLabel('menuitem', /Move section up/);
    row.focus();
    key(row, { key: 'Enter' });
    row.click();
    await tick();
    await new Promise((r) => setTimeout(r, 0));
    expect(editor.snapshot().slides.length).toBe(before);
    expect((editor.snapshot().sections ?? []).map((s) => s.name)).toEqual(['Details', 'Opening']);
    expect(document.activeElement?.getAttribute('aria-label') ?? '').toMatch(/^Section Details,/);
  });
});

describe('rail keys', () => {
  test('in the viewing mode, Enter, Cmd+D, and Delete say why they do nothing', async () => {
    const editor = await open(flatDeck());
    editor.setMode('view');
    await mount(() => h(PptxEditorSlideRail), editor);
    for (const init of [{ key: 'Enter' }, { key: 'd', metaKey: true }, { key: 'Delete' }]) {
      key(byLabel('option', 'Slide 1'), init);
      expect(editor.snapshot().lastRejection?.reason).toBe('disabled.viewing');
    }
    expect(editor.snapshot().slides.length).toBe(3);
  });

  test('Home and End go to the first and the last slide of a rail without sections', async () => {
    const editor = await open(flatDeck(4));
    await mount(() => h(PptxEditorSlideRail), editor);
    key(byLabel('option', 'Slide 1'), { key: 'End' });
    expect(editor.snapshot().currentSlide).toBe(3);
    await tick();
    key(byLabel('option', 'Slide 4'), { key: 'Home' });
    expect(editor.snapshot().currentSlide).toBe(0);
  });
});

describe('slide sorter', () => {
  test('Cmd+F in the sorter shows why find does not open', async () => {
    const editor = await open(flatDeck());
    editor.exec({ type: 'setPresentationView', view: 'slideSorter' });
    await mount(
      () => [h(PptxEditorSlideSorter), h(PptxEditorFindBar), h(PptxEditorEditNotice)],
      editor
    );
    const tile = byLabel('option', 'Slide 1');
    tile.focus();
    key(tile, { key: 'f', metaKey: true });
    await tick();
    expect(notice()).toBe('Open the Normal view to find text in the slides');
    expect(document.querySelector<HTMLElement>('.pptx-find')!.hidden).toBe(true);
  });

  test('the Find button turns off in the slide sorter and on again in the Normal view', async () => {
    let editor: PptxEditorInstance | null = null;
    host = document.createElement('div');
    document.body.append(host);
    app = createApp(
      defineComponent({
        setup: () => () =>
          h(PptxEditorRoot, { onReady: (e: PptxEditorInstance) => (editor = e) }, () =>
            h(PptxEditorViewport, null, () => [h(PptxEditorToolbar), h(PptxEditorFindBar)])
          ),
      }) as Component
    );
    app.mount(host);
    await tick();
    expect((await editor!.load(flatDeck())).ok).toBe(true);
    await tick();
    const off = (): string | null =>
      document.querySelector('button[data-slot="find.find"]')!.getAttribute('aria-disabled');
    expect(off()).toBe(null);
    editor!.exec({ type: 'setPresentationView', view: 'slideSorter' });
    await tick();
    expect(off()).toBe('true');
    editor!.exec({ type: 'setPresentationView', view: 'normal' });
    await tick();
    expect(off()).toBe(null);
  });
});

describe('find bar', () => {
  const type = (label: string, value: string): void => {
    const input = document.querySelector<HTMLInputElement>(`input[aria-label="${label}"]`)!;
    input.value = value;
    input.dispatchEvent(new Event('input', { bubbles: true }));
  };
  const count = () => document.querySelector('.pptx-find__count')!.textContent;

  test('a new document closes the bar and clears the search', async () => {
    const editor = await open(flatDeck());
    await mount(() => h(PptxEditorFindBar), editor);
    requestFindBar(editor);
    await tick();
    type('Find', 'S2');
    await tick();
    expect(count()).toBe('1 of 1');
    await editor.load(flatDeck(2));
    await tick();
    expect(document.querySelector<HTMLElement>('.pptx-find')!.hidden).toBe(true);
    requestFindBar(editor);
    await tick();
    expect(document.querySelector<HTMLInputElement>('input[aria-label="Find"]')!.value).toBe('');
  });

  test('matches with none shown give the count alone; Replace all keeps focus in the field', async () => {
    const editor = await open(flatDeck());
    await mount(() => h(PptxEditorFindBar), editor);
    requestFindBar(editor, true);
    await tick();
    type('Find', 's');
    type('Replace', 'S');
    await tick();
    const all = () =>
      [...document.querySelectorAll<HTMLButtonElement>('.pptx-find button')].find(
        (b) => b.textContent === 'Replace all'
      )!;
    all().focus();
    all().click();
    await tick();
    expect(count()).toBe('3 matches');
    type('Find', 'S');
    type('Replace', 'Q');
    await tick();
    all().focus();
    all().click();
    await tick();
    expect(count()).toBe('No results');
    expect(document.activeElement?.getAttribute('aria-label')).toBe('Find');
    const nav = [...document.querySelectorAll('.pptx-find button')].filter((b) =>
      /match/.test(b.getAttribute('aria-label') ?? '')
    );
    expect(nav.every((b) => b.getAttribute('aria-disabled') === 'true')).toBe(true);
  });
});

describe('notes formats', () => {
  const field = () => document.querySelector<HTMLElement>('.pptx-notes__text')!;

  test('text typed into an underlined run stays underlined after the commit', async () => {
    const editor = await open(flatDeck(1, 'Hi'));
    await mount(() => h(PptxEditorNotesPane), editor);
    field().focus();
    const span = document.createElement('span');
    span.className = 'pptx-notes__run--underline';
    span.textContent = ' under\nmore';
    field().append(span);
    field().dispatchEvent(new Event('input', { bubbles: true }));
    field().blur();
    await tick();
    expect(
      editor.snapshot().notesFormatted!.map((p) => p.runs.map((r) => [r.text, r.underline]))
    ).toEqual([
      [
        ['Hi', false],
        [' under', true],
      ],
      [['more', true]],
    ]);
  });

  test('B with a caret outside a word arms bold for the next typed text', async () => {
    const editor = await open(flatDeck(1, 'Hi'));
    await mount(() => h(PptxEditorNotesPane), editor);
    field().focus();
    const sel = document.getSelection()!;
    sel.selectAllChildren(field());
    sel.collapseToEnd();
    document.dispatchEvent(new Event('selectionchange'));
    await tick();
    const bold = byLabel('group', 'Notes formatting').querySelector<HTMLButtonElement>('button')!;
    expect(bold.getAttribute('aria-disabled') === 'true').toBe(false);
    bold.click();
    await tick();
    expect(bold.getAttribute('aria-pressed')).toBe('true');
    field().dispatchEvent(
      new InputEvent('beforeinput', {
        inputType: 'insertText',
        data: '!',
        bubbles: true,
        cancelable: true,
      })
    );
    field().blur();
    await tick();
    expect(editor.snapshot().notesFormatted![0]!.runs.map((r) => [r.text, r.bold])).toEqual([
      ['Hi', false],
      ['!', true],
    ]);
  });
});

describe('view switch focus', () => {
  test('a sorter hidden while its slide has focus gives focus to the current slide of the rail', async () => {
    const editor = await open(flatDeck());
    editor.exec({ type: 'setPresentationView', view: 'slideSorter' });
    // A host layout that swaps the rail and the sorter, as the examples do.
    const Views = defineComponent({
      setup() {
        const sorter = useEditorState((s) => s.presentationView === 'slideSorter');
        return () => (sorter.value ? h(PptxEditorSlideSorter) : h(PptxEditorSlideRail));
      },
    });
    const el = await mount(() => h(Views), editor);
    editor.goToSlide(1);
    await tick();
    el.querySelectorAll<HTMLElement>('.pptx-sorter__item')[1]!.focus();
    editor.exec({ type: 'setPresentationView', view: 'normal' });
    await tick();
    await new Promise<void>((r) => requestAnimationFrame(() => r()));
    const current = el.querySelector('.pptx-rail__list [aria-current="true"]');
    expect(current).not.toBeNull();
    expect(document.activeElement).toBe(current);
  });
});
