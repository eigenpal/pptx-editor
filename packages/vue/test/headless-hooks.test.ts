// The composables with DOCX Editor's names: `useChromeTranslate` and the compounds' `t`, the
// toolbar label composables, `useHistoryGroup`, `useEditorCaret`, and `useDocumentSearch`. The
// twin of the React test (packages/react/src/__tests__/headless-hooks.test.tsx).

import { afterEach, describe, expect, test } from 'bun:test';
import { createApp, defineComponent, h, nextTick, shallowRef, type App } from 'vue';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import { providePptxEditor } from '../src/composables.ts';
import {
  PptxEditorSlideRail,
  PptxEditorStatusBar,
  useChromeTranslate,
  useDocumentSearch,
  useEditorCaret,
  useHistoryGroup,
  useToolbarContext,
  useToolbarLabel,
  type ChromeTranslate,
  type UseDocumentSearchReturn,
  type UseEditorCaretReturn,
  type UseHistoryGroupReturn,
} from '../src/index.ts';

const apps: [App, HTMLElement][] = [];
const editors: PptxEditorInstance[] = [];
afterEach(() => {
  for (const [app, host] of apps.splice(0)) {
    app.unmount();
    host.remove();
  }
  for (const editor of editors.splice(0)) editor.destroy();
});

function mount(render: () => ReturnType<typeof h>): void {
  const host = document.createElement('div');
  document.body.append(host);
  const app = createApp({ render });
  app.mount(host);
  apps.push([app, host]);
}

const IN = 914400;
async function textDeck(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  editors.push(editor);
  const text = (name: string, words: string) =>
    shape({ id: 2, name, x: IN, y: IN, cx: 6 * IN, cy: IN, text: para(run(words)) });
  await editor.load(
    buildPptx({ slides: [{ shapes: text('T', 'one two one') }, { shapes: text('U', 'one') }] })
  );
  return editor;
}

/** Render `children` under a host component that provides `editor`. */
function inEditor(editor: PptxEditorInstance, children: () => ReturnType<typeof h>[]) {
  return h(
    defineComponent({
      setup() {
        providePptxEditor(shallowRef(editor));
        return () => children();
      },
    })
  );
}

const host: ChromeTranslate = (key) => `host:${key}`;

const Labels = defineComponent({
  props: { id: { type: String, required: true } },
  setup(props) {
    const t = useChromeTranslate();
    const label = useToolbarLabel();
    const context = useToolbarContext();
    return () =>
      h(
        'output',
        { 'data-testid': props.id },
        `${t('statusBar.notes')}|${label('statusBar.notes')}|${context.t.value ? 'host' : 'none'}`
      );
  },
});

const byTestId = (id: string) => document.querySelector(`[data-testid="${id}"]`)!.textContent;

describe('the t prop and the translation composables (Vue)', () => {
  test('a compound with t labels itself and its parts from the host', async () => {
    const editor = await textDeck();
    mount(() =>
      inEditor(editor, () => [
        h(PptxEditorStatusBar, { t: host }, () => [h(Labels, { id: 'inside' })]),
        h(PptxEditorSlideRail, { t: host }),
        h(Labels, { id: 'outside' }),
      ])
    );
    await nextTick();
    expect(document.querySelector('.pptx-statusbar')!.getAttribute('aria-label')).toBe(
      'host:statusBar.ariaLabel'
    );
    expect(document.querySelector('[aria-label="host:statusBar.notes"]') !== null).toBe(true);
    expect(document.querySelector('[aria-label="host:rail.collapse"]') !== null).toBe(true);
    expect(byTestId('inside')).toBe('host:statusBar.notes|host:statusBar.notes|host');
    expect(byTestId('outside')).toBe('Notes|Notes|none');
  });

  test('a change of t keeps the compound mounted, with its state and focus', async () => {
    const editor = await textDeck();
    const t = shallowRef<ChromeTranslate | undefined>(host);
    const Counter = defineComponent({
      setup() {
        const count = shallowRef(0);
        return () =>
          h(
            'button',
            { type: 'button', 'data-testid': 'count', onClick: () => (count.value += 1) },
            String(count.value)
          );
      },
    });
    mount(() =>
      inEditor(editor, () => [h(PptxEditorStatusBar, { t: t.value }, () => [h(Counter)])])
    );
    await nextTick();
    const button = document.querySelector<HTMLElement>('[data-testid="count"]')!;
    button.click();
    button.focus();
    t.value = undefined;
    await nextTick();
    expect(document.querySelector('[data-testid="count"]')).toBe(button);
    expect(button.textContent).toBe('1');
    expect(document.activeElement).toBe(button);
    expect(document.querySelector('[aria-label="Notes"]') !== null).toBe(true);
    t.value = host;
    await nextTick();
    expect(document.querySelector('[data-testid="count"]')).toBe(button);
    expect(document.querySelector('[aria-label="host:statusBar.notes"]') !== null).toBe(true);
  });

  test('PptxEditorStatusBar.Action is a host button after Notes', async () => {
    const editor = await textDeck();
    let sent = 0;
    mount(() =>
      inEditor(editor, () => [
        h(PptxEditorStatusBar, null, () => [
          h(PptxEditorStatusBar.Action, { label: 'Share', onSelect: () => (sent += 1) }),
        ]),
      ])
    );
    await nextTick();
    const share = document.querySelector<HTMLElement>('[aria-label="Share"]')!;
    const notes = document.querySelector<HTMLElement>('[aria-label="Notes"]')!;
    expect(notes.compareDocumentPosition(share) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    share.click();
    expect(sent).toBe(1);
  });
});

describe('useHistoryGroup (Vue)', () => {
  test('the commands of one gesture are one undo step', async () => {
    const editor = await textDeck();
    let group!: UseHistoryGroupReturn;
    const Slider = defineComponent({
      setup() {
        group = useHistoryGroup({ kind: 'range' });
        return () => h('input', { type: 'range', 'data-testid': 'slider', ref: group.ref });
      },
    });
    mount(() => inEditor(editor, () => [h(Slider)]));
    await nextTick();
    const slider = document.querySelector<HTMLElement>('[data-testid="slider"]')!;
    slider.dispatchEvent(new PointerEvent('pointerdown', { button: 0, pointerId: 1 }));
    editor.exec({ type: 'addSlide' }, group.options());
    editor.exec({ type: 'addSlide' }, group.options());
    document.dispatchEvent(new PointerEvent('pointerup', { pointerId: 1 }));
    await new Promise((resolve) => setTimeout(resolve, 5));
    expect(editor.snapshot().slides.length).toBe(4);
    editor.exec({ type: 'undo' });
    expect(editor.snapshot().slides.length).toBe(2);
  });
});

describe('useHistoryGroup with an element from another window (Vue)', () => {
  test('binds the element, and ignores a node that is not an element', async () => {
    const editor = await textDeck();
    let group!: UseHistoryGroupReturn;
    const Probe = defineComponent({
      setup() {
        group = useHistoryGroup({ kind: 'range' });
        return () => null;
      },
    });
    mount(() => inEditor(editor, () => [h(Probe)]));
    const input = document.createElement('input');
    document.body.append(input);
    // An iframe's element is no instance of this window's `HTMLElement`.
    const own = globalThis.HTMLElement;
    globalThis.HTMLElement = class Foreign {} as unknown as typeof HTMLElement;
    try {
      expect(input instanceof HTMLElement).toBe(false);
      group.ref(input);
    } finally {
      globalThis.HTMLElement = own;
    }
    expect(() => group.options()).not.toThrow();
    group.ref(document.createTextNode('x'));
    expect(() => group.options()).toThrow();
  });
});

describe('useEditorCaret (Vue)', () => {
  test('reads the caret of a text edit, and null otherwise', async () => {
    const editor = await textDeck();
    let caret!: UseEditorCaretReturn;
    const Probe = defineComponent({
      setup() {
        caret = useEditorCaret();
        return () => null;
      },
    });
    mount(() => inEditor(editor, () => [h(Probe)]));
    expect(caret.value).toBeNull();
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    editor.setSelection({
      kind: 'text',
      nodeId,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 3 },
    });
    expect(caret.value).toEqual({ nodeId, paragraph: 0, offset: 3 });
    const first = caret.value;
    editor.setSelection({
      kind: 'text',
      nodeId,
      anchor: { paragraph: 0, offset: 1 },
      focus: { paragraph: 0, offset: 3 },
    });
    expect(caret.value).toBe(first);
    editor.setSelection({ kind: 'none' });
    expect(caret.value).toBeNull();
  });
});

describe('useDocumentSearch (Vue)', () => {
  test('finds, steps, goes to a match, and clears', async () => {
    const editor = await textDeck();
    let search!: UseDocumentSearchReturn;
    const Probe = defineComponent({
      setup() {
        search = useDocumentSearch();
        return () => null;
      },
    });
    mount(() => inEditor(editor, () => [h(Probe)]));
    search.setQuery('one');
    expect(search.matches.value.length).toBe(3);
    expect(search.activeIndex.value).toBe(-1);
    search.next();
    expect(search.activeIndex.value).toBe(0);
    search.goTo(2);
    expect(search.activeIndex.value).toBe(2);
    expect(editor.snapshot().currentSlide).toBe(1);
    search.clear();
    expect(search.query.value).toBe('');
    expect(search.matches.value.length).toBe(0);
  });

  test('a composable that mounts again reads the matches of the deck as it is now', async () => {
    const editor = await textDeck();
    let search!: UseDocumentSearchReturn;
    const Probe = defineComponent({
      setup() {
        search = useDocumentSearch();
        return () => null;
      },
    });
    const show = shallowRef(true);
    mount(() => inEditor(editor, () => (show.value ? [h(Probe)] : [])));
    search.setQuery('one');
    expect(search.matches.value.length).toBe(3);
    show.value = false;
    await nextTick();
    // An edit with no composable mounted: the second slide goes, with its match.
    expect(editor.exec({ type: 'deleteSlides', indices: [1] }).ok).toBe(true);
    show.value = true;
    await nextTick();
    expect(search.query.value).toBe('one');
    expect(search.matches.value.length).toBe(2);
  });

  test('outside a root it has no matches and does nothing', () => {
    let search!: UseDocumentSearchReturn;
    const Probe = defineComponent({
      setup() {
        search = useDocumentSearch();
        return () => null;
      },
    });
    mount(() => h(Probe));
    expect(search.matches.value.length).toBe(0);
    expect(search.next()).toBe(false);
  });
});
