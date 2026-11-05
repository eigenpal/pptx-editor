import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import {
  createApp,
  defineComponent,
  h,
  nextTick,
  ref,
  shallowRef,
  type App,
  type Component,
} from 'vue';
import {
  createPptxEditor,
  toolbarGroupsFor,
  resetDevWarnings,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import { buildSlideLayoutsDeck } from '../../../scripts/fixtures/slide-layouts.ts';
import {
  PptxEditor,
  PptxEditorContent,
  PptxEditorRoot,
  PptxEditorToolbar,
  PptxEditorViewport,
  ToolbarGroup,
  ToolbarSlot,
  useEditorCommand,
  useEditorState,
  useEditorValueCommand,
} from '../src/index.ts';
import { editorStateActiveSubscriptionCount, providePptxEditor } from '../src/composables.ts';

const IN = 914_400;
const deckBytes = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run('Hi')) }),
      },
    ],
  });

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(deckBytes())).ok).toBe(true);
  return editor;
}

/** Select the slide's first shape. */
const selectBox = (editor: PptxEditorInstance): void =>
  editor.setSelection({
    kind: 'shapes',
    nodeIds: [editor.session!.layout(0)!.items[0]!.element.nodeId],
  });

let app: App | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
  resetDevWarnings();
});

/** Mount a component whose `setup` is `setup`, with `editor` provided as `PptxEditorRoot` does. */
function mountWith(editor: PptxEditorInstance | null, setup: () => () => unknown): HTMLElement {
  // A second mount in one test replaces the first: no app stays in the document for a later
  // file's `document.querySelector` to find.
  app?.unmount();
  host?.remove();
  host = document.createElement('div');
  document.body.append(host);
  const Child = defineComponent({ setup });
  const Provider = defineComponent({
    setup() {
      providePptxEditor(shallowRef(editor));
      return () => h(Child);
    },
  });
  app = createApp(Provider as Component);
  app.mount(host);
  return host;
}

describe('PptxEditorRoot', () => {
  test('creates an instance, emits ready, and attaches the canvas', async () => {
    let ready: PptxEditorInstance | null = null;
    host = document.createElement('div');
    document.body.append(host);
    app = createApp({
      render: () =>
        h(
          PptxEditorRoot,
          { onReady: (e: PptxEditorInstance) => (ready = e) },
          { default: () => h(PptxEditorViewport, () => h(PptxEditorContent)) }
        ),
    });
    app.mount(host);
    await nextTick();
    expect(ready !== null).toBe(true);
    expect(host.querySelector('.pptx-editor .pptx-canvas[data-pptx-canvas]') !== null).toBe(true);
  });

  test('the PptxEditor sugar renders the toolbar and the canvas, with the host class', async () => {
    host = document.createElement('div');
    document.body.append(host);
    app = createApp({ render: () => h(PptxEditor, { class: 'acme-editor' }) });
    app.mount(host);
    await nextTick();
    expect(host.querySelector('.pptx-editor.pptx-default.acme-editor') !== null).toBe(true);
    expect(host.querySelector('.pptx-default [role="toolbar"]') !== null).toBe(true);
    expect(host.querySelector('.pptx-default .pptx-canvas') !== null).toBe(true);
  });

  test('a document change reloads, and mode and locale apply without a reload', async () => {
    let editor: PptxEditorInstance | null = null;
    const bytes = shallowRef<Uint8Array | undefined>(undefined);
    const mode = ref<'edit' | 'view'>('edit');
    const locale = shallowRef<Record<string, unknown> | undefined>(undefined);
    host = document.createElement('div');
    document.body.append(host);
    app = createApp({
      render: () =>
        h(
          PptxEditorRoot,
          {
            document: bytes.value,
            mode: mode.value,
            locale: locale.value,
            onReady: (e: PptxEditorInstance) => (editor = e),
          },
          { default: () => h(PptxEditorToolbar) }
        ),
    });
    app.mount(host);
    await nextTick();
    expect(editor!.snapshot().status).toBe('empty');

    bytes.value = deckBytes();
    await nextTick();
    for (let i = 0; i < 200 && editor!.snapshot().status !== 'ready'; i += 1)
      await new Promise((r) => setTimeout(r, 10));
    expect(editor!.snapshot().slides.length).toBe(1);

    mode.value = 'view';
    await nextTick();
    expect(editor!.snapshot().mode).toBe('view');
    expect(editor!.snapshot().slides.length).toBe(1);

    locale.value = { toolbar: { ariaLabel: 'Werkzeuge' } };
    await nextTick();
    expect(host.querySelector('[role="toolbar"]')!.getAttribute('aria-label')).toBe('Werkzeuge');
  });

  test('a locale change repaints the canvas prompts', async () => {
    let editor: PptxEditorInstance | null = null;
    const locale = shallowRef<Record<string, unknown> | undefined>(undefined);
    const bytes = buildSlideLayoutsDeck();
    host = document.createElement('div');
    document.body.append(host);
    app = createApp({
      render: () =>
        h(
          PptxEditorRoot,
          {
            document: bytes,
            locale: locale.value,
            onReady: (e: PptxEditorInstance) => (editor = e),
          },
          { default: () => h(PptxEditorContent) }
        ),
    });
    app.mount(host);
    await nextTick();
    for (let i = 0; i < 200 && editor!.snapshot().status !== 'ready'; i += 1)
      await new Promise((r) => setTimeout(r, 10));
    // Slide 5 has an empty title: the canvas shows its prompt.
    editor!.goToSlide(4);
    const canvas = (): string => host!.querySelector('.pptx-surface__stage')?.textContent ?? '';
    expect(canvas()).toContain('Click to add title');
    locale.value = { canvas: { prompts: { title: 'Titel hinzufügen' } } };
    await nextTick();
    expect(canvas()).toContain('Titel hinzufügen');
  });

  test('an inline locale with a change handler that sets state does not loop', async () => {
    let editor: PptxEditorInstance | null = null;
    const changes = ref(0);
    const bytes = buildSlideLayoutsDeck();
    const warn = console.warn;
    const warnings: string[] = [];
    console.warn = (...args: unknown[]) => void warnings.push(args.map(String).join(' '));
    try {
      host = document.createElement('div');
      document.body.append(host);
      app = createApp({
        render: () =>
          // Read the counter, so each change renders the host again.
          (
            void changes.value,
            h(
              PptxEditorRoot,
              {
                document: bytes,
                // A new object with the same content on each render.
                locale: { canvas: { prompts: { title: `Titel hinzufügen` } } },
                onReady: (e: PptxEditorInstance) => (editor = e),
                onChange: () => (changes.value += 1),
              },
              { default: () => h(PptxEditorContent) }
            )
          ),
      });
      app.mount(host);
      await nextTick();
      for (let i = 0; i < 200 && editor!.snapshot().status !== 'ready'; i += 1)
        await new Promise((r) => setTimeout(r, 10));
      editor!.goToSlide(4);
      await new Promise((r) => setTimeout(r, 50));
      const settled = changes.value;
      await new Promise((r) => setTimeout(r, 50));
      expect(changes.value).toBe(settled);
      expect(warnings.some((w) => /recursive/i.test(w))).toBe(false);
      expect(host.querySelector('.pptx-surface__stage')?.textContent).toContain('Titel hinzufügen');
    } finally {
      console.warn = warn;
    }
  });

  test('unmounting releases every editor subscription', async () => {
    const before = editorStateActiveSubscriptionCount();
    host = document.createElement('div');
    document.body.append(host);
    app = createApp({ render: () => h(PptxEditor) });
    app.mount(host);
    await nextTick();
    // One subscription for the whole tree: the root's change tick.
    expect(editorStateActiveSubscriptionCount()).toBe(before + 1);
    app.unmount();
    app = null;
    expect(editorStateActiveSubscriptionCount()).toBe(before);
  });
});

describe('composables', () => {
  test('useEditorState updates when its slice changes', async () => {
    const editor = await open();
    const el = mountWith(editor, () => {
      const count = useEditorState((s) => s.slides.length);
      return () => h('output', String(count.value));
    });
    expect(el.textContent).toBe('1');
    expect(editor.exec({ type: 'addSlide' }).ok).toBe(true);
    await nextTick();
    expect(el.textContent).toBe('2');
  });

  test('useEditorState keeps the same value when the slice is equal', async () => {
    const editor = await open();
    let renders = 0;
    mountWith(editor, () => {
      const status = useEditorState((s) => s.status);
      return () => {
        renders += 1;
        return h('output', status.value);
      };
    });
    const before = renders;
    editor.setZoom(2);
    await nextTick();
    expect(renders).toBe(before);
  });

  test('useEditorValueCommand runs a typed value', async () => {
    const editor = await open();
    const id = editor.session!.layout(0)!.items[0].element.nodeId;
    editor.setSelection({ kind: 'shapes', nodeIds: [id] });
    let rotate: ((v: 'right90') => { ok: boolean }) | null = null;
    let options: readonly string[] = [];
    mountWith(editor, () => {
      const cmd = useEditorValueCommand('arrange.rotate');
      rotate = cmd.execute;
      options = cmd.options.value;
      // @ts-expect-error A typo in a typed slot value does not compile.
      if (false as boolean) cmd.execute('right91');
      return () => h('output', String(cmd.isEnabled.value));
    });
    expect(rotate!('right90').ok).toBe(true);
    expect(editor.snapshot().geometry?.rot).toBe(90);
    expect(options).toEqual(['right90', 'left90', 'flipVertical', 'flipHorizontal']);
  });

  test('useEditorValueCommand gives numbers for a numeric slot', async () => {
    const editor = await open();
    const id = editor.session!.layout(0)!.items[0].element.nodeId;
    editor.setSelection({ kind: 'shapes', nodeIds: [id] });
    let width: ReturnType<typeof useEditorValueCommand<'shape.lineWidth'>> | null = null;
    mountWith(editor, () => {
      width = useEditorValueCommand('shape.lineWidth');
      return () => h('output', String(width!.value.value));
    });
    expect(width!.execute(3).ok).toBe(true);
    await nextTick();
    expect(width!.value.value).toBe(3);
    expect(width!.options.value.every((o) => typeof o === 'number')).toBe(true);
  });

  test('useEditorCommand follows a slot given as a ref', async () => {
    const editor = await open();
    expect(editor.exec({ type: 'addSlide' }).ok).toBe(true);
    const slot = ref<'history.undo' | 'history.redo'>('history.undo');
    const el = mountWith(editor, () => {
      const cmd = useEditorCommand(slot);
      return () => h('output', String(cmd.isEnabled.value));
    });
    expect(el.textContent).toBe('true');
    slot.value = 'history.redo';
    await nextTick();
    await nextTick();
    expect(el.textContent).toBe('false');
  });

  test('useEditorCommand reports the engine reason when disabled, and runs when enabled', async () => {
    const empty = mountWith(null, () => {
      const cmd = useEditorCommand('history.undo');
      return () => h('output', `${cmd.isEnabled.value}|${cmd.disabledReason.value}`);
    });
    expect(empty.textContent).toBe('false|disabled.noEditor');
    app?.unmount();
    host?.remove();

    const editor = await open();
    expect(editor.exec({ type: 'addSlide' }).ok).toBe(true);
    let run: (() => void) | null = null;
    const el = mountWith(editor, () => {
      const cmd = useEditorCommand('history.undo');
      run = () => cmd.execute();
      return () => h('output', String(cmd.isEnabled.value));
    });
    expect(el.textContent).toBe('true');
    run!();
    await nextTick();
    expect(editor.snapshot().slides.length).toBe(1);
  });
});

describe('PptxEditorToolbar', () => {
  const groups = (el: HTMLElement): string[] =>
    [...el.querySelectorAll<HTMLElement>('[role="toolbar"] > [role="group"]')].map(
      (g) => g.dataset.group ?? ''
    );

  test('the preset draws the groups for the selection, in order', async () => {
    const editor = await open();
    const el = mountWith(editor, () => () => h(PptxEditorToolbar));
    // Nothing selected: Insert, the slide, and Find, as in React.
    expect(groups(el)).toEqual(toolbarGroupsFor('none').map((g) => g.id));
    selectBox(editor);
    await nextTick();
    // The text box: its text, shape, and arrange groups.
    expect(groups(el)).toEqual(toolbarGroupsFor('shape').map((g) => g.id));
    const bold = el.querySelector<HTMLButtonElement>('[data-slot="text.bold"]')!;
    expect(bold.getAttribute('aria-pressed')).toBe('false');
    expect(bold.getAttribute('aria-label')!.startsWith('Bold')).toBe(true);
  });

  test('a disabled control carries the engine reason in its tooltip', async () => {
    const el = mountWith(null, () => () => h(PptxEditorToolbar));
    const undo = el.querySelector<HTMLButtonElement>('[data-slot="history.undo"]')!;
    expect(undo.getAttribute('aria-disabled')).toBe('true');
    expect(undo.hasAttribute('data-disabled')).toBe(true);
    expect(
      document.getElementById(undo.getAttribute('aria-describedby')!)?.textContent
    ).toBeTruthy();
    // The tooltip names the control, with the reason on its own line.
    expect(undo.getAttribute('data-pptx-tooltip')).toBe(undo.getAttribute('aria-label'));
    expect(undo.getAttribute('data-pptx-tooltip-reason')).toBeTruthy();
  });

  test('ToolbarGroup hides or adds groups; ToolbarSlot hides one control', async () => {
    const editor = await open();
    selectBox(editor);
    const el = mountWith(
      editor,
      () => () =>
        h(PptxEditorToolbar, null, () => [
          h(ToolbarGroup, { id: 'insert', hidden: true }),
          h(ToolbarSlot, { slotId: 'text.italic', hidden: true }),
          h(ToolbarGroup, { id: 'acme', label: 'Acme' }, () => h('button', { id: 'ask' }, 'Ask')),
        ])
    );
    expect(groups(el)).not.toContain('insert');
    expect(groups(el).at(-1)).toBe('acme');
    expect(el.querySelector('[data-group="acme"]')!.getAttribute('aria-label')).toBe('Acme');
    expect(el.querySelector('#ask') !== null).toBe(true);
    expect(el.querySelector('[data-slot="text.italic"]') === null).toBe(true);
    expect(el.querySelector('[data-slot="text.bold"]') !== null).toBe(true);
  });

  test('preset false renders only the children', async () => {
    const el = mountWith(
      await open(),
      () => () =>
        h(PptxEditorToolbar, { preset: false }, () => [h(ToolbarSlot, { slotId: 'history.undo' })])
    );
    expect(el.querySelectorAll('[data-slot]').length).toBe(1);
    expect(el.querySelector('[data-slot="history.undo"]') !== null).toBe(true);
  });

  test('an unknown slot id warns once and renders nothing', async () => {
    const warn = spyOn(console, 'warn').mockImplementation(() => undefined);
    try {
      const el = mountWith(
        await open(),
        () => () =>
          h(PptxEditorToolbar, { preset: false }, () => [
            h(ToolbarSlot, { slotId: 'text.nope' as 'text.bold' }),
          ])
      );
      expect(el.querySelector('[data-slot]') === null).toBe(true);
      expect(warn.mock.calls.some(([m]) => String(m).includes('"text.nope"'))).toBe(true);
    } finally {
      warn.mockRestore();
    }
  });

  test('a picker slot (the table grid) draws its control, as in React', async () => {
    const el = mountWith(
      await open(),
      () => () =>
        h(PptxEditorToolbar, { preset: false }, () => [h(ToolbarSlot, { slotId: 'table.insert' })])
    );
    expect(el.querySelector('[data-slot="table.insert"]') !== null).toBe(true);
  });
});
