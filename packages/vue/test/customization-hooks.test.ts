// The composables a host builds its own chrome with: `useZoom`, `useEditorEvent`,
// `useEditorSnapshot`, and `useScopeClassName`. The twin of the React test
// (packages/react/src/__tests__/customization-hooks.test.tsx).

import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { createApp, defineComponent, h, nextTick, shallowRef, type App } from 'vue';
import {
  createPptxEditor,
  resetDevWarnings,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx } from '../../../scripts/fixtures/builder.ts';
import { providePptxEditor } from '../src/composables.ts';
import {
  PptxEditorRoot,
  PptxEditorViewport,
  useEditorEvent,
  useEditorSnapshot,
  useScopeClassName,
  useZoom,
  usePptxEditor,
  type UseZoomReturn,
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

function mount(render: () => ReturnType<typeof h>): App {
  const host = document.createElement('div');
  document.body.append(host);
  const app = createApp({ render });
  app.mount(host);
  apps.push([app, host]);
  return app;
}

async function editorWithDeck(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true });
  editors.push(editor);
  await editor.load(buildPptx({ slides: [{ shapes: '' }, { shapes: '' }] }));
  return editor;
}

/** Render `child` under a host component that provides `editor`. */
const Provide = defineComponent({
  props: { editor: { type: Object, required: true }, child: { type: Object, required: true } },
  setup(props) {
    providePptxEditor(shallowRef(props.editor as PptxEditorInstance));
    return () => h(props.child as ReturnType<typeof defineComponent>);
  },
});

describe('useZoom (Vue)', () => {
  test('reads the zoom and sets it, steps it, fits, and resets', async () => {
    const editor = await editorWithDeck();
    let z!: UseZoomReturn;
    const Probe = defineComponent({
      setup() {
        z = useZoom();
        return () => null;
      },
    });
    mount(() => h(Provide, { editor, child: Probe }));
    expect(z.isFit.value).toBe(true);
    z.reset();
    await nextTick();
    expect([z.zoom.value, z.isFit.value]).toEqual([1, false]);
    z.zoomIn();
    await nextTick();
    expect(z.zoom.value).toBe(1.1);
    z.zoomOut();
    z.zoomOut();
    await nextTick();
    expect(editor.snapshot().zoom).toBe(0.9);
    z.setZoom(9);
    await nextTick();
    expect([z.zoom.value, z.canZoomIn.value, z.canZoomOut.value]).toEqual([4, false, true]);
    z.fit();
    await nextTick();
    expect(z.isFit.value).toBe(true);
    expect(z.levels).toContain(1);
  });

  test('outside a root it reports 100 % and its actions do nothing', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => undefined);
    let z!: UseZoomReturn;
    mount(() =>
      h(
        defineComponent({
          setup() {
            z = useZoom();
            return () => null;
          },
        })
      )
    );
    await nextTick();
    expect([z.zoom.value, z.canZoomIn.value]).toEqual([1, false]);
    z.reset();
    expect(z.zoom.value).toBe(1);
    warn.mockRestore();
  });
});

describe('useEditorEvent (Vue)', () => {
  test('calls its handler for its event and stops on unmount', async () => {
    const editor = await editorWithDeck();
    const log: string[] = [];
    const Listener = defineComponent({
      setup() {
        useEditorEvent('selectionChange', (s) => void log.push(`select:${s.currentSlide}`));
        useEditorEvent('rejection', (r) => void log.push(`rejection:${r.reason}`));
        return () => null;
      },
    });
    const app = mount(() => h(Provide, { editor, child: Listener }));
    await nextTick();
    editor.selectSlide(1);
    editor.exec({ type: 'undo' });
    expect(log).toEqual(['select:1', 'rejection:disabled.nothingToUndo']);
    app.unmount();
    apps.splice(0);
    editor.selectSlide(0);
    expect(log.length).toBe(2);
  });
});

describe('useEditorSnapshot (Vue)', () => {
  test('changes its counter on each editor change', async () => {
    const editor = await editorWithDeck();
    let revision!: { value: number };
    const Probe = defineComponent({
      setup() {
        revision = useEditorSnapshot(usePptxEditor());
        return () => null;
      },
    });
    mount(() => h(Provide, { editor, child: Probe }));
    await nextTick();
    const before = revision.value;
    editor.selectSlide(1);
    expect(revision.value > before).toBe(true);
  });
});

describe('useScopeClassName (Vue)', () => {
  const Part = defineComponent({
    setup() {
      const scope = useScopeClassName();
      return () => h('aside', { 'data-testid': 'part', class: `${scope.value}my-panel` });
    },
  });
  const part = (): string => document.querySelector('[data-testid="part"]')!.className;

  test('outside the viewport it is the scope class, with dark for a dark root', async () => {
    mount(() => h(PptxEditorRoot, null, { default: () => h(Part) }));
    await nextTick();
    expect(part()).toBe('pptx-editor my-panel');
    apps.splice(0).forEach(([app, host]) => (app.unmount(), host.remove()));
    mount(() => h(PptxEditorRoot, { colorMode: 'dark' }, { default: () => h(Part) }));
    await nextTick();
    expect(part()).toBe('pptx-editor dark my-panel');
  });

  test('inside the viewport it is empty: the viewport carries the scope', async () => {
    mount(() =>
      h(
        PptxEditorRoot,
        { colorMode: 'dark' },
        { default: () => h(PptxEditorViewport, null, { default: () => h(Part) }) }
      )
    );
    await nextTick();
    expect(part()).toBe('my-panel');
    expect(document.querySelector('[data-testid="part"]')!.closest('.pptx-editor.dark')).not.toBe(
      null
    );
  });
});
