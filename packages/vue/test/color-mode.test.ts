import { afterEach, describe, expect, test } from 'bun:test';
import { createApp, defineComponent, h, nextTick, type App, type Component } from 'vue';
import {
  PageSetupDialog,
  PptxEditorRoot,
  PptxEditorViewport,
  useEditorColorMode,
} from '../src/index.ts';

// The Vue twin of `colorMode` (packages/react/src/__tests__/panes.test.tsx).

let app: App | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
});

async function mountRoot(colorMode?: 'light' | 'dark'): Promise<HTMLElement> {
  // A second mount in one test replaces the first: no app stays in the document for a later
  // file's `document.querySelector` to find.
  app?.unmount();
  host?.remove();
  host = document.createElement('div');
  document.body.append(host);
  app = createApp(
    defineComponent({
      setup: () => () =>
        h(PptxEditorRoot, colorMode ? { colorMode } : null, () => [h(PptxEditorViewport)]),
    }) as Component
  );
  app.mount(host);
  await nextTick();
  return host;
}

describe('colorMode', () => {
  test('dark puts dark on the viewport', async () => {
    const el = await mountRoot('dark');
    expect(el.querySelector('.pptx-editor.dark') !== null).toBe(true);
  });

  test('light is the default', async () => {
    const el = await mountRoot();
    expect(el.querySelector('.pptx-editor')!.classList.contains('dark')).toBe(false);
  });
});

describe('colorMode, review round 1', () => {
  test('system reads the operating system after mount and follows it', async () => {
    const listeners = new Set<() => void>();
    let dark = true;
    const original = globalThis.matchMedia;
    globalThis.matchMedia = ((query: string) => ({
      matches: dark,
      media: query,
      addEventListener: (_: string, l: () => void) => listeners.add(l),
      removeEventListener: (_: string, l: () => void) => listeners.delete(l),
    })) as unknown as typeof matchMedia;
    let mode = '';
    const Probe = defineComponent({
      setup() {
        const { colorMode } = useEditorColorMode();
        return () => ((mode = colorMode.value), null);
      },
    });
    try {
      host = document.createElement('div');
      document.body.append(host);
      app = createApp(
        defineComponent({
          setup: () => () =>
            h(PptxEditorRoot, { colorMode: 'system' }, () => [
              h(PptxEditorViewport, null, () => [h(Probe)]),
            ]),
        }) as Component
      );
      app.mount(host);
      await nextTick();
      expect(host.querySelector('.pptx-editor.dark') !== null).toBe(true);
      expect(mode).toBe('dark');
      dark = false;
      listeners.forEach((l) => l());
      await nextTick();
      expect(host.querySelector('.pptx-editor.dark') === null).toBe(true);
      expect(mode).toBe('light');
      app.unmount();
      app = null;
      expect(listeners.size).toBe(0);
    } finally {
      globalThis.matchMedia = original;
    }
  });

  test('a dialog outside the viewport gets dark beside its own scope', async () => {
    host = document.createElement('div');
    document.body.append(host);
    app = createApp(
      defineComponent({
        setup: () => () =>
          h(PptxEditorRoot, { colorMode: 'dark' }, () => [
            h(PageSetupDialog, { open: true, onClose: () => {} }),
          ]),
      }) as Component
    );
    app.mount(host);
    await nextTick();
    const dialog = document.querySelector('[data-pptx-dialog="pageSetup"]')!;
    expect(dialog.classList.contains('pptx-editor')).toBe(true);
    expect(dialog.classList.contains('dark')).toBe(true);
  });
});
