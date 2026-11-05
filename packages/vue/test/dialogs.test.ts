import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { resetDevWarnings } from '@pptx-editor.dev/core/editor';
import { createApp, defineComponent, h, nextTick, type App, type Component } from 'vue';
import {
  PageSetupDialog,
  PptxEditorRoot,
  definePopup,
  useEditorDialog,
  usePageSetupDialog,
  type PptxEditorPopups,
} from '../src/index.ts';

// The Vue twin of the dialog parts and the popups registry
// (packages/react/src/__tests__/dialog-parts.test.tsx, customization.test.tsx).

let app: App | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
});

async function mount(render: () => unknown): Promise<HTMLElement> {
  // A second mount in one test replaces the first: no app stays in the document for a later
  // file's `document.querySelector` to find.
  app?.unmount();
  host?.remove();
  host = document.createElement('div');
  document.body.append(host);
  app = createApp(defineComponent({ setup: () => render }) as Component);
  app.mount(host);
  await nextTick();
  return host;
}

const part = (name: string) => document.querySelector<HTMLElement>(`[data-pptx-part="${name}"]`);

describe('PageSetupDialog parts', () => {
  test('the default arrangement marks every part, and a part keeps its wiring', async () => {
    let closed = 0;
    await mount(() =>
      h(
        PageSetupDialog,
        { open: true, onClose: () => (closed += 1), onThemeChange: () => {} },
        () => [h(PageSetupDialog.Cancel, { className: 'acme-cancel', onClick: () => {} })]
      )
    );
    for (const name of ['header', 'title', 'body', 'footer', 'error', 'apply', 'cancel'])
      expect(part(name) !== null).toBe(true);
    const cancel = part('cancel')!;
    expect(cancel.classList.contains('pptx-button')).toBe(true);
    expect(cancel.classList.contains('acme-cancel')).toBe(true);
    cancel.click();
    expect(closed).toBe(1);
  });

  test('hidden removes a field, and slot content replaces a part', async () => {
    await mount(() =>
      h(PageSetupDialog, { open: true, onClose: () => {} }, () => [
        h(PageSetupDialog.Field, { name: 'substitutes', hidden: true }),
        h(PageSetupDialog.Title, null, () => 'Slide size'),
      ])
    );
    expect(document.querySelector('[data-pptx-field="substitutes"]') === null).toBe(true);
    expect(part('title')!.textContent).toBe('Slide size');
  });

  test('preset false renders the host arrangement over the same draft', async () => {
    const themes: string[] = [];
    const Light = defineComponent({
      setup() {
        const draft = usePageSetupDialog();
        return () =>
          h(
            'button',
            { 'data-testid': 'light', onClick: () => draft.setValue('theme', 'light') },
            draft.values.value.theme
          );
      },
    });
    await mount(() =>
      h(
        PageSetupDialog,
        {
          open: true,
          preset: false,
          theme: 'dark',
          onClose: () => {},
          onThemeChange: (t: string) => themes.push(t),
        },
        () => [h(Light), h(PageSetupDialog.Apply, null, () => 'Save')]
      )
    );
    expect(part('body') === null).toBe(true);
    const light = document.querySelector<HTMLElement>('[data-testid="light"]')!;
    expect(light.textContent).toBe('dark');
    light.click();
    await nextTick();
    expect(light.textContent).toBe('light');
    expect(themes).toEqual([]);
    part('apply')!.click();
    expect(themes).toEqual(['light']);
  });
});

describe('popups registry', () => {
  const Opener = defineComponent({
    props: {
      dialog: { type: String as () => 'pageSetup' | 'background' | 'paragraph', required: true },
    },
    setup(props) {
      const handle = useEditorDialog(props.dialog);
      return () =>
        h('button', {
          'data-testid': 'opener',
          'data-available': String(handle.isAvailable.value),
          'data-open': String(handle.isOpen.value),
          onClick: () => handle.open(),
        });
    },
  });
  const opener = () => document.querySelector<HTMLElement>('[data-testid="opener"]')!;
  const dialogEl = (name: string) => document.querySelector(`[data-pptx-dialog="${name}"]`);
  const root = (dialog: 'pageSetup' | 'background' | 'paragraph', popups?: PptxEditorPopups) =>
    mount(() => h(PptxEditorRoot, popups ? { popups } : {}, () => [h(Opener, { dialog })]));

  test('without an entry, the packaged Page setup dialog opens and closes', async () => {
    await root('pageSetup');
    expect(dialogEl('pageSetup') === null).toBe(true);
    opener().click();
    await nextTick();
    expect(dialogEl('pageSetup') !== null).toBe(true);
    expect(opener().getAttribute('data-open')).toBe('true');
    part('cancel')!.click();
    await nextTick();
    expect(dialogEl('pageSetup') === null).toBe(true);
  });

  test('a definePopup entry replaces the packaged dialog', async () => {
    const Mine = defineComponent({
      props: { open: Boolean, onClose: Function },
      setup: (props) => () =>
        h('div', { 'data-testid': 'mine', onClick: () => (props.onClose as () => void)() }),
    });
    await root('pageSetup', { pageSetup: definePopup(Mine) });
    opener().click();
    await nextTick();
    expect(dialogEl('pageSetup') === null).toBe(true);
    document.querySelector<HTMLElement>('[data-testid="mine"]')!.click();
    await nextTick();
    expect(document.querySelector('[data-testid="mine"]') === null).toBe(true);
  });

  test('false turns a dialog off, and the packaged Background dialog opens without one', async () => {
    await root('pageSetup', { pageSetup: false });
    expect(opener().getAttribute('data-available')).toBe('false');
    app?.unmount();
    host?.remove();
    await root('background');
    expect(opener().getAttribute('data-available')).toBe('true');
    opener().click();
    await nextTick();
    expect(dialogEl('background') !== null).toBe(true);
  });

  test('the packaged Paragraph dialog opens with no warning; false keeps it closed', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    await root('paragraph');
    opener().click();
    await nextTick();
    expect(dialogEl('paragraph') !== null).toBe(true);
    app?.unmount();
    host?.remove();
    await root('paragraph', { paragraph: false });
    opener().click();
    await nextTick();
    expect(dialogEl('paragraph') === null).toBe(true);
    expect(warn).not.toHaveBeenCalled();
    warn.mockRestore();
  });
});
