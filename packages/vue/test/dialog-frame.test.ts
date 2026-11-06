import { afterEach, beforeEach, describe, expect, spyOn, test } from 'bun:test';
import { createApp, defineComponent, h, nextTick, shallowRef, type App, type Component } from 'vue';
import { resetDevWarnings } from '@pptx-editor.dev/core/editor';
import { buildPptx } from '../../../scripts/fixtures/builder.ts';
import {
  PageSetupDialog,
  PptxEditorRoot,
  useEditorDialog,
  usePageSetupDialog,
} from '../src/index.ts';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';

// The Vue twin of packages/react/src/__tests__/dialog-frame.test.tsx: the frame's keys, focus,
// and backdrop, text in the default arrangement, and the warnings for misused parts.

let app: App | null = null;
let host: HTMLElement | null = null;
let warn: ReturnType<typeof spyOn>;
beforeEach(() => {
  resetDevWarnings();
  warn = spyOn(console, 'warn').mockImplementation(() => {});
});
afterEach(() => {
  warn.mockRestore();
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
  // Vue skips a listener for an event stamped in the same millisecond as the listener was
  // attached (runtime-dom's `invoker.attached`), so a key sent at once could miss the dialog.
  const mounted = Date.now();
  while (Date.now() <= mounted) await new Promise((resolve) => setTimeout(resolve, 1));
  return host;
}

const part = (name: string) => document.querySelector<HTMLElement>(`[data-pptx-part="${name}"]`);
const frame = () => document.querySelector<HTMLDialogElement>('[data-pptx-dialog="pageSetup"]')!;
const key = (target: Element, init: KeyboardEventInit) =>
  target.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init }));
const mouse = (target: Element, type: string, x: number, y: number) =>
  target.dispatchEvent(new MouseEvent(type, { bubbles: true, clientX: x, clientY: y }));

describe('Vue DialogFrame', () => {
  test('text beside elements stays in the default arrangement', async () => {
    await mount(() => h(PageSetupDialog, { open: true, onClose: () => {} }));
    // The Width label: its text, then the unit for screen readers, then the box.
    const row = document
      .querySelector('[data-pptx-field="slideSize"] input[name="width"]')!
      .closest('label')!;
    const unit = row.querySelector('.pptx-visually-hidden')!;
    expect(row.textContent!.length).toBeGreaterThan(unit.textContent!.length);
  });

  test('Escape closes, Tab wraps, and the dialog starts on OK', async () => {
    let closed = 0;
    await mount(() => h(PageSetupDialog, { open: true, onClose: () => (closed += 1) }));
    expect(document.activeElement).toBe(part('apply'));
    // Tab wraps to the header's close button, the first control, as in the other dialogs.
    key(part('apply')!, { key: 'Tab' });
    const close = document.activeElement as HTMLElement;
    expect(close.getAttribute('aria-label')).toBe('Close');
    expect(part('header')!.contains(close)).toBe(true);
    key(frame(), { key: 'Escape' });
    expect(closed).toBe(1);
    close.click();
    expect(closed).toBe(2);
  });

  test('only a press and a release on the backdrop close it', async () => {
    let closed = 0;
    await mount(() => h(PageSetupDialog, { open: true, onClose: () => (closed += 1) }));
    mouse(part('title')!, 'mousedown', 5, 5);
    mouse(frame(), 'click', -10, -10);
    expect(closed).toBe(0);
    mouse(frame(), 'mousedown', -10, -10);
    mouse(frame(), 'click', -10, -10);
    expect(closed).toBe(1);
  });

  test('arrow keys move the appearance choice', async () => {
    await mount(() =>
      h(PageSetupDialog, { open: true, onClose: () => {}, theme: 'light', onThemeChange: () => {} })
    );
    const light = document.querySelector<HTMLElement>('[role="radio"][data-value="light"]')!;
    const dark = document.querySelector<HTMLElement>('[role="radio"][data-value="dark"]')!;
    expect([light.tabIndex, dark.tabIndex]).toEqual([0, -1]);
    key(light, { key: 'ArrowRight' });
    await nextTick();
    expect(dark.getAttribute('aria-checked')).toBe('true');
    expect(document.activeElement).toBe(dark);
  });

  test('focus goes back to the element open() names', async () => {
    const Opener = defineComponent({
      setup() {
        const handle = useEditorDialog('pageSetup');
        return () => [
          h('button', { 'data-testid': 'target' }),
          h('button', {
            'data-testid': 'opener',
            onClick: () =>
              handle.open(document.querySelector<HTMLElement>('[data-testid="target"]')),
          }),
        ];
      },
    });
    await mount(() => h(PptxEditorRoot, null, () => [h(Opener)]));
    const opener = document.querySelector<HTMLElement>('[data-testid="opener"]')!;
    opener.focus();
    opener.click();
    await nextTick();
    part('cancel')!.click();
    await nextTick();
    await nextTick();
    expect(document.activeElement).toBe(document.querySelector('[data-testid="target"]'));
  });

  test('asChild puts the wiring on the host element; className reaches the surface', async () => {
    let closed = 0;
    await mount(() =>
      h(
        PageSetupDialog,
        { open: true, onClose: () => (closed += 1), className: 'acme-dialog' },
        () => [
          h(PageSetupDialog.Cancel, { asChild: true, className: 'acme-cancel' }, () => [
            h('a', { href: '#close' }, 'Close'),
          ]),
        ]
      )
    );
    expect(frame().classList.contains('acme-dialog')).toBe(true);
    const cancel = part('cancel')!;
    expect(cancel.localName).toBe('a');
    expect(cancel.classList.contains('acme-cancel')).toBe(true);
    expect(cancel.classList.contains('pptx-button')).toBe(false);
    cancel.click();
    expect(closed).toBe(1);
  });

  test('Cancel discards the appearance draft', async () => {
    const themes: string[] = [];
    await mount(() =>
      h(PageSetupDialog, {
        open: true,
        onClose: () => {},
        theme: 'light',
        onThemeChange: (t: string) => themes.push(t),
      })
    );
    document.querySelector<HTMLElement>('[role="radio"][data-value="dark"]')!.click();
    await nextTick();
    part('cancel')!.click();
    expect(themes).toEqual([]);
  });

  test('a part outside its dialog throws for its draft hook', async () => {
    const Stray = defineComponent({
      setup() {
        usePageSetupDialog();
        return () => null;
      },
    });
    host = document.createElement('div');
    document.body.append(host);
    app = createApp(Stray as Component);
    let error: unknown;
    app.config.errorHandler = (e) => (error = e);
    app.mount(host);
    expect(String(error)).toContain(
      'PageSetupDialog parts must be rendered inside PageSetupDialog.'
    );
  });

  test('Tab skips the unchecked radio with a dark choice', async () => {
    await mount(() =>
      h(PageSetupDialog, { open: true, onClose: () => {}, theme: 'dark', onThemeChange: () => {} })
    );
    const light = document.querySelector<HTMLElement>('[role="radio"][data-value="light"]')!;
    const dark = document.querySelector<HTMLElement>('[role="radio"][data-value="dark"]')!;
    // One Tab stop per radio group: the checked radio only.
    expect([light.tabIndex, dark.tabIndex]).toEqual([-1, 0]);
  });

  test('the registry closes a dialog when the editor opens another document', async () => {
    let editor: PptxEditorInstance | null = null;
    const Opener = defineComponent({
      setup() {
        const handle = useEditorDialog('pageSetup');
        return () => h('button', { 'data-testid': 'opener', onClick: () => handle.open() });
      },
    });
    await mount(() =>
      h(PptxEditorRoot, { onReady: (e: PptxEditorInstance) => (editor = e) }, () => [h(Opener)])
    );
    await editor!.load(buildPptx({ slides: [{ shapes: '' }] }));
    document.querySelector<HTMLElement>('[data-testid="opener"]')!.click();
    await nextTick();
    expect(frame() !== null).toBe(true);
    await editor!.load(buildPptx({ slides: [{ shapes: '' }, { shapes: '' }] }));
    await nextTick();
    expect(document.querySelector('[data-pptx-dialog="pageSetup"]') === null).toBe(true);
  });

  test('Return on a radio applies the draft', async () => {
    const themes: string[] = [];
    let closed = 0;
    await mount(() =>
      h(PageSetupDialog, {
        open: true,
        onClose: () => (closed += 1),
        theme: 'light',
        onThemeChange: (t: string) => themes.push(t),
      })
    );
    const light = document.querySelector<HTMLElement>('[role="radio"][data-value="light"]')!;
    key(light, { key: 'ArrowRight' });
    await nextTick();
    const dark = document.querySelector<HTMLElement>('[role="radio"][data-value="dark"]')!;
    key(dark, { key: 'Enter' });
    expect(themes).toEqual(['dark']);
    expect(closed).toBe(1);
  });

  test('a dialog the host renders closes when the editor opens another document', async () => {
    let editor: PptxEditorInstance | null = null;
    let closed = 0;
    await mount(() =>
      h(PptxEditorRoot, { onReady: (e: PptxEditorInstance) => (editor = e) }, () => [
        h(PageSetupDialog, { open: true, onClose: () => (closed += 1) }),
      ])
    );
    await editor!.load(buildPptx({ slides: [{ shapes: '' }] }));
    await nextTick();
    const before = closed;
    await editor!.load(buildPptx({ slides: [{ shapes: '' }, { shapes: '' }] }));
    await nextTick();
    expect(closed).toBe(before + 1);
  });

  test('a popups entry that turns false closes its open dialog', async () => {
    const popups = shallowRef<{ pageSetup?: false }>({});
    const Opener = defineComponent({
      setup() {
        const handle = useEditorDialog('pageSetup');
        return () => h('button', { 'data-testid': 'opener', onClick: () => handle.open() });
      },
    });
    await mount(() => h(PptxEditorRoot, { popups: popups.value }, () => [h(Opener)]));
    document.querySelector<HTMLElement>('[data-testid="opener"]')!.click();
    await nextTick();
    expect(document.querySelector('[data-pptx-dialog="pageSetup"]') !== null).toBe(true);
    popups.value = { pageSetup: false };
    await nextTick();
    await nextTick();
    expect(document.querySelector('[data-pptx-dialog="pageSetup"]') === null).toBe(true);
  });

  test('misused parts warn, and a bad asChild falls back to the packaged part', async () => {
    await mount(() =>
      h(PageSetupDialog, { open: true, onClose: () => {} }, () => [
        h(PageSetupDialog.Field, { name: 'theme' }),
        h(PageSetupDialog.Title, { className: 'a' }),
        h(PageSetupDialog.Title, { className: 'b' }),
        h('div', null, [h(PageSetupDialog.Cancel)]),
        h(PageSetupDialog.Apply, { asChild: true }, () => 'Save'),
      ])
    );
    const text = warn.mock.calls.map((c: unknown[]) => String(c[0])).join('\n');
    expect(text).toContain('PageSetupDialog.Field name="theme" matches no part');
    expect(text).toContain('PageSetupDialog.Title appears more than once');
    expect(text).toContain('PageSetupDialog.Cancel is inside another element');
    expect(text).toContain('PageSetupDialog.Apply has asChild but not exactly one element child');
    expect(part('apply')!.textContent).toBe('OK');
  });

  test('a misused asChild on the owner does not claim its parts render nothing', async () => {
    await mount(() =>
      h(PageSetupDialog, { open: true, onClose: () => {} }, () => [
        h(PageSetupDialog.Footer, { asChild: true }, () => [h('b'), h('i')]),
        h(PageSetupDialog.Apply, { className: 'host-apply' }),
      ])
    );
    expect(document.querySelector('.host-apply') !== null).toBe(true);
    const text = warn.mock.calls.map((c: unknown[]) => String(c[0])).join('\n');
    expect(text).not.toContain('replaces a default inside');
  });

  test('a part inside a part that brings its own children says it renders nothing', async () => {
    await mount(() =>
      h(PageSetupDialog, { open: true, onClose: () => {} }, () => [
        h(PageSetupDialog.Body, null, () => [h('p')]),
        h(PageSetupDialog.Field, { name: 'slideSize', className: 'x' }),
      ])
    );
    const text = warn.mock.calls.map((c: unknown[]) => String(c[0])).join('\n');
    expect(text).toContain(
      'PageSetupDialog.Field name="slideSize" replaces a default inside PageSetupDialog.Body'
    );
  });
});
