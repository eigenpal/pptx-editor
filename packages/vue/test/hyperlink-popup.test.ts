import { afterEach, describe, expect, test } from 'bun:test';
import { createApp, defineComponent, h, nextTick, type App, type Component } from 'vue';
import {
  requestLinkPopup,
  surfaceRootOf,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import {
  definePopup,
  PptxEditorContent,
  PptxEditorHyperLink,
  PptxEditorRoot,
  PptxEditorViewport,
  useHyperlinkPopup,
  type PptxEditorPopups,
} from '../src/index.ts';

const IN = 914_400;

let app: App | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
  document.body.replaceChildren();
});

async function mountRoot(popups?: PptxEditorPopups): Promise<PptxEditorInstance> {
  host = document.createElement('div');
  document.body.append(host);
  let ready: PptxEditorInstance | null = null;
  app = createApp(
    defineComponent({
      setup: () => () =>
        h(PptxEditorRoot, { popups, onReady: (e: PptxEditorInstance) => (ready = e) }, () =>
          h(PptxEditorViewport, null, () => h(PptxEditorContent))
        ),
    }) as Component
  );
  app.mount(host);
  await nextTick();
  const editor = ready as PptxEditorInstance | null;
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({ id: 2, x: IN, y: IN, cx: 3 * IN, cy: IN, text: para(run('Hello')) }),
      },
    ],
  });
  expect((await editor!.load(bytes)).ok).toBe(true);
  await nextTick();
  return editor!;
}

const select = (editor: PptxEditorInstance, from: number, to: number): void => {
  const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
  editor.setSelection({
    kind: 'text',
    nodeId,
    anchor: { paragraph: 0, offset: from },
    focus: { paragraph: 0, offset: to },
  });
};

async function linkAndCaret(editor: PptxEditorInstance): Promise<void> {
  select(editor, 0, 5);
  editor.exec({ type: 'setLink', url: 'https://example.com/a' });
  select(editor, 2, 2);
  await nextTick();
}

async function clickSlide(editor: PptxEditorInstance): Promise<void> {
  const root = surfaceRootOf(editor)!;
  const at = { button: 0, clientX: 5, clientY: 5, bubbles: true };
  root.dispatchEvent(new PointerEvent('pointerdown', at));
  root.dispatchEvent(new PointerEvent('pointerup', at));
  await new Promise((r) => setTimeout(r, 40));
  await nextTick();
}

/**
 * A plain click between the "e" and the "l" of "Hello", through the surface's own pointer path,
 * as React's test: the stage gets a 1280 px box, and the point comes from the caret stops.
 */
async function clickText(editor: PptxEditorInstance): Promise<void> {
  const root = surfaceRootOf(editor)!;
  const stage = root.querySelector<HTMLElement>('.pptx-surface__stage')!;
  const W = 1280;
  stage.getBoundingClientRect = () =>
    ({ left: 0, top: 0, right: W, bottom: W, width: W, height: W, x: 0, y: 0 }) as DOMRect;
  const layout = editor.session!.layout(0)!;
  const item = layout.items[0]!;
  if (item.kind !== 'shape' || !item.text || !item.textMatrix) throw new Error('no text');
  const line = item.text.lines[0]!;
  const fx = item.text.box.x + (line.caretX[2]! + line.caretX[3]!) / 2;
  const fy = item.text.box.y + line.top + line.height / 2;
  const [a, b, c, d, e, f] = item.textMatrix;
  const px = (v: number) => (v / layout.slide.width) * W;
  const at = {
    button: 0,
    clientX: px(a * fx + c * fy + e),
    clientY: px(b * fx + d * fy + f),
    bubbles: true,
  };
  stage.dispatchEvent(new PointerEvent('pointerdown', at));
  stage.dispatchEvent(new PointerEvent('pointerup', at));
  window.dispatchEvent(new PointerEvent('pointerup', at));
  await new Promise((r) => setTimeout(r, 40));
  await nextTick();
}

const panel = () => document.querySelector<HTMLElement>('.pptx-link');
const button = (label: string) =>
  document.querySelector<HTMLButtonElement>(`.pptx-link [aria-label="${label}"]`);

describe('PptxEditorHyperLink', () => {
  test('a click into linked text opens reading mode; Edit and Apply change the link', async () => {
    const editor = await mountRoot();
    await linkAndCaret(editor);
    await clickSlide(editor);
    expect(panel()?.getAttribute('data-mode')).toBe('reading');
    const url = document.querySelector<HTMLAnchorElement>('.pptx-link__url')!;
    expect(url.getAttribute('href')).toBe('https://example.com/a');
    expect(url.getAttribute('rel')).toBe('noopener noreferrer');
    expect(button('Copy link') !== null).toBe(true);
    expect(document.querySelector('.pptx-link-popup')!.contains(document.activeElement)).toBe(
      false
    );
    button('Edit link')!.click();
    await nextTick();
    expect(panel()?.getAttribute('data-mode')).toBe('editing');
    const field = document.querySelector<HTMLInputElement>('input[aria-label="Link"]')!;
    expect(field.value).toBe('https://example.com/a');
    field.value = 'example.org';
    field.dispatchEvent(new Event('input'));
    await nextTick();
    field.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })
    );
    await nextTick();
    expect(panel() === null).toBe(true);
    expect(editor.snapshot().format?.link).toBe('https://example.org');
  });

  test('Unlink removes the link; a caret move closes reading mode', async () => {
    const editor = await mountRoot();
    await linkAndCaret(editor);
    await clickSlide(editor);
    select(editor, 3, 3);
    await nextTick();
    expect(panel() === null).toBe(true);
    await clickSlide(editor);
    button('Remove link')!.click();
    await nextTick();
    expect(panel() === null).toBe(true);
    expect(editor.snapshot().format?.link ?? '').toBe('');
  });

  test('a popups.hyperlink renderer overrides parts in place', async () => {
    // A renderer that shows reading mode asks for it.
    const AcmeLink = defineComponent({
      setup: () => () =>
        h(PptxEditorHyperLink, { className: 'acme-link' }, () => [
          h(PptxEditorHyperLink.Unlink, { hidden: true }),
          h(PptxEditorHyperLink.Edit, { asChild: true }, () => [
            h('a', { class: 'acme-edit' }, 'Change'),
          ]),
        ]),
    });
    const editor = await mountRoot({ hyperlink: definePopup(AcmeLink, { reading: true }) });
    await linkAndCaret(editor);
    await clickSlide(editor);
    expect(document.querySelector('.pptx-link-popup.acme-link') !== null).toBe(true);
    expect(button('Remove link') === null).toBe(true);
    expect(document.querySelectorAll('.pptx-link__copy').length).toBe(1);
    const edit = document.querySelector<HTMLElement>('a.acme-edit')!;
    expect(edit.getAttribute('aria-label')).toBe('Edit link');
    edit.click();
    await nextTick();
    expect(panel()?.getAttribute('data-mode')).toBe('editing');
  });

  test('preset: false and useHyperlinkPopup drive a host panel', async () => {
    const Custom = defineComponent({
      setup() {
        const popup = useHyperlinkPopup();
        return () =>
          h(PptxEditorHyperLink, { preset: false }, () => [
            h('p', { class: 'acme-mode' }, popup.state.value.mode),
            h('button', { class: 'acme-fill', onClick: () => popup.setUrl('https://acme.test/') }),
            h('button', { class: 'acme-save', onClick: () => popup.commitEdit() }),
          ]);
      },
    });
    const editor = await mountRoot({ hyperlink: () => h(Custom) });
    select(editor, 0, 5);
    await nextTick();
    requestLinkPopup(editor);
    await nextTick();
    expect(document.querySelector('.acme-mode')?.textContent).toBe('editing');
    expect(document.querySelector('.pptx-link__apply') === null).toBe(true);
    document.querySelector<HTMLButtonElement>('.acme-fill')!.click();
    document.querySelector<HTMLButtonElement>('.acme-save')!.click();
    await nextTick();
    expect(editor.snapshot().format?.link).toBe('https://acme.test/');
    expect(panel() === null).toBe(true);
  });

  test('a viewer reads a link; Edit and Unlink are absent', async () => {
    const editor = await mountRoot();
    await linkAndCaret(editor);
    editor.setMode('view');
    editor.setSelection({ kind: 'none' });
    await nextTick();
    // A viewer's own click on the linked word places the caret there and opens the panel.
    await clickText(editor);
    expect(editor.snapshot().selection.kind).toBe('text');
    expect(panel()?.getAttribute('data-mode')).toBe('reading');
    expect(document.querySelector('.pptx-link__copy') !== null).toBe(true);
    expect(button('Edit link') === null).toBe(true);
    expect(button('Remove link') === null).toBe(true);
  });

  test('a viewer click on text with no link selects the shape and opens nothing', async () => {
    const editor = await mountRoot();
    editor.setMode('view');
    await nextTick();
    await clickText(editor);
    expect(editor.snapshot().selection.kind).toBe('shapes');
    expect(panel()).toBe(null);
  });

  test('a plain popups.hyperlink renderer opens for Ctrl/Cmd+K, not on a click', async () => {
    const editor = await mountRoot({ hyperlink: () => h('p', { class: 'acme-form' }) });
    await linkAndCaret(editor);
    await clickSlide(editor);
    expect(document.querySelector('.acme-form')).toBe(null);
    requestLinkPopup(editor);
    await nextTick();
    expect(document.querySelector('.acme-form') !== null).toBe(true);
  });

  test('F6 moves focus into the reading panel, and a live region announces it', async () => {
    const editor = await mountRoot();
    await linkAndCaret(editor);
    await clickSlide(editor);
    const status = document.querySelector('.pptx-sr-only[role="status"]');
    expect(status?.textContent).toBe('Link: https://example.com/a. Press F6 for its actions.');
    const input = surfaceRootOf(editor)!.querySelector<HTMLElement>('.pptx-surface__input')!;
    input.focus();
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'F6', bubbles: true }));
    await nextTick();
    expect(document.activeElement?.closest('.pptx-link-popup') !== null).toBe(true);
  });
});
