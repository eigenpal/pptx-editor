// The Vue twin of `Pro effect controls` in packages/react/src/__tests__/shape-format.test.tsx:
// without the drawing module's effects renderer, Shape Effects offers only the outer shadows and
// Picture Styles leaves out the reflection and soft edge styles. A module installed later shows
// them the next time the menu opens.

import { afterAll, afterEach, beforeAll, describe, expect, spyOn, test } from 'bun:test';
import { createApp, defineComponent, h, nextTick, shallowRef, type App, type Component } from 'vue';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { installDrawingRenderers, type DrawingEffectsRenderer } from '@pptx-editor.dev/core/layout';
import { buildPptx, shape, solid } from '../../../scripts/fixtures/builder.ts';
import { providePptxEditor } from '../src/composables.ts';
import { PptxEditorMenu, PptxEditorToolbar, ToolbarSlot } from '../src/index.ts';

const IN = 914_400;
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
  ),
  (c) => c.charCodeAt(0)
);

/** An effects renderer that paints nothing: the chrome reads only that one is installed. */
const EFFECTS: DrawingEffectsRenderer = {
  body: () => null,
  glow: () => null,
  reflection: () => null,
};

// happy-dom lays nothing out: give the bar room, so no control moves into More.
const width = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth');
let warn: ReturnType<typeof spyOn>;
beforeAll(() => {
  Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
    configurable: true,
    get: () => 100_000,
  });
  // A late install logs a warning once, which these tests do not need to see.
  warn = spyOn(console, 'warn').mockImplementation(() => {});
});
afterAll(() => {
  if (width) Object.defineProperty(HTMLElement.prototype, 'clientWidth', width);
  warn.mockRestore();
});

let app: App | null = null;
let host: HTMLElement | null = null;
function unmount(): void {
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
  document.body.replaceChildren();
}
afterEach(() => {
  unmount();
  installDrawingRenderers({ effects: undefined });
});

async function mount(editor: PptxEditorInstance, render: () => unknown): Promise<void> {
  unmount();
  host = document.createElement('div');
  document.body.append(host);
  app = createApp(
    defineComponent({
      setup() {
        providePptxEditor(shallowRef(editor));
        return render;
      },
    }) as Component
  );
  app.mount(host);
  await nextTick();
}

async function load(shapes: string, media = false): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes,
        ...(media ? { rels: [['rIdImg', IMAGE_REL, '../media/image1.png']] } : {}),
      },
    ],
    ...(media ? { media: [{ name: 'image1.png', bytes: PNG, contentType: 'image/png' }] } : {}),
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  const id = editor.session!.layout(0)!.items[0]!.element.nodeId;
  editor.setSelection({ kind: 'shapes', nodeIds: [id] });
  return editor;
}

const openShape = () =>
  load(shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, fill: solid('4472C4') }));
const openPicture = () =>
  load(
    '<p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture 1"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
      `<p:spPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${IN}" cy="${IN}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`,
    true
  );

/** The labels of the rows the open menus show. */
const rows = (): string[] =>
  [...document.querySelectorAll('[role="menuitem"], [role="menuitemradio"]')].map(
    (r) => r.textContent?.trim() ?? ''
  );
async function click(el: Element | undefined): Promise<void> {
  expect(el).toBeTruthy();
  (el as HTMLElement).click();
  await nextTick();
  await nextTick();
}
const row = (label: string) =>
  [...document.querySelectorAll('[role^="menuitem"]')].find((r) => r.textContent?.trim() === label);
const button = (label: string) =>
  document.querySelector(`button[aria-label="${label}"]`) ?? undefined;

describe('Pro effect controls (Vue)', () => {
  test('without the drawing module, Shape effects offers only the outer shadows', async () => {
    const editor = await openShape();
    await mount(editor, () =>
      h(PptxEditorToolbar, { preset: false }, () => [h(ToolbarSlot, { slotId: 'shape.effects' })])
    );
    await click(button('Shape effects'));
    for (const name of ['Reflection', 'Glow', 'Soft edges']) expect(row(name)).toBeUndefined();
    await click(row('Shadow'));
    expect(rows()).toContain('Offset: top');
    expect(rows().some((r) => r.startsWith('Inside:'))).toBe(false);
    await click(row('Offset: top'));
    expect(editor.snapshot().geometry?.effects?.shadow).toBe('outerTop');

    // A module installed later shows its effects the next time the menu opens.
    installDrawingRenderers({ effects: EFFECTS });
    await click(button('Shape effects'));
    for (const name of ['Shadow', 'Reflection', 'Glow', 'Soft edges'])
      expect(row(name)).toBeTruthy();
  });

  test('without the drawing module, a Shape effects menu row has only the Shadow submenu', async () => {
    const editor = await openShape();
    await mount(editor, () =>
      h(PptxEditorMenu, null, () => [
        h(PptxEditorMenu.Entry, { id: 'acme', label: 'Acme' }, () => [
          h(PptxEditorMenu.Item, { slotId: 'shape.effects' }),
        ]),
      ])
    );
    await click(
      [...document.querySelectorAll('.pptx-menubar__trigger')].find((b) => b.textContent === 'Acme')
    );
    await click(row('Shape effects'));
    expect(row('Shadow')).toBeTruthy();
    for (const name of ['Reflection', 'Glow', 'Soft edges']) expect(row(name)).toBeUndefined();
    await click(row('Shadow'));
    expect(rows().some((r) => r.startsWith('Inside:'))).toBe(false);
  });

  test('Picture styles leaves out the reflection and soft edge styles without the module', async () => {
    const styles = async (): Promise<string[]> => {
      const editor = await openPicture();
      await mount(editor, () =>
        h(PptxEditorToolbar, { preset: false }, () => [h(ToolbarSlot, { slotId: 'image.styles' })])
      );
      await click(button('Picture styles'));
      return rows();
    };
    expect(await styles()).toEqual(['Simple frame, white', 'Simple frame, black']);
    installDrawingRenderers({ effects: EFFECTS });
    expect(await styles()).toEqual([
      'Simple frame, white',
      'Reflected rounded rectangle',
      'Soft edge rectangle',
      'Simple frame, black',
      'Soft edge oval',
    ]);
  });
});
