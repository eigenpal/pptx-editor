// The Vue twin of `Pro effect controls` in packages/react/src/__tests__/shape-format.test.tsx:
// without the drawing module's effects renderer, Shape Effects offers only the outer shadows and
// Picture Styles leaves out the reflection and soft edge styles. A module installed later shows
// them the next time the menu opens. Removing an effect needs no module: the No row of an effect
// stays while the selection has that effect.

import { afterAll, afterEach, beforeAll, describe, expect, spyOn, test } from 'bun:test';
import {
  createApp,
  createSSRApp,
  defineComponent,
  effectScope,
  h,
  nextTick,
  shallowRef,
  type App,
  type Component,
} from 'vue';
import { renderToString } from 'vue/server-renderer';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  drawingRenderersListenerCount,
  installDrawingRenderers,
  type DrawingEffectsRenderer,
} from '@pptx-editor.dev/core/layout';
import { buildPptx, shape, solid } from '../../../scripts/fixtures/builder.ts';
import { providePptxEditor } from '../src/composables.ts';
import {
  PptxEditorMenu,
  PptxEditorToolbar,
  ToolbarSlot,
  useEditorValueCommand,
} from '../src/index.ts';

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

  test('useEditorValueCommand options refresh after a late install', async () => {
    const editor = await openShape();
    let options: { value: readonly string[] } | null = null;
    const Probe = defineComponent({
      setup() {
        const cmd = useEditorValueCommand('shape.effects');
        options = cmd.options;
        return () => h('output', String(cmd.options.value.length));
      },
    });
    await mount(editor, () => h(Probe));
    expect(options!.value.some((v) => v.startsWith('glow.'))).toBe(false);
    expect(options!.value.some((v) => v.startsWith('shadow.inner'))).toBe(false);
    const shown = document.querySelector('output')!.textContent;
    installDrawingRenderers({ effects: EFFECTS });
    await nextTick();
    expect(options!.value).toContain('glow.8.accent2');
    expect(options!.value).toContain('shadow.innerTop');
    expect(options!.value).toContain('reflection.none');
    expect(options!.value).toContain('softEdge.none');
    // The component that reads the options draws again.
    expect(document.querySelector('output')!.textContent).not.toBe(shown);
  });

  test('useEditorValueCommand subscribes on mount and stops on unmount', async () => {
    const editor = await openShape();
    const before = drawingRenderersListenerCount();
    const Probe = defineComponent({
      setup() {
        const cmd = useEditorValueCommand('shape.effects');
        return () => h('output', String(cmd.options.value.length));
      },
    });
    await mount(editor, () => h(Probe));
    expect(drawingRenderersListenerCount()).toBe(before + 1);
    unmount();
    expect(drawingRenderersListenerCount()).toBe(before);
  });

  test('a server render of useEditorValueCommand leaves no renderer subscription', async () => {
    const editor = await openShape();
    const before = drawingRenderersListenerCount();
    const html = await renderToString(
      createSSRApp(
        defineComponent({
          setup() {
            providePptxEditor(shallowRef(editor));
            const cmd = useEditorValueCommand('shape.effects');
            return () => h('output', String(cmd.options.value.length));
          },
        })
      )
    );
    expect(html).toContain('<output>');
    expect(drawingRenderersListenerCount()).toBe(before);
  });

  test('useEditorValueCommand in an effect scope stops with the scope', async () => {
    const before = drawingRenderersListenerCount();
    // An app context gives `inject` its providers outside a component, as a host store does.
    const context = createApp({ render: () => null });
    const scope = effectScope();
    const options = context.runWithContext(() =>
      scope.run(() => useEditorValueCommand('shape.effects').options)!
    );
    expect(drawingRenderersListenerCount()).toBe(before + 1);
    expect(options.value.some((v) => v.startsWith('glow.'))).toBe(false);
    installDrawingRenderers({ effects: EFFECTS });
    expect(options.value).toContain('glow.8.accent2');
    scope.stop();
    expect(drawingRenderersListenerCount()).toBe(before);
  });

  test('an open toolbar shows the Pro picture styles after a late install', async () => {
    const editor = await openPicture();
    await mount(editor, () =>
      h(PptxEditorToolbar, { preset: false }, () => [h(ToolbarSlot, { slotId: 'image.styles' })])
    );
    await click(button('Picture styles'));
    expect(rows()).toEqual(['Simple frame, white', 'Simple frame, black']);
    await click(button('Picture styles'));
    installDrawingRenderers({ effects: EFFECTS });
    await click(button('Picture styles'));
    for (const name of ['Reflected rounded rectangle', 'Soft edge rectangle', 'Soft edge oval'])
      expect(rows()).toContain(name);
  });
});

// Removing an effect needs no drawing module, as in React: without it, the No row of an effect
// stays while the selection has that effect.

const GLOW = '<a:glow rad="63500"><a:srgbClr val="FF0000"/></a:glow>';
const REFLECTION =
  '<a:reflection blurRad="6350" stA="50000" endPos="55000" dist="5000" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>';

/** A shape with a glow (selected), a plain shape, and a shape with a reflection. */
async function openEffects(): Promise<{ editor: PptxEditorInstance; ids: string[] }> {
  const box = (id: number, x: number, effects = ''): string =>
    shape({
      id,
      x,
      y: IN,
      cx: IN,
      cy: IN,
      fill: solid('4472C4'),
      ...(effects ? { effects: `<a:effectLst>${effects}</a:effectLst>` } : {}),
    });
  const editor = await load(box(2, IN, GLOW) + box(3, 3 * IN) + box(4, 5 * IN, REFLECTION));
  return { editor, ids: editor.session!.layout(0)!.items.map((i) => i.element.nodeId) };
}

const kinds = (): string[] =>
  ['Shadow', 'Reflection', 'Glow', 'Soft edges'].filter((name) => row(name) !== undefined);

describe('Removing an effect without the drawing module (Vue)', () => {
  test('Shape effects offers No glow for a shape with a glow, and it removes the glow', async () => {
    const { editor, ids } = await openEffects();
    const [glow, plain] = ids;
    editor.setSelection({ kind: 'shapes', nodeIds: [plain!] });
    await mount(editor, () =>
      h(PptxEditorToolbar, { preset: false }, () => [h(ToolbarSlot, { slotId: 'shape.effects' })])
    );
    await click(button('Shape effects'));
    expect(kinds()).toEqual(['Shadow']);
    // The open menu follows the selection.
    editor.setSelection({ kind: 'shapes', nodeIds: [glow!] });
    await nextTick();
    await nextTick();
    expect(kinds()).toEqual(['Shadow', 'Glow']);
    await click(row('Glow'));
    // Only the row that removes: no glow preset and no Glow Variations grid.
    expect(row('No glow')).toBeTruthy();
    expect(document.querySelector('.pptx-glow-grid')).toBeNull();
    await click(row('No glow'));
    expect(editor.snapshot().geometry?.effects?.glow).toBe(0);
    await click(button('Shape effects'));
    expect(kinds()).toEqual(['Shadow']);
    await click(button('Shape effects'));
    // Undo brings the glow and its row back.
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    await click(button('Shape effects'));
    expect(kinds()).toEqual(['Shadow', 'Glow']);
  });

  test('a menu row offers the No row of each effect that a selected shape has', async () => {
    const { editor, ids } = await openEffects();
    editor.setSelection({ kind: 'shapes', nodeIds: [ids[0]!, ids[2]!] });
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
    expect(kinds()).toEqual(['Shadow', 'Reflection', 'Glow']);
    await click(row('Reflection'));
    expect(rows()).toContain('No reflection');
    expect(rows().some((r) => r.startsWith('Tight reflection'))).toBe(false);
  });

  test('useEditorValueCommand options follow the selection', async () => {
    const { editor, ids } = await openEffects();
    editor.setSelection({ kind: 'shapes', nodeIds: [ids[1]!] });
    let options: { value: readonly string[] } | null = null;
    const Probe = defineComponent({
      setup() {
        const cmd = useEditorValueCommand('shape.effects');
        options = cmd.options;
        return () => h('output', String(cmd.options.value.length));
      },
    });
    await mount(editor, () => h(Probe));
    expect(options!.value).not.toContain('glow.none');
    editor.setSelection({ kind: 'shapes', nodeIds: [ids[0]!, ids[1]!] });
    await nextTick();
    expect(options!.value).toContain('glow.none');
    expect(options!.value).not.toContain('reflection.none');
    expect(options!.value.some((v) => /^glow\.\d/.test(v))).toBe(false);
  });
});
