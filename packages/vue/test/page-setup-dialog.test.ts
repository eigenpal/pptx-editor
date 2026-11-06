import { afterEach, describe, expect, test } from 'bun:test';
import { createApp, defineComponent, h, nextTick, type App, type Component } from 'vue';
import { strFromU8, unzipSync } from 'fflate';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx } from '../../../scripts/fixtures/builder.ts';
import { PageSetupDialog, PptxEditorRoot } from '../src/index.ts';

// The Vue twin of packages/react/src/__tests__/page-setup-dialog.test.tsx: the slide size
// fields of the Page setup dialog.

const BOX =
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="1219200" y="685800"/><a:ext cx="2438400" cy="1371600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>';

let app: App | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
});

/** Mount the dialog in a root with a loaded deck; the dialog opens after the load. */
async function open(onClose: () => void = () => {}): Promise<PptxEditorInstance> {
  let editor: PptxEditorInstance | null = null;
  let shown = false;
  const Shell = defineComponent({
    setup: () => () =>
      h(PptxEditorRoot, { onReady: (e: PptxEditorInstance) => (editor = e) }, () =>
        shown ? [h(PageSetupDialog, { open: true, onClose })] : []
      ),
  });
  host = document.createElement('div');
  document.body.append(host);
  app = createApp(Shell as Component);
  const vm = app.mount(host);
  await editor!.load(buildPptx({ slides: [{ layout: 'blank', shapes: BOX }] }));
  shown = true;
  vm.$forceUpdate();
  await nextTick();
  return editor!;
}

const field = (name: string) => document.querySelector<HTMLInputElement>(`[name="${name}"]`)!;
const button = (label: string) =>
  [...document.querySelectorAll<HTMLButtonElement>('button')].find((b) => b.textContent === label)!;

async function type(name: string, value: string): Promise<void> {
  const el = field(name);
  el.value = value;
  el.dispatchEvent(new Event(el.tagName === 'SELECT' ? 'change' : 'input', { bubbles: true }));
  await nextTick();
}

function boxWidth(editor: PptxEditorInstance): number {
  const slide = strFromU8(unzipSync(editor.save()!)['ppt/slides/slide1.xml']!);
  return Number(/name="Box"[\s\S]*?<a:ext cx="(\d+)"/.exec(slide)![1]);
}

describe('Vue Page setup slide size', () => {
  test('a preset applies its exact size and scales the content, in one undo step', async () => {
    let closed = 0;
    const editor = await open(() => (closed += 1));
    expect(field('size').value).toBe('widescreen');
    expect([field('width').value, field('height').value]).toEqual(['13.33', '7.5']);
    await type('size', 'standard');
    expect([field('width').value, field('height').value]).toEqual(['10', '7.5']);
    button('OK').click();
    await nextTick();
    expect(closed).toBe(1);
    expect([editor.snapshot().slideWidth, editor.snapshot().slideHeight]).toEqual([
      9_144_000, 6_858_000,
    ]);
    expect(boxWidth(editor)).toBe(2_438_400 * 0.75);
    editor.exec({ type: 'undo' });
    expect(editor.snapshot().slideWidth).toBe(12_192_000);
    expect(boxWidth(editor)).toBe(2_438_400);
  });

  test('typed inches give a custom size; portrait swaps them; keep leaves the boxes', async () => {
    const editor = await open();
    await type('width', '12');
    expect(field('size').value).toBe('custom');
    button('Portrait').click();
    await nextTick();
    expect([field('width').value, field('height').value]).toEqual(['7.5', '12']);
    field('scaleContent').click();
    await nextTick();
    button('OK').click();
    await nextTick();
    expect([editor.snapshot().slideWidth, editor.snapshot().slideHeight]).toEqual([
      6_858_000, 10_972_800,
    ]);
    expect(boxWidth(editor)).toBe(2_438_400);
  });

  test('Custom with the same numbers, or orientation there and back, sends no command', async () => {
    const editor = await open();
    await type('size', 'custom');
    button('Portrait').click();
    await nextTick();
    button('Landscape').click();
    await nextTick();
    button('OK').click();
    await nextTick();
    expect(editor.snapshot().canUndo).toBe(false);
    expect(editor.snapshot().slideWidth).toBe(12_192_000);
  });

  test('a size out of 1 to 56 inches keeps the dialog open with its reason', async () => {
    let closed = 0;
    const editor = await open(() => (closed += 1));
    await type('height', '0.5');
    button('OK').click();
    await nextTick();
    expect(closed).toBe(0);
    const alerts = [...document.querySelectorAll('[role="alert"]')].map((a) => a.textContent);
    expect(alerts).toContain('The slide width and height must each be from 1 to 56 inches.');
    expect(editor.snapshot().slideHeight).toBe(6_858_000);
  });
});
