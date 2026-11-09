import { afterAll, afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { PNG } from 'pngjs';
import { createApp, defineComponent, h, nextTick, shallowRef, type App, type Component } from 'vue';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx } from '../../../scripts/fixtures/builder.ts';
import { providePptxEditor } from '../src/composables.ts';
import {
  ImageAltTextDialog,
  ImagePropertiesDialog,
  PptxEditorToolbar,
  useImageAltTextDialog,
} from '../src/index.ts';

// The Vue twins of the picture chrome (packages/react/src/__tests__/picture-format.test.tsx):
// the `image` toolbar group, and the dialogs with the same fields, the same draft rules from
// core, and one undo step on OK.

const IN = 914_400;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

let app: App | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
});

async function open(locks = 'noChangeAspect="1"'): Promise<PptxEditorInstance> {
  const img = new PNG({ width: 4, height: 3 });
  img.data.fill(200);
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const r = await editor.load(
    buildPptx({
      slides: [
        {
          layout: 'blank',
          rels: [['rIdImg', R_IMAGE, '../media/image1.png']],
          shapes:
            `<p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture 1" descr="A chart"/><p:cNvPicPr><a:picLocks ${locks}/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
            '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
            `<p:spPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${4 * IN}" cy="${3 * IN}"/></a:xfrm>` +
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>',
        },
      ],
      media: [
        {
          name: 'image1.png',
          bytes: new Uint8Array(PNG.sync.write(img)),
          contentType: 'image/png',
        },
      ],
    })
  );
  expect(r.ok).toBe(true);
  const id = editor.session!.layout(0)!.items[0]!.element.nodeId;
  editor.setSelection({ kind: 'shapes', nodeIds: [id] });
  return editor;
}

async function mount(
  editor: PptxEditorInstance,
  dialog: Component,
  props: Record<string, unknown> = { open: true }
): Promise<{ closed: number }> {
  const state = { closed: 0 };
  host = document.createElement('div');
  document.body.append(host);
  app = createApp(
    defineComponent({
      setup() {
        providePptxEditor(shallowRef(editor));
        return () => h(dialog, { ...props, onClose: () => (state.closed += 1) });
      },
    }) as Component
  );
  app.mount(host);
  await nextTick();
  return state;
}

const part = (name: string) => document.querySelector<HTMLElement>(`[data-pptx-part="${name}"]`);
const input = (field: string, index = 0) =>
  document
    .querySelector(`[data-pptx-field="${field}"]`)!
    .querySelectorAll<HTMLInputElement | HTMLTextAreaElement>('input, textarea')[index]!;

function type(el: HTMLInputElement | HTMLTextAreaElement, value: string): void {
  el.value = value;
  el.dispatchEvent(new Event('input', { bubbles: true }));
}

describe('the image group on the Vue toolbar', () => {
  // happy-dom lays nothing out: give the bar room, so no group moves into More.
  const width = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth');
  beforeAll(() => {
    Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
      configurable: true,
      get: () => 100_000,
    });
  });
  afterAll(() => {
    if (width) Object.defineProperty(HTMLElement.prototype, 'clientWidth', width);
  });

  test('joins the bar while a picture is selected, and leaves with it', async () => {
    const editor = await open();
    await mount(editor, PptxEditorToolbar as Component, {});
    const groups = () =>
      [...document.querySelectorAll('.pptx-toolbar > [data-group]')].map((g) =>
        g.getAttribute('data-group')
      );
    // A picture shows the picture tools and none of the text or shape groups, as in React.
    expect(groups()).toEqual(['history', 'insert', 'image', 'arrange', 'more']);
    // The picture group's bar controls, as in React. Format > Picture and the picture context
    // menu hold Color, Transparency, Reset, Crop to shape, and Crop fit.
    const slots = [...document.querySelectorAll('[data-group="image"] [data-slot]')].map((e) =>
      e.getAttribute('data-slot')
    );
    expect(slots).toEqual([
      'image.corrections',
      'image.change',
      'image.styles',
      'image.altText',
      'image.crop',
      'image.properties',
    ]);
    editor.setSelection({ kind: 'none' });
    await nextTick();
    expect(groups()).not.toContain('image');
  });

  test('in a preset bar, an image named part replaces its control in the picture group', async () => {
    const editor = await open();
    editor.setSelection({ kind: 'none' });
    await mount(
      editor,
      defineComponent({
        setup: () => () =>
          h(PptxEditorToolbar, null, () => [
            h(PptxEditorToolbar.ImageCrop, { class: 'acme-crop' }),
          ]),
      }) as Component,
      {}
    );
    // The picture group shows only while every selected element is a picture.
    expect(document.querySelector('[data-slot="image.crop"]') === null).toBe(true);
    const id = editor.session!.layout(0)!.items[0]!.element.nodeId;
    editor.setSelection({ kind: 'shapes', nodeIds: [id] });
    await nextTick();
    const crops = document.querySelectorAll<HTMLElement>('[data-slot="image.crop"]');
    expect(crops.length).toBe(1);
    expect(crops[0]!.classList.contains('acme-crop')).toBe(true);
    expect(crops[0]!.closest('[data-group]')?.getAttribute('data-group')).toBe('image');
  });
});

describe('ImageAltTextDialog (Vue)', () => {
  test('opens with the alt text; OK applies the draft as one undo step', async () => {
    const editor = await open();
    const state = await mount(editor, ImageAltTextDialog);
    const text = input('description');
    expect(text.value).toBe('A chart');
    type(text, 'Sales by region');
    await nextTick();
    part('apply')!.click();
    expect(state.closed).toBe(1);
    expect(editor.snapshot().geometry?.altText?.description).toBe('Sales by region');
    editor.exec({ type: 'undo' });
    expect(editor.snapshot().geometry?.altText?.description).toBe('A chart');
  });

  test('Mark as decorative turns the description off and says why to a screen reader', async () => {
    const editor = await open();
    await mount(editor, ImageAltTextDialog);
    const text = input('description');
    expect(text.getAttribute('aria-describedby')).not.toContain('decorative-reason');
    const box = input('decorative') as HTMLInputElement;
    box.checked = true;
    box.dispatchEvent(new Event('change', { bubbles: true }));
    await nextTick();
    expect(input('description').disabled).toBe(true);
    const ids = input('description').getAttribute('aria-describedby')!.split(' ');
    const reason = document.getElementById(ids.find((x) => x.endsWith('decorative-reason'))!)!;
    expect(reason.textContent).toBe(
      'The description is off because the object is marked as decorative.'
    );
    expect(reason.classList.contains('pptx-sr-only')).toBe(true);
  });
});

describe('ImageAltTextDialog (Vue) errors and Enter', () => {
  test('Enter in the title applies; a long title gets a named, focused error', async () => {
    const editor = await open();
    const state = await mount(editor, ImageAltTextDialog);
    const title = input('title') as HTMLInputElement;
    type(title, 'x'.repeat(4001));
    await nextTick();
    title.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    await nextTick();
    expect(state.closed).toBe(0);
    expect(part('error')!.textContent).toMatch(/^Title: /);
    expect(title.getAttribute('aria-describedby')).toBe(part('error')!.id);
    expect(document.activeElement).toBe(title);
    type(title, 'Sales');
    await nextTick();
    title.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(state.closed).toBe(1);
    expect(editor.snapshot().geometry?.altText?.title).toBe('Sales');
  });
});

describe('ImagePropertiesDialog (Vue)', () => {
  test('Lock aspect ratio scales the height, and OK applies size and crop together', async () => {
    const editor = await open();
    const state = await mount(editor, ImagePropertiesDialog);
    for (const field of ['size', 'crop']) {
      const group = document.querySelector(`[data-pptx-field="${field}"]`)!;
      expect(group.getAttribute('role')).toBe('group');
      const label = document.getElementById(group.getAttribute('aria-labelledby')!)!;
      expect(label.textContent).toMatch(field === 'size' ? /^Size$/ : /^Crop/);
    }
    const width = input('size', 1) as HTMLInputElement;
    const height = input('size', 0) as HTMLInputElement;
    expect([width.value, height.value]).toEqual(['4', '3']);
    type(width, '2');
    await nextTick();
    expect((input('size', 0) as HTMLInputElement).value).toBe('1.5');
    type(input('crop', 0) as HTMLInputElement, '25');
    await nextTick();
    part('apply')!.click();
    expect(state.closed).toBe(1);
    const g = editor.snapshot().geometry!;
    expect([g.cx, g.cy]).toEqual([2 * IN, 1.5 * IN]);
    expect(editor.snapshot().image?.crop.left).toBe(25);
    editor.exec({ type: 'undo' });
    expect(editor.snapshot().geometry?.cx).toBe(4 * IN);
  });

  test('a size out of range keeps the dialog open with the reason', async () => {
    const editor = await open();
    const state = await mount(editor, ImagePropertiesDialog);
    type(input('size', 1) as HTMLInputElement, '0');
    await nextTick();
    part('apply')!.click();
    await nextTick();
    expect(state.closed).toBe(0);
    // Lock aspect ratio makes the height 0 too: the first field on screen, Height, is named.
    expect(part('error')!.textContent).toBe('Height: Enter a size from 0.01 to 56 inches.');
    const height = input('size', 0) as HTMLInputElement;
    expect(height.getAttribute('aria-describedby')).toBe(part('error')!.id);
    expect(document.activeElement).toBe(height);
    expect(editor.snapshot().canUndo).toBe(false);
  });
});

describe('picture dialogs (Vue): target, errors, and locks', () => {
  test('a dialog closes when the selection leaves its object or a document opens', async () => {
    for (const dialog of [ImageAltTextDialog, ImagePropertiesDialog]) {
      const editor = await open();
      const state = await mount(editor, dialog);
      editor.setSelection({ kind: 'none' });
      await nextTick();
      expect(state.closed).toBe(1);
      app?.unmount();
      host?.remove();
      const again = await open();
      const reopened = await mount(again, dialog);
      const next = await open();
      expect((await again.load((await next.save())!)).ok).toBe(true);
      await nextTick();
      expect(reopened.closed).toBeGreaterThan(0);
      app?.unmount();
      host?.remove();
      app = null;
      host = null;
    }
  });

  test('errors in the dialog state are messages, not keys', async () => {
    const editor = await open();
    let errors: Readonly<Record<string, string | undefined>> = {};
    const Probe = defineComponent({
      setup() {
        const dialog = useImageAltTextDialog();
        return () => {
          errors = dialog.errors.value;
          return null;
        };
      },
    });
    host = document.createElement('div');
    document.body.append(host);
    app = createApp(
      defineComponent({
        setup() {
          providePptxEditor(shallowRef(editor));
          return () =>
            h(ImageAltTextDialog, { open: true, onClose: () => {} }, { default: () => h(Probe) });
        },
      }) as Component
    );
    app.mount(host);
    await nextTick();
    const title = input('title') as HTMLInputElement;
    type(title, 'x'.repeat(4001));
    await nextTick();
    part('apply')!.click();
    await nextTick();
    expect(typeof errors.title).toBe('string');
    expect(errors.title).not.toMatch(/^errors\./);
  });

  test('a locked crop or size: the fields are off, and one note under the heading says why', async () => {
    const editor = await open('noCrop="1" noResize="1"');
    await mount(editor, ImagePropertiesDialog);
    const cropNote = [...document.querySelectorAll('p')].find(
      (p) => p.textContent === 'This picture is locked against cropping'
    )!;
    const sizeNote = [...document.querySelectorAll('p')].find(
      (p) => p.textContent === 'This picture is locked against resizing'
    )!;
    for (let i = 0; i < 4; i++) {
      const field = input('crop', i) as HTMLInputElement;
      expect(field.disabled).toBe(true);
      expect(field.getAttribute('aria-describedby')).toBe(cropNote.id);
    }
    for (let i = 0; i < 2; i++) {
      const field = input('size', i) as HTMLInputElement;
      expect(field.disabled).toBe(true);
      expect(field.getAttribute('aria-describedby')).toBe(sizeNote.id);
    }
    const lock = document.querySelector('input[type="checkbox"]') as HTMLInputElement;
    expect(lock.disabled).toBe(true);
    expect(lock.getAttribute('aria-describedby')).toBe(sizeNote.id);
    expect((input('rotation') as HTMLInputElement).disabled).toBe(false);
  });

  test('a picture locked against rotation: the rotation field is off with the reason', async () => {
    const editor = await open('noRot="1"');
    expect(editor.snapshot().image?.canRotate).toBe(false);
    await mount(editor, ImagePropertiesDialog);
    const note = [...document.querySelectorAll('p')].find(
      (p) => p.textContent === 'This picture is locked against rotation'
    )!;
    const rotation = input('rotation') as HTMLInputElement;
    expect(rotation.disabled).toBe(true);
    expect(rotation.getAttribute('aria-describedby')).toBe(note.id);
    expect((input('size', 0) as HTMLInputElement).disabled).toBe(false);
  });
});
