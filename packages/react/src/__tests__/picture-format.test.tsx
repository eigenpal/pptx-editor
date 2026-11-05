import { afterAll, afterEach, beforeAll, describe, expect, spyOn, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import { PNG } from 'pngjs';
import {
  createPptxEditor,
  resetDevWarnings,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../context.ts';
import {
  ImageAltTextDialog,
  ImagePropertiesDialog,
  PptxEditor,
  useEditorCommand,
  useImageAltTextDialog,
  useImagePropertiesDialog,
} from '../index.ts';
import { PptxEditorContent, PptxEditorRoot } from '../PptxEditorRoot.tsx';
import { tooltipText } from './tooltip-text.ts';

// The picture chrome: the contextual `image`
// group on the toolbar, Crop, and the Alt Text dialog in the popups registry.

const IN = 914_400;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

function png(): Uint8Array {
  const img = new PNG({ width: 4, height: 3 });
  img.data.fill(200);
  return new Uint8Array(PNG.sync.write(img));
}

const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        rels: [['rIdImg', R_IMAGE, '../media/image1.png']],
        shapes:
          '<p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture 1" descr="A chart"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>' +
          '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
          `<p:spPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${4 * IN}" cy="${3 * IN}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>` +
          shape({ id: 3, x: 6 * IN, y: IN, cx: IN, cy: IN }) +
          '<p:pic><p:nvPicPr><p:cNvPr id="4" name="Picture 2"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
          '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
          `<p:spPr><a:xfrm><a:off x="${IN}" y="${5 * IN}"/><a:ext cx="${IN}" cy="${IN}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`,
      },
    ],
    media: [{ name: 'image1.png', bytes: png(), contentType: 'image/png' }],
  });

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(deck())).ok).toBe(true);
  return editor;
}

const nodeAt = (editor: PptxEditorInstance, i: number): string =>
  editor.session!.layout(0)!.items[i]!.element.nodeId;

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
afterEach(cleanup);

const slot = (id: string): HTMLElement =>
  document.querySelector<HTMLElement>(`.pptx-toolbar [data-slot="${id}"]`)!;
const barGroups = (): (string | null)[] =>
  [...document.querySelectorAll('.pptx-toolbar > [data-group]')].map((g) =>
    g.getAttribute('data-group')
  );

describe('the image group on the toolbar', () => {
  test('joins the bar while pictures are selected, as DOCX Editor’s image group', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Toolbar />
      </PptxEditorContext.Provider>
    );
    expect(barGroups()).not.toContain('image');
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 1)] }));
    expect(barGroups()).not.toContain('image');
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 0)] }));
    // A picture shows the picture tools and none of the text or shape groups.
    expect(barGroups()).toEqual(['history', 'insert', 'image', 'arrange', 'more']);
    for (const id of ['image.crop', 'image.change', 'image.altText', 'image.transparency'])
      expect(slot(id) !== null).toBe(true);
    // Two pictures keep the group: the Picture Format tab stays.
    act(() =>
      editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 0), nodeAt(editor, 2)] })
    );
    expect(barGroups()).toContain('image');
    expect(editor.snapshot().image?.count).toBe(2);
    // A shape in the selection removes it.
    act(() =>
      editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 0), nodeAt(editor, 1)] })
    );
    expect(barGroups()).not.toContain('image');
  });

  test('Crop is a pressed toggle over crop mode; a pick runs its command', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Toolbar />
      </PptxEditorContext.Provider>
    );
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 0)] }));
    const crop = slot('image.crop');
    expect(crop.getAttribute('aria-pressed')).toBe('false');
    act(() => void fireEvent.click(crop));
    expect(editor.snapshot().image?.cropping).toBe(true);
    expect(slot('image.crop').getAttribute('aria-pressed')).toBe('true');
    act(() => void fireEvent.click(slot('image.transparency')));
    act(
      () => void fireEvent.click(screen.getByRole('menuitemradio', { name: 'Transparency: 50%' }))
    );
    expect(editor.snapshot().image?.transparency).toBe(0.5);
    expect(editor.snapshot().canUndo).toBe(true);
  });

  test('Corrections is a grid of samples, a row per contrast; a cell sets both', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Toolbar />
      </PptxEditorContext.Provider>
    );
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 0)] }));
    act(() => void fireEvent.click(slot('image.corrections')));
    const grid = document.querySelector('.pptx-corrections-grid')!;
    expect(grid.getAttribute('data-columns')).toBe('5');
    const rows = [...grid.querySelectorAll('.pptx-style-gallery__row')];
    expect(rows.map((r) => r.querySelectorAll('[role="menuitemradio"]').length)).toEqual([
      5, 5, 5, 5, 5,
    ]);
    // Each cell is a sample picture, named by its values; the current values are checked.
    expect(grid.querySelectorAll('.pptx-corrections-grid__swatch').length).toBe(25);
    const checked = grid.querySelector('[aria-checked="true"]')!;
    expect(checked.getAttribute('data-value')).toBe('0,0');
    expect(checked.getAttribute('aria-label')).toBe('Brightness: 0% Contrast: 0%');
    // The axes are named: Brightness over the columns, Contrast beside the rows.
    expect(grid.getAttribute('aria-label')).toBe('Brightness and contrast');
    expect(grid.querySelector('.pptx-corrections-grid__title')?.textContent).toBe('Brightness');
    expect(grid.querySelector('.pptx-corrections-grid__side')?.textContent).toBe('Contrast');
    expect(
      [...grid.querySelectorAll('.pptx-corrections-grid__head span')].map((l) => l.textContent)
    ).toEqual(['-40%', '-20%', '0%', '+20%', '+40%']);
    // The grid opens on the current values; Down moves to the next contrast.
    expect(document.activeElement).toBe(checked);
    fireEvent.keyDown(checked, { key: 'ArrowDown' });
    expect((document.activeElement as HTMLElement).getAttribute('data-value')).toBe('0,20');
    const cell = grid.querySelector<HTMLElement>('[data-value="20,-20"]')!;
    expect(cell.getAttribute('aria-label')).toBe('Brightness: +20% Contrast: -20%');
    act(() => void fireEvent.click(cell));
    expect(editor.snapshot().image).toMatchObject({ brightness: 0.2, contrast: -0.2 });
    expect(document.querySelector('.pptx-corrections-grid')).toBeNull();
  });

  test('Change picture takes the same file a second time', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Toolbar />
      </PptxEditorContext.Provider>
    );
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 0)] }));
    const input =
      slot('image.change').parentElement!.querySelector<HTMLInputElement>('input[type="file"]')!;
    // Browsers fire `change` only for a new value: the input clears after each pick.
    Object.defineProperty(input, 'files', { configurable: true, value: [] });
    fireEvent.change(input);
    expect(input.value).toBe('');
  });

  test('the context menu of a picture is the picture menu, named for it', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditorContent />
      </PptxEditorContext.Provider>
    );
    const stage = document.querySelector('.pptx-surface__stage') as HTMLElement;
    stage.getBoundingClientRect = () =>
      ({
        left: 0,
        top: 0,
        right: 1280,
        bottom: 720,
        width: 1280,
        height: 720,
        x: 0,
        y: 0,
      }) as DOMRect;
    // The picture spans 1 in to 5 in: one CSS pixel is 9525 EMU at zoom 1.
    act(() => {
      fireEvent.contextMenu(stage, { clientX: (2 * IN) / 9525, clientY: (2 * IN) / 9525 });
    });
    screen.getByRole('menu', { name: 'Picture actions' });
    screen.getByRole('menuitem', { name: 'Change picture…' });
  });

  test('in a preset bar, an image named part replaces its control in the picture group', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Toolbar>
          <PptxEditor.Toolbar.ImageCrop className="acme-crop" />
        </PptxEditor.Toolbar>
      </PptxEditorContext.Provider>
    );
    // The picture group shows only while every selected element is a picture.
    expect(document.querySelector('[data-slot="image.crop"]') === null).toBe(true);
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 0)] }));
    const crops = document.querySelectorAll<HTMLElement>('[data-slot="image.crop"]');
    expect(crops.length).toBe(1);
    expect(crops[0]!.classList.contains('acme-crop')).toBe(true);
    expect(crops[0]!.closest('[data-group]')?.getAttribute('data-group')).toBe('image');
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 1)] }));
    expect(document.querySelector('[data-slot="image.crop"]') === null).toBe(true);
  });

  test('a named part customizes an image control', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Toolbar preset={false}>
          <PptxEditor.Toolbar.ImageCrop className="acme-crop" />
        </PptxEditor.Toolbar>
      </PptxEditorContext.Provider>
    );
    const button = document.querySelector<HTMLElement>('[data-slot="image.crop"]')!;
    expect(button.classList.contains('acme-crop')).toBe(true);
    // Disabled with the reason until a picture is selected.
    expect(button.getAttribute('aria-disabled') ?? String(button.hasAttribute('disabled'))).toBe(
      'true'
    );
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 0)] }));
    const enabled = document.querySelector<HTMLElement>('[data-slot="image.crop"]')!;
    expect(enabled.hasAttribute('disabled')).toBe(false);
  });
});

describe('Alt Text dialog', () => {
  test('opens with the alt text; OK applies the draft as one undo step', async () => {
    const editor = await open();
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 0)] }));
    let closed = 0;
    render(
      <PptxEditorContext.Provider value={editor}>
        <ImageAltTextDialog open onClose={() => (closed += 1)} />
      </PptxEditorContext.Provider>
    );
    const text = screen.getByLabelText('Description') as HTMLTextAreaElement;
    expect(text.value).toBe('A chart');
    act(() => void fireEvent.change(text, { target: { value: 'Sales by region' } }));
    expect(editor.snapshot().geometry?.altText?.description).toBe('A chart');
    act(() => void fireEvent.click(document.querySelector('[data-pptx-part="apply"]')!));
    expect(closed).toBe(1);
    expect(editor.snapshot().geometry?.altText?.description).toBe('Sales by region');
    act(() => void editor.exec({ type: 'undo' }));
    expect(editor.snapshot().geometry?.altText?.description).toBe('A chart');
  });

  test('Enter in the title applies; a long title gets a named, linked, focused error', async () => {
    const editor = await open();
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 0)] }));
    let closed = 0;
    render(
      <PptxEditorContext.Provider value={editor}>
        <ImageAltTextDialog open onClose={() => (closed += 1)} />
      </PptxEditorContext.Provider>
    );
    const description = screen.getByLabelText('Description') as HTMLTextAreaElement;
    const title = screen.getByLabelText('Title') as HTMLInputElement;
    // Enter in the description is a new line, not OK.
    act(() => void fireEvent.keyDown(description, { key: 'Enter' }));
    expect(closed).toBe(0);
    act(() => void fireEvent.change(title, { target: { value: 'x'.repeat(4001) } }));
    act(() => void fireEvent.keyDown(title, { key: 'Enter' }));
    expect(closed).toBe(0);
    const error = document.querySelector('[data-pptx-part="error"]')!;
    expect(error.textContent).toMatch(/^Title: /);
    expect(title.getAttribute('aria-invalid')).toBe('true');
    expect(title.getAttribute('aria-describedby')).toBe(error.id);
    expect(document.activeElement).toBe(title);
    act(() => void fireEvent.change(title, { target: { value: 'Sales' } }));
    act(() => void fireEvent.keyDown(title, { key: 'Enter' }));
    expect(closed).toBe(1);
    expect(editor.snapshot().geometry?.altText?.title).toBe('Sales');
  });

  test('decorative turns the description off; Cancel discards', async () => {
    const editor = await open();
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 0)] }));
    render(
      <PptxEditorContext.Provider value={editor}>
        <ImageAltTextDialog open onClose={() => {}} />
      </PptxEditorContext.Provider>
    );
    const text = screen.getByLabelText('Description') as HTMLTextAreaElement;
    expect(text.getAttribute('aria-describedby')).not.toContain('decorative-reason');
    act(() => void fireEvent.click(screen.getByLabelText('Mark as decorative')));
    expect(text.disabled).toBe(true);
    // A screen reader hears why the field is off.
    const reason = document.getElementById(
      text
        .getAttribute('aria-describedby')!
        .split(' ')
        .find((x) => x.endsWith('decorative-reason'))!
    )!;
    expect(reason.textContent).toBe(
      'The description is off because the object is marked as decorative.'
    );
    expect(reason.classList.contains('pptx-sr-only')).toBe(true);
    act(() => void fireEvent.click(document.querySelector('[data-pptx-part="cancel"]')!));
    expect(editor.snapshot().canUndo).toBe(false);
  });

  test('a host field writes the draft through useImageAltTextDialog', async () => {
    const editor = await open();
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 1)] }));
    function Suggest() {
      const { setValue } = useImageAltTextDialog();
      return (
        <button type="button" onClick={() => setValue('description', 'A blue square')}>
          Suggest
        </button>
      );
    }
    render(
      <PptxEditorContext.Provider value={editor}>
        <ImageAltTextDialog open onClose={() => {}}>
          <ImageAltTextDialog.Body>
            <Suggest />
          </ImageAltTextDialog.Body>
        </ImageAltTextDialog>
      </PptxEditorContext.Provider>
    );
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Suggest' })));
    act(() => void fireEvent.click(document.querySelector('[data-pptx-part="apply"]')!));
    expect(editor.snapshot().geometry?.altText?.description).toBe('A blue square');
  });

  test('the toolbar opens it from the registry; popups false opens nothing', async () => {
    let editor: PptxEditorInstance | null = null;
    render(
      <PptxEditorRoot onReady={(e) => (editor = e)}>
        <PptxEditor.Toolbar preset={false}>
          <PptxEditor.Toolbar.ImageAltText />
        </PptxEditor.Toolbar>
      </PptxEditorRoot>
    );
    await act(async () => void (await editor!.load(deck())));
    act(() => editor!.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor!, 0)] }));
    act(() => void fireEvent.click(document.querySelector('[data-slot="image.altText"]')!));
    expect(document.querySelector('[data-pptx-dialog="imageAltText"]') !== null).toBe(true);
    cleanup();
    render(
      <PptxEditorRoot onReady={(e) => (editor = e)} popups={{ imageAltText: false }}>
        <PptxEditor.Toolbar preset={false}>
          <PptxEditor.Toolbar.ImageAltText />
        </PptxEditor.Toolbar>
      </PptxEditorRoot>
    );
    await act(async () => void (await editor!.load(deck())));
    act(() => editor!.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor!, 0)] }));
    const off = document.querySelector<HTMLElement>('[data-slot="image.altText"]')!;
    // The button is off with the reason; a click opens nothing.
    expect(off.getAttribute('aria-disabled')).toBe('true');
    expect(tooltipText(off)).toContain('This dialog is turned off');
    act(() => void fireEvent.click(off));
    expect(document.querySelector('[data-pptx-dialog="imageAltText"]') === null).toBe(true);
  });

  test('popups false turns Format Picture off with the reason; alt text values still apply', async () => {
    let editor: PptxEditorInstance | null = null;
    let alt: ReturnType<typeof useEditorCommand> | null = null;
    function AltProbe() {
      alt = useEditorCommand('image.altText');
      return null;
    }
    render(
      <PptxEditorRoot
        onReady={(e) => (editor = e)}
        popups={{ imageProperties: false, imageAltText: false }}
      >
        <PptxEditor.Toolbar preset={false}>
          <PptxEditor.Toolbar.ImageProperties />
        </PptxEditor.Toolbar>
        <AltProbe />
      </PptxEditorRoot>
    );
    await act(async () => void (await editor!.load(deck())));
    act(() => editor!.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor!, 0)] }));
    const button = document.querySelector<HTMLElement>('[data-slot="image.properties"]')!;
    expect(button.getAttribute('aria-disabled')).toBe('true');
    expect(tooltipText(button)).toContain('This dialog is turned off');
    // Hiding the popup never turns off the operation: a value runs `setAltText`.
    act(() => void alt!.execute('Sales by region'));
    expect(editor!.snapshot().geometry?.altText?.description).toBe('Sales by region');
    // `execute()` with no description would open the dialog, which is off: it says so.
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      let ran = true;
      act(() => void (ran = alt!.execute()));
      expect(ran).toBe(false);
      expect(warn.mock.calls.map((c: unknown[]) => String(c[0])).join('')).toContain(
        'Pass the description instead: execute(text)'
      );
    } finally {
      warn.mockRestore();
    }
  });

  test('the dialog closes when the selection moves to another object or a document opens', async () => {
    const editor = await open();
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 0)] }));
    let closed = 0;
    render(
      <PptxEditorContext.Provider value={editor}>
        <ImageAltTextDialog open onClose={() => (closed += 1)} />
      </PptxEditorContext.Provider>
    );
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 1)] }));
    expect(closed).toBe(1);
    cleanup();
    closed = 0;
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 0)] }));
    render(
      <PptxEditorContext.Provider value={editor}>
        <ImageAltTextDialog open onClose={() => (closed += 1)} />
      </PptxEditorContext.Provider>
    );
    await act(async () => void (await editor.load(deck())));
    expect(closed).toBeGreaterThan(0);
  });

  test('errors in the dialog state are messages, not keys', async () => {
    const editor = await open();
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 0)] }));
    let seen: Record<string, unknown> = {};
    function LongTitle() {
      const dialog = useImageAltTextDialog();
      seen = dialog.errors;
      return (
        <button type="button" onClick={() => dialog.setValue('title', 'x'.repeat(4001))}>
          Long
        </button>
      );
    }
    render(
      <PptxEditorContext.Provider value={editor}>
        <ImageAltTextDialog open onClose={() => {}}>
          <ImageAltTextDialog.Footer>
            <LongTitle />
          </ImageAltTextDialog.Footer>
        </ImageAltTextDialog>
      </PptxEditorContext.Provider>
    );
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Long' })));
    // The footer part replaced the OK button: Enter in the title applies and runs the checks.
    act(() => void fireEvent.keyDown(screen.getByLabelText('Title'), { key: 'Enter' }));
    expect(typeof seen.title).toBe('string');
    expect(String(seen.title)).not.toMatch(/^errors\./);
    expect(String(seen.title).length).toBeGreaterThan(0);
  });
});

describe('Format Picture dialog', () => {
  test('size with Lock aspect ratio and the crop apply as one undo step', async () => {
    const editor = await open();
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 0)] }));
    let closed = 0;
    render(
      <PptxEditorContext.Provider value={editor}>
        <ImagePropertiesDialog open onClose={() => (closed += 1)} />
      </PptxEditorContext.Provider>
    );
    // Size and Crop are labelled groups, so a screen reader names the field's section.
    for (const name of ['Size', /^Crop/]) expect(screen.getByRole('group', { name })).toBeTruthy();
    const width = screen.getByLabelText(/^Width/) as HTMLInputElement;
    const height = screen.getByLabelText(/^Height/) as HTMLInputElement;
    expect([width.value, height.value]).toEqual(['4', '3']);
    act(() => void fireEvent.change(width, { target: { value: '2' } }));
    expect(height.value).toBe('1.5');
    act(() => void fireEvent.change(screen.getByLabelText(/^Left/), { target: { value: '25' } }));
    act(() => void fireEvent.click(document.querySelector('[data-pptx-part="apply"]')!));
    expect(closed).toBe(1);
    const g = editor.snapshot().geometry!;
    expect([g.cx, g.cy]).toEqual([2 * IN, 1.5 * IN]);
    expect(editor.snapshot().image?.crop.left).toBe(25);
    act(() => void editor.exec({ type: 'undo' }));
    expect(editor.snapshot().geometry?.cx).toBe(4 * IN);
    expect(editor.snapshot().image?.crop.left).toBe(0);
  });

  test('two setValue calls in one event keep both values (the live draft)', async () => {
    const editor = await open();
    // Picture 2 has no aspect lock: width and height are independent.
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 2)] }));
    function Both() {
      const { setValue } = useImagePropertiesDialog();
      const both = (): void => {
        setValue('width', 2);
        setValue('height', 3);
      };
      return (
        <button type="button" onClick={both}>
          Both
        </button>
      );
    }
    render(
      <PptxEditorContext.Provider value={editor}>
        <ImagePropertiesDialog open onClose={() => {}}>
          <ImagePropertiesDialog.Body>
            <Both />
          </ImagePropertiesDialog.Body>
        </ImagePropertiesDialog>
      </PptxEditorContext.Provider>
    );
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Both' })));
    act(() => void fireEvent.click(document.querySelector('[data-pptx-part="apply"]')!));
    const g = editor.snapshot().geometry!;
    expect([g.cx, g.cy]).toEqual([2 * IN, 3 * IN]);
  });

  test('a size out of range keeps the dialog open with the reason', async () => {
    const editor = await open();
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 0)] }));
    let closed = 0;
    render(
      <PptxEditorContext.Provider value={editor}>
        <ImagePropertiesDialog open onClose={() => (closed += 1)} />
      </PptxEditorContext.Provider>
    );
    act(() => void fireEvent.click(screen.getByLabelText('Lock aspect ratio')));
    act(() => void fireEvent.change(screen.getByLabelText(/^Width/), { target: { value: '0' } }));
    expect((screen.getByLabelText(/^Height/) as HTMLInputElement).value).toBe('3');
    act(() => void fireEvent.click(document.querySelector('[data-pptx-part="apply"]')!));
    expect(closed).toBe(0);
    // The message names the field, the field points at it, and focus goes to it (WCAG 3.3.1).
    const error = document.querySelector('[data-pptx-part="error"]')!;
    expect(error.textContent).toBe('Width: Enter a size from 0.01 to 56 inches.');
    const width = screen.getByLabelText(/^Width/) as HTMLInputElement;
    expect(width.getAttribute('aria-invalid')).toBe('true');
    expect(width.getAttribute('aria-describedby')).toBe(error.id);
    expect(document.activeElement).toBe(width);
  });

  test('a locked crop or size: the fields are off, and one note under the heading says why', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect(
      (
        await editor.load(
          buildPptx({
            slides: [
              {
                layout: 'blank',
                rels: [['rIdImg', R_IMAGE, '../media/image1.png']],
                shapes:
                  '<p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture 1"/><p:cNvPicPr><a:picLocks noCrop="1" noResize="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>' +
                  '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
                  `<p:spPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${4 * IN}" cy="${3 * IN}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`,
              },
            ],
            media: [{ name: 'image1.png', bytes: png(), contentType: 'image/png' }],
          })
        )
      ).ok
    ).toBe(true);
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 0)] }));
    render(
      <PptxEditorContext.Provider value={editor}>
        <ImagePropertiesDialog open onClose={() => {}} />
      </PptxEditorContext.Provider>
    );
    const cropNote = screen.getByText('This picture is locked against cropping');
    const sizeNote = screen.getByText('This picture is locked against resizing');
    for (const name of [/^Left/, /^Top/, /^Right/, /^Bottom/]) {
      const input = screen.getByLabelText(name) as HTMLInputElement;
      expect(input.disabled).toBe(true);
      expect(input.getAttribute('aria-describedby')).toBe(cropNote.id);
    }
    for (const name of [/^Width/, /^Height/]) {
      const input = screen.getByLabelText(name) as HTMLInputElement;
      expect(input.disabled).toBe(true);
      expect(input.getAttribute('aria-describedby')).toBe(sizeNote.id);
    }
    const lock = screen.getByLabelText(/Lock aspect ratio/i) as HTMLInputElement;
    expect(lock.disabled).toBe(true);
    expect(lock.getAttribute('aria-describedby')).toBe(sizeNote.id);
    // The rotation field stays on.
    const rotation = screen.getByLabelText(/^Rotation/) as HTMLInputElement;
    expect(rotation.disabled).toBe(false);
    expect(editor.snapshot().image?.canRotate).toBe(true);
  });

  test('a picture locked against rotation: the rotation field is off with the reason', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const shapes =
      '<p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture 1"/><p:cNvPicPr><a:picLocks noRot="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
      `<p:spPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${4 * IN}" cy="${3 * IN}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
    const bytes = buildPptx({
      slides: [{ layout: 'blank', rels: [['rIdImg', R_IMAGE, '../media/image1.png']], shapes }],
      media: [{ name: 'image1.png', bytes: png(), contentType: 'image/png' }],
    });
    expect((await editor.load(bytes)).ok).toBe(true);
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 0)] }));
    expect(editor.snapshot().image?.canRotate).toBe(false);
    render(
      <PptxEditorContext.Provider value={editor}>
        <ImagePropertiesDialog open onClose={() => {}} />
      </PptxEditorContext.Provider>
    );
    const note = screen.getByText('This picture is locked against rotation');
    const rotation = screen.getByLabelText(/^Rotation/) as HTMLInputElement;
    expect(rotation.disabled).toBe(true);
    expect(rotation.getAttribute('aria-describedby')).toBe(note.id);
    expect((screen.getByLabelText(/^Width/) as HTMLInputElement).disabled).toBe(false);
  });

  test('an error names and focuses the first field in screen order: Height', async () => {
    const editor = await open();
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeAt(editor, 0)] }));
    render(
      <PptxEditorContext.Provider value={editor}>
        <ImagePropertiesDialog open onClose={() => {}} />
      </PptxEditorContext.Provider>
    );
    // Lock aspect ratio is on: a height of 0 makes the width 0 too.
    act(() => void fireEvent.change(screen.getByLabelText(/^Height/), { target: { value: '0' } }));
    act(() => void fireEvent.click(document.querySelector('[data-pptx-part="apply"]')!));
    expect(document.querySelector('[data-pptx-part="error"]')!.textContent).toMatch(/^Height: /);
    expect(document.activeElement).toBe(screen.getByLabelText(/^Height/));
    expect(editor.snapshot().canUndo).toBe(false);
  });
});
