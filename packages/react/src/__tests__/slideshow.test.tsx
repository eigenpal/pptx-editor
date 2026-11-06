import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import type { ReactNode } from 'react';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import {
  PptxEditorContext,
  SlideshowContext,
  useSlideshow,
  type UseSlideshowReturn,
} from '../context.ts';
import { PptxEditorViewport } from '../PptxEditorRoot.tsx';
import { Slideshow, useSlideshowHandle } from '../Slideshow.tsx';

// The slide show in a DOM: start and its shortcuts, the show keys, hidden slides, the
// control bar, click to advance, the end screen, and exit back to the editor.

const IN = 914_400;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
/** A picture with alt text: the show reads it out. */
const PICTURE =
  '<p:pic><p:nvPicPr><p:cNvPr id="3" name="Lake" descr="A lake at dusk"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
  '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
  '<p:spPr><a:xfrm><a:off x="5486400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
const MEDIA = [
  { name: 'image1.png', bytes: new Uint8Array([137, 80, 78, 71]), contentType: 'image/png' },
];

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const box = (id: number, text: string) =>
    shape({ id, x: IN, y: IN, cx: 3 * IN, cy: IN, text: para(run(text)) });
  const r = await editor.load(
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: box(2, 'One') + PICTURE,
          rels: [['rIdImg', R_IMAGE, '../media/image1.png']],
        },
        { layout: 'blank', shapes: box(2, 'Two'), hidden: true },
        { layout: 'blank', shapes: box(2, 'Three') },
      ],
      media: MEDIA,
    })
  );
  expect(r.ok).toBe(true);
  return editor;
}

let handle: UseSlideshowReturn | null = null;
function Probe() {
  handle = useSlideshow();
  return null;
}

/** The providers `PptxEditor.Root` publishes, around one editor made for the test. */
function Host({ editor, children }: { editor: PptxEditorInstance; children: ReactNode }) {
  const slideshow = useSlideshowHandle(editor);
  return (
    <PptxEditorContext.Provider value={editor}>
      <SlideshowContext.Provider value={slideshow}>{children}</SlideshowContext.Provider>
    </PptxEditorContext.Provider>
  );
}

async function setup(): Promise<PptxEditorInstance> {
  const editor = await open();
  render(
    <Host editor={editor}>
      <PptxEditorViewport>
        <button type="button">inside</button>
        <Slideshow />
        <Probe />
      </PptxEditorViewport>
    </Host>
  );
  return editor;
}

const dialog = () => screen.queryByRole('dialog', { name: 'Slide show' });
const position = () => document.querySelector('.pptx-show__position')?.textContent;
const key = (k: string, init: Record<string, unknown> = {}) =>
  act(() => {
    fireEvent.keyDown(document.activeElement ?? document.body, { key: k, ...init });
  });

afterEach(() => {
  cleanup();
  handle = null;
  document.body.replaceChildren();
});

describe('Slideshow', () => {
  test('starts from the first slide and paints it with the editor painter', async () => {
    const editor = await setup();
    expect(dialog() === null).toBe(true);
    act(() => handle!.start('start'));
    expect(dialog() !== null).toBe(true);
    expect(position()).toBe('Slide 1 of 3');
    const slide = document.querySelector('.pptx-show__slide')!;
    expect(slide.getAttribute('data-slide-index')).toBe('0');
    const svg = slide.querySelector('svg')!;
    expect(svg.getAttribute('data-slide-part')).toBe(editor.session!.info.slides[0].partName);
    // No editing hooks in the show.
    expect(svg.querySelector('[data-node-id]') === null).toBe(true);
    // A screen reader reads the slide: its text, and the picture by its alt text.
    expect(svg.getAttribute('role')).toBe('group');
    expect(svg.textContent).toContain('One');
    expect(svg.querySelector('[role="img"]')?.getAttribute('aria-label')).toBe('A lake at dusk');
  });

  test('keys step through the shown slides, skip the hidden one, and stay in the show', async () => {
    const editor = await setup();
    act(() => handle!.start('start'));
    let leaked = 0;
    const count = () => (leaked += 1);
    document.body.addEventListener('keydown', count);
    await key('ArrowRight');
    expect(position()).toBe('Slide 3 of 3');
    await key('Backspace');
    expect(position()).toBe('Slide 1 of 3');
    await key('End');
    expect(position()).toBe('Slide 3 of 3');
    await key('2');
    expect(screen.getByRole('status').textContent).toBe('Go to slide 2');
    await key('Enter');
    expect(position()).toBe('Slide 2 of 3');
    await key('b');
    expect(document.querySelector('.pptx-show__blank')?.getAttribute('data-color')).toBe('black');
    await key('b');
    expect(document.querySelector('.pptx-show__blank') === null).toBe(true);
    expect(leaked).toBe(0);
    document.body.removeEventListener('keydown', count);
    // The deck is unchanged: no key reached the editor.
    expect(editor.snapshot().slides).toHaveLength(3);
  });

  test('Escape ends the show and the editor shows the slide shown last', async () => {
    const editor = await setup();
    act(() => handle!.start('start'));
    await key('ArrowRight');
    await key('Escape');
    expect(dialog() === null).toBe(true);
    expect(editor.snapshot().currentSlide).toBe(2);
  });

  test('Escape gives focus back to the opener, or to the slide canvas when the opener left', async () => {
    const editor = await open();
    render(
      <Host editor={editor}>
        <PptxEditorViewport>
          <div className="pptx-surface" tabIndex={0} data-testid="canvas" />
          <button type="button">opener</button>
          <Slideshow />
          <Probe />
        </PptxEditorViewport>
      </Host>
    );
    const opener = screen.getByRole('button', { name: 'opener' });
    opener.focus();
    act(() => handle!.start('start'));
    await key('Escape');
    expect(document.activeElement).toBe(opener);
    // A menu row starts the show and closes: it is gone when the show ends.
    opener.focus();
    act(() => handle!.start('start'));
    opener.remove();
    await key('Escape');
    await new Promise((r) => setTimeout(r, 50));
    expect(document.activeElement).toBe(screen.getByTestId('canvas'));
  });

  test('the control bar steps and exits; a click advances through the end screen', async () => {
    const editor = await setup();
    act(() => handle!.start('start'));
    const bar = screen.getByRole('toolbar', { name: 'Slide show controls' });
    expect(bar.getAttribute('data-visible')).toBe('false');
    fireEvent.mouseMove(dialog()!);
    expect(bar.getAttribute('data-visible')).toBe('true');
    fireEvent.click(screen.getByRole('button', { name: 'Next slide' }));
    expect(position()).toBe('Slide 3 of 3');
    fireEvent.click(screen.getByRole('button', { name: 'Previous slide' }));
    expect(position()).toBe('Slide 1 of 3');
    // Clicks on the slide advance; after the last slide comes the end screen.
    fireEvent.click(dialog()!);
    fireEvent.click(dialog()!);
    expect(screen.getByText('End of slide show. Click to exit.') !== null).toBe(true);
    expect(document.querySelector('.pptx-show__slide svg') === null).toBe(true);
    fireEvent.click(dialog()!);
    expect(dialog() === null).toBe(true);
    expect(editor.snapshot().currentSlide).toBe(2);

    act(() => handle!.start('start'));
    fireEvent.click(screen.getByRole('button', { name: 'End show' }));
    expect(dialog() === null).toBe(true);
    expect(editor.snapshot().currentSlide).toBe(0);
  });

  test('F5 starts from the start, Ctrl+Enter from the current slide, in viewing mode too', async () => {
    const editor = await setup();
    editor.setMode('view');
    act(() => editor.goToSlide(1));
    await key('F5');
    expect(position()).toBe('Slide 1 of 3');
    await key('Escape');
    act(() => editor.goToSlide(1));
    screen.getByRole('button', { name: 'inside' }).focus();
    await key('Enter', { ctrlKey: true });
    // From the current slide shows it even when it is hidden.
    expect(position()).toBe('Slide 2 of 3');
  });

  test('does not start without a document', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    render(
      <Host editor={editor}>
        <PptxEditorViewport>
          <Slideshow />
          <Probe />
        </PptxEditorViewport>
      </Host>
    );
    act(() => handle!.start('start'));
    await key('F5');
    expect(dialog() === null).toBe(true);
  });
});
