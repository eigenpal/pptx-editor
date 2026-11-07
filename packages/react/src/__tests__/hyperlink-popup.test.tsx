import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import type { ReactElement } from 'react';
import { surfaceRootOf, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContent, PptxEditorRoot } from '../PptxEditorRoot.tsx';
import { requestLinkPopup } from '../chrome-actions.tsx';
import { definePopup, PptxEditor, PptxEditorViewport, useHyperlinkPopup } from '../index.ts';

const IN = 914_400;

afterEach(() => {
  cleanup();
  document.body.replaceChildren();
});

async function loaded(ui: (onReady: (e: PptxEditorInstance) => void) => ReactElement) {
  let editor: PptxEditorInstance | null = null;
  render(ui((e) => (editor = e)));
  await act(
    async () =>
      void (await editor!.load(
        buildPptx({
          slides: [
            {
              layout: 'blank',
              shapes: shape({
                id: 2,
                x: IN,
                y: IN,
                cx: 3 * IN,
                cy: IN,
                text: para(run('Hello world')),
              }),
            },
          ],
        })
      ))
  );
  return editor!;
}

/** Link "Hello", then put a caret inside it. */
function linkAndCaret(editor: PptxEditorInstance, url = 'https://example.com/a'): void {
  const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
  act(() =>
    editor.setSelection({
      kind: 'text',
      nodeId,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 5 },
    })
  );
  act(() => void editor.exec({ type: 'setLink', url }));
  act(() =>
    editor.setSelection({
      kind: 'text',
      nodeId,
      anchor: { paragraph: 0, offset: 2 },
      focus: { paragraph: 0, offset: 2 },
    })
  );
}

/** A plain click on the slide: the press and the release at one point. */
async function clickSlide(editor: PptxEditorInstance): Promise<void> {
  const root = surfaceRootOf(editor)!;
  const at = { button: 0, clientX: 5, clientY: 5, bubbles: true };
  root.dispatchEvent(new PointerEvent('pointerdown', at));
  root.dispatchEvent(new PointerEvent('pointerup', at));
  await act(async () => void (await new Promise((r) => setTimeout(r, 40))));
}

/**
 * A plain click between the "e" and the "l" of "Hello", through the surface's own pointer path.
 * happy-dom lays nothing out, so the stage gets a 1280 px box; the point comes from the laid-out
 * caret stops (slide = textMatrix · (box + frame point), unrotated text).
 */
async function clickText(editor: PptxEditorInstance, index = 2): Promise<void> {
  const root = surfaceRootOf(editor)!;
  const stage = root.querySelector<HTMLElement>('.pptx-surface__stage')!;
  const W = 1280;
  stage.getBoundingClientRect = () =>
    ({ left: 0, top: 0, right: W, bottom: W, width: W, height: W, x: 0, y: 0 }) as DOMRect;
  const layout = editor.session!.layout(0)!;
  const item = layout.items[0]!;
  if (item.kind !== 'shape' || !item.text || !item.textMatrix) throw new Error('no text');
  const line = item.text.lines[0]!;
  const fx = item.text.box.x + (line.caretX[index]! + line.caretX[index + 1]!) / 2;
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
  await act(async () => void (await new Promise((r) => setTimeout(r, 40))));
}

const panel = () => document.querySelector<HTMLElement>('.pptx-link');

describe('PptxEditor.HyperLink', () => {
  test('the link popup stays beside the selected text while its draft changes', async () => {
    const editor = await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady}>
        <PptxEditorContent />
      </PptxEditorRoot>
    ));
    linkAndCaret(editor);
    const surface = surfaceRootOf(editor)!;
    const caret = surface.querySelector<HTMLElement>('.pptx-surface__caret')!;
    expect(caret).not.toBeNull();
    caret.getBoundingClientRect = () =>
      ({ left: 220, right: 222, top: 160, bottom: 180, width: 2, height: 20 }) as DOMRect;
    act(() => requestLinkPopup(editor));
    const popup = document.querySelector<HTMLElement>('.pptx-link-popup')!;
    expect(popup.style.left).toBe('220px');
    expect(popup.style.top).toBe('186px');
    caret.getBoundingClientRect = () =>
      ({ left: 500, right: 502, top: 360, bottom: 380, width: 2, height: 20 }) as DOMRect;
    act(
      () =>
        void fireEvent.change(document.querySelector('.pptx-link__field')!, {
          target: { value: 'https://example.com/new' },
        })
    );
    expect(popup.style.left).toBe('220px');
    expect(popup.style.top).toBe('186px');
  });

  test('a click into linked text opens reading mode with the target, Copy, Edit, and Unlink', async () => {
    const editor = await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady}>
        <PptxEditorContent />
      </PptxEditorRoot>
    ));
    linkAndCaret(editor);
    await clickSlide(editor);
    expect(panel()?.getAttribute('data-mode')).toBe('reading');
    const url = document.querySelector<HTMLAnchorElement>('.pptx-link__url')!;
    expect(url.getAttribute('href')).toBe('https://example.com/a');
    expect(url.getAttribute('target')).toBe('_blank');
    expect(url.getAttribute('rel')).toBe('noopener noreferrer');
    expect(screen.getByRole('button', { name: 'Copy link' })).toBeDefined();
    expect(screen.getByRole('button', { name: 'Remove link' })).toBeDefined();
    // The caret keeps the keys: the panel takes no focus.
    expect(document.querySelector('.pptx-link-popup')!.contains(document.activeElement)).toBe(
      false
    );
    // Edit switches to the URL field, seeded with the link.
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Edit link' })));
    expect(panel()?.getAttribute('data-mode')).toBe('editing');
    const field = screen.getByRole('textbox', { name: 'Link' }) as HTMLInputElement;
    expect(field.value).toBe('https://example.com/a');
    act(() => void fireEvent.change(field, { target: { value: 'example.org' } }));
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Apply' })));
    expect(panel()).toBe(null);
    expect(editor.snapshot().format?.link).toBe('https://example.org');
  });

  test('a click with no link, Escape, and a caret move leave or close it', async () => {
    const editor = await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady}>
        <PptxEditorContent />
      </PptxEditorRoot>
    ));
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    act(() =>
      editor.setSelection({
        kind: 'text',
        nodeId,
        anchor: { paragraph: 0, offset: 1 },
        focus: { paragraph: 0, offset: 1 },
      })
    );
    await clickSlide(editor);
    expect(panel()).toBe(null);
    linkAndCaret(editor);
    await clickSlide(editor);
    expect(panel()).not.toBe(null);
    const input = document.querySelector<HTMLElement>('.pptx-surface__input')!;
    act(() => void fireEvent.keyDown(input, { key: 'Escape' }));
    expect(panel()).toBe(null);
    // One Escape does one thing: the caret stays in the text.
    expect(editor.snapshot().selection.kind).toBe('text');
    await clickSlide(editor);
    expect(panel()).not.toBe(null);
    act(() =>
      editor.setSelection({
        kind: 'text',
        nodeId,
        anchor: { paragraph: 0, offset: 3 },
        focus: { paragraph: 0, offset: 3 },
      })
    );
    expect(panel()).toBe(null);
  });

  test('Unlink removes the link and gives focus back to the slide text', async () => {
    const editor = await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady}>
        <PptxEditorContent />
      </PptxEditorRoot>
    ));
    linkAndCaret(editor);
    await clickSlide(editor);
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Remove link' })));
    expect(panel()).toBe(null);
    expect(editor.snapshot().format?.link ?? '').toBe('');
    expect(document.activeElement).toBe(document.querySelector('.pptx-surface__input'));
  });

  test('a popups.hyperlink renderer overrides parts in place: hidden, icon, className, asChild', async () => {
    const editor = await loaded((onReady) => (
      <PptxEditorRoot
        onReady={onReady}
        popups={{
          // A renderer that shows reading mode asks for it.
          hyperlink: definePopup(
            () => (
              <PptxEditor.HyperLink className="acme-link">
                <PptxEditor.HyperLink.Unlink hidden />
                <PptxEditor.HyperLink.Copy icon={<span className="acme-copy-glyph" />} />
                <PptxEditor.HyperLink.Edit asChild>
                  <a className="acme-edit">Change</a>
                </PptxEditor.HyperLink.Edit>
              </PptxEditor.HyperLink>
            ),
            { reading: true }
          ),
        }}
      >
        <PptxEditorContent />
      </PptxEditorRoot>
    ));
    linkAndCaret(editor);
    await clickSlide(editor);
    expect(document.querySelector('.pptx-link-popup.acme-link')).not.toBe(null);
    expect(screen.queryByRole('button', { name: 'Remove link' })).toBe(null);
    expect(document.querySelectorAll('.pptx-link__copy').length).toBe(1);
    expect(document.querySelector('.pptx-link__copy .acme-copy-glyph')).not.toBe(null);
    const edit = document.querySelector<HTMLElement>('a.acme-edit')!;
    expect(edit.getAttribute('aria-label')).toBe('Edit link');
    act(() => void fireEvent.click(edit));
    expect(panel()?.getAttribute('data-mode')).toBe('editing');
  });

  test('preset={false} renders only the host parts, and useHyperlinkPopup drives a custom panel', async () => {
    function Custom() {
      const { state, setUrl, commitEdit, close } = useHyperlinkPopup();
      return (
        <PptxEditor.HyperLink preset={false}>
          <p className="acme-mode">{state.mode}</p>
          <PptxEditor.HyperLink.Fields />
          <button type="button" onClick={() => setUrl('https://acme.test/')}>
            Fill
          </button>
          <button type="button" onClick={() => commitEdit()}>
            Save
          </button>
          <button type="button" onClick={close}>
            Dismiss
          </button>
        </PptxEditor.HyperLink>
      );
    }
    const editor = await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady} popups={{ hyperlink: () => <Custom /> }}>
        <PptxEditorContent />
      </PptxEditorRoot>
    ));
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    act(() =>
      editor.setSelection({
        kind: 'text',
        nodeId,
        anchor: { paragraph: 0, offset: 0 },
        focus: { paragraph: 0, offset: 5 },
      })
    );
    act(() => requestLinkPopup(editor));
    expect(document.querySelector('.acme-mode')?.textContent).toBe('editing');
    expect(screen.queryByRole('button', { name: 'Apply' })).toBe(null);
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Fill' })));
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Save' })));
    expect(panel()).toBe(null);
    expect(editor.snapshot().format?.link).toBe('https://acme.test/');
  });

  test('a refused URL keeps the panel open and names the reason', async () => {
    const editor = await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady}>
        <PptxEditorContent />
      </PptxEditorRoot>
    ));
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    act(() =>
      editor.setSelection({
        kind: 'text',
        nodeId,
        anchor: { paragraph: 0, offset: 0 },
        focus: { paragraph: 0, offset: 5 },
      })
    );
    act(() => requestLinkPopup(editor));
    const field = screen.getByRole('textbox', { name: 'Link' });
    act(() => void fireEvent.change(field, { target: { value: 'javascript:alert(1)' } }));
    act(() => void fireEvent.keyDown(field, { key: 'Enter' }));
    expect(panel()?.getAttribute('data-mode')).toBe('editing');
    const alert = screen.getByRole('alert');
    expect(field.getAttribute('aria-describedby')).toBe(alert.id);
  });

  test('a viewer reads, copies, and opens a link; Edit and Unlink are absent', async () => {
    const editor = await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady}>
        <PptxEditorContent />
      </PptxEditorRoot>
    ));
    linkAndCaret(editor);
    act(() => editor.setMode('view'));
    act(() => editor.setSelection({ kind: 'none' }));
    // A viewer's own click on the linked word places the caret there and opens the panel.
    await clickText(editor);
    expect(editor.snapshot().selection.kind).toBe('text');
    expect(panel()?.getAttribute('data-mode')).toBe('reading');
    expect(document.querySelector('.pptx-link__url')?.getAttribute('href')).toBe(
      'https://example.com/a'
    );
    expect(document.querySelector('.pptx-link__copy')).not.toBe(null);
    expect(screen.queryByRole('button', { name: 'Edit link' })).toBe(null);
    expect(screen.queryByRole('button', { name: 'Remove link' })).toBe(null);
    // The text input takes the keys but opens no on-screen keyboard for a viewer.
    expect(surfaceRootOf(editor)!.querySelector<HTMLTextAreaElement>('textarea')!.readOnly).toBe(
      true
    );
    // A click on text after the link, with the caret still in the shape, selects the shape.
    await clickText(editor, 8);
    expect(editor.snapshot().selection.kind).toBe('shapes');
    expect(panel()).toBe(null);
  });

  test('a viewer click on text with no link selects the shape and opens nothing', async () => {
    const editor = await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady}>
        <PptxEditorContent />
      </PptxEditorRoot>
    ));
    act(() => editor.setMode('view'));
    await clickText(editor);
    expect(editor.snapshot().selection.kind).toBe('shapes');
    expect(panel()).toBe(null);
  });

  test('open() reads the link at the caret; open(null) asks for the URL field', async () => {
    function Panel() {
      const { state, open } = useHyperlinkPopup();
      return (
        <div className="probe" data-mode={state.mode}>
          <button type="button" onClick={() => open(null)}>
            field
          </button>
          <button type="button" onClick={() => open()}>
            read
          </button>
        </div>
      );
    }
    const editor = await loaded((onReady) => (
      <PptxEditorRoot
        onReady={onReady}
        popups={{ hyperlink: definePopup(Panel, { reading: true }) }}
      >
        <PptxEditorContent />
      </PptxEditorRoot>
    ));
    linkAndCaret(editor);
    await clickSlide(editor);
    const probe = () => document.querySelector<HTMLElement>('.probe');
    expect(probe()?.getAttribute('data-mode')).toBe('reading');
    act(() => screen.getByRole('button', { name: 'field' }).click());
    expect(probe()?.getAttribute('data-mode')).toBe('editing');
    act(() => screen.getByRole('button', { name: 'read' }).click());
    expect(probe()?.getAttribute('data-mode')).toBe('reading');
  });

  test('the panel keeps its place and scope after a repaint, and one Escape closes only it', async () => {
    const editor = await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady}>
        <PptxEditorViewport>
          <PptxEditorContent />
        </PptxEditorViewport>
      </PptxEditorRoot>
    ));
    linkAndCaret(editor);
    await clickSlide(editor);
    expect(panel()?.getAttribute('data-mode')).toBe('reading');
    // A repaint with no caret move (a peer's edit, a font load) rebuilds the painted caret.
    act(() => void editor.exec({ type: 'toggleSmartGuides' }));
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Edit link' })));
    const popup = document.querySelector<HTMLElement>('.pptx-link-popup')!;
    expect(panel()?.getAttribute('data-mode')).toBe('editing');
    expect(popup.closest('.pptx-editor')).not.toBe(null);
    const field = document.querySelector<HTMLInputElement>('.pptx-link__field')!;
    act(() => void fireEvent.keyDown(field, { key: 'Escape' }));
    expect(panel()).toBe(null);
    // The Escape closed the panel only: the caret stays in the text.
    expect(editor.snapshot().selection.kind).toBe('text');
  });

  test('a plain popups.hyperlink renderer opens for Ctrl/Cmd+K, not on a click', async () => {
    const seen: string[] = [];
    const editor = await loaded((onReady) => (
      <PptxEditorRoot
        onReady={onReady}
        popups={{
          hyperlink: (props) => {
            seen.push(props.mode ?? 'editing');
            return <p className="acme-form" />;
          },
        }}
      >
        <PptxEditorContent />
      </PptxEditorRoot>
    ));
    linkAndCaret(editor);
    await clickSlide(editor);
    expect(document.querySelector('.acme-form')).toBe(null);
    act(() => requestLinkPopup(editor));
    expect(document.querySelector('.acme-form')).not.toBe(null);
    expect(seen.every((m) => m === 'editing')).toBe(true);
  });

  test('edit mode keeps the target and Copy of a link, and names a new link Insert link', async () => {
    const editor = await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady}>
        <PptxEditorContent />
      </PptxEditorRoot>
    ));
    linkAndCaret(editor);
    act(() => requestLinkPopup(editor));
    expect(screen.getByRole('dialog', { name: 'Edit link' })).toBeTruthy();
    expect(document.querySelector('.pptx-link__url')?.getAttribute('href')).toBe(
      'https://example.com/a'
    );
    expect(document.querySelector('.pptx-link__copy')).not.toBe(null);
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Cancel' })));
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    act(() => void editor.exec({ type: 'setLink', url: null }));
    act(() =>
      editor.setSelection({
        kind: 'text',
        nodeId,
        anchor: { paragraph: 0, offset: 0 },
        focus: { paragraph: 0, offset: 5 },
      })
    );
    act(() => requestLinkPopup(editor));
    expect(screen.getByRole('dialog', { name: 'Insert link' })).toBeTruthy();
    expect(document.querySelector('.pptx-link__url')).toBe(null);
  });

  test('F6 moves focus into the reading panel, and a live region announces it', async () => {
    const editor = await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady}>
        <PptxEditorContent />
      </PptxEditorRoot>
    ));
    linkAndCaret(editor);
    await clickSlide(editor);
    const status = document.querySelector('.pptx-sr-only[role="status"]');
    expect(status?.textContent).toBe('Link: https://example.com/a. Press F6 for its actions.');
    const input = surfaceRootOf(editor)!.querySelector<HTMLElement>('.pptx-surface__input')!;
    input.focus();
    act(() => void fireEvent.keyDown(input, { key: 'F6' }));
    expect(document.activeElement?.closest('.pptx-link-popup')).not.toBe(null);
    expect(panel()?.getAttribute('data-mode')).toBe('reading');
  });
});

describe('PptxEditor.HyperLink link kinds', () => {
  test('Link to writes an email link and a slide link; a slide link shows Go to slide', async () => {
    const editor = await loaded((onReady) => (
      <PptxEditorRoot onReady={onReady}>
        <PptxEditorContent />
      </PptxEditorRoot>
    ));
    linkAndCaret(editor);
    await clickSlide(editor);
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Edit link' })));
    const kind = () => screen.getByRole('combobox', { name: 'Link to' }) as HTMLSelectElement;
    expect(kind().value).toBe('web');
    // Email: focus stays on Link to, and the subject is percent-encoded into the one draft.
    kind().focus();
    act(() => void fireEvent.change(kind(), { target: { value: 'email' } }));
    const address = screen.getByRole('textbox', { name: 'Email address' }) as HTMLInputElement;
    expect(document.activeElement).toBe(kind());
    act(() => void fireEvent.change(address, { target: { value: 'ana@example.com' } }));
    const subject = screen.getByRole('textbox', { name: 'Subject' });
    act(() => void fireEvent.change(subject, { target: { value: 'Q3 & plans' } }));
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Apply' })));
    expect(editor.snapshot().format?.link).toBe('mailto:ana@example.com?subject=Q3%20%26%20plans');
    // Edit again: the form starts on Email with its fields.
    await clickSlide(editor);
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Edit link' })));
    expect(kind().value).toBe('email');
    expect((screen.getByRole('textbox', { name: 'Subject' }) as HTMLInputElement).value).toBe(
      'Q3 & plans'
    );
    // Slide: the jumps and each slide by number.
    act(() => void fireEvent.change(kind(), { target: { value: 'slide' } }));
    const slide = screen.getByRole('combobox', { name: 'Slide in this presentation' });
    expect([...(slide as HTMLSelectElement).options].map((o) => o.text)).toEqual([
      'First slide',
      'Last slide',
      'Next slide',
      'Previous slide',
      'Slide 1',
    ]);
    act(() => void fireEvent.change(slide, { target: { value: '#slide-1' } }));
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Apply' })));
    expect(editor.snapshot().format?.link).toBe('#slide-1');
    // Reading mode: Go to slide, and no Copy.
    await clickSlide(editor);
    const go = screen.getByRole('button', { name: 'Go to Slide 1' });
    expect(go.textContent).toBe('Slide 1');
    expect(screen.queryByRole('button', { name: 'Copy link' })).toBe(null);
    act(() => void fireEvent.click(go));
    expect(panel()).toBe(null);
    expect(editor.snapshot().currentSlide).toBe(0);
  });
});
