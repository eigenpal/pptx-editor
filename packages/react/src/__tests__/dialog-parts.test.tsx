import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../context.ts';
import {
  BackgroundDialog,
  PageSetupDialog,
  useBackgroundDialog,
  usePageSetupDialog,
} from '../index.ts';

// Dialog named parts over a draft controller, as in DOCX Editor
// (docs/site/content/customization/dialogs.mdx): parts replace defaults in place, keep their
// wiring, and Apply commits the draft once while Cancel discards it. DOCX Editor commit adf0e178.

const IN = 914_400;

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run('Hi')) }),
      },
    ],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

const part = (name: string) => document.querySelector<HTMLElement>(`[data-pptx-part="${name}"]`);
const field = (name: string) =>
  document.querySelector<HTMLElement>(`[data-pptx-part="field"][data-pptx-field="${name}"]`);

afterEach(() => cleanup());

describe('Page setup dialog parts', () => {
  test('the default arrangement marks every part', () => {
    render(<PageSetupDialog open onClose={() => {}} onThemeChange={() => {}} />);
    for (const name of ['header', 'title', 'body', 'footer', 'error', 'apply', 'cancel'])
      expect(part(name) !== null).toBe(true);
    for (const name of ['slideSize', 'theme', 'fonts']) expect(field(name) !== null).toBe(true);
    // No substituted fonts: the Fonts field says so, and no empty Substituted fonts section.
    expect(field('substitutes') === null).toBe(true);
    expect(field('fonts')!.textContent).toContain('All fonts are available');
    expect(screen.getByRole('button', { name: 'Close' })).toBeTruthy();
  });

  test('a named part adds its class and attributes and keeps the wiring', () => {
    let closed = 0;
    render(
      <PageSetupDialog open onClose={() => (closed += 1)}>
        <PageSetupDialog.Cancel className="acme-cancel" data-testid="cancel" />
      </PageSetupDialog>
    );
    const cancel = screen.getByTestId('cancel');
    expect(cancel.classList.contains('pptx-button')).toBe(true);
    expect(cancel.classList.contains('acme-cancel')).toBe(true);
    act(() => void fireEvent.click(cancel));
    // The part keeps the dialog's Cancel wiring.
    expect(closed).toBe(1);
  });

  test('hidden removes a part, and children replace its content', () => {
    render(
      <PageSetupDialog open onClose={() => {}}>
        <PageSetupDialog.Field name="substitutes" hidden />
        <PageSetupDialog.Title>Slide size</PageSetupDialog.Title>
      </PageSetupDialog>
    );
    expect(field('substitutes') === null).toBe(true);
    expect(part('title')!.textContent).toBe('Slide size');
  });

  test('asChild puts the wiring on the host element without the packaged classes', () => {
    let closed = 0;
    render(
      <PageSetupDialog open onClose={() => (closed += 1)}>
        <PageSetupDialog.Apply asChild className="acme-ok">
          <a href="#ok">Done</a>
        </PageSetupDialog.Apply>
      </PageSetupDialog>
    );
    const ok = screen.getByText('Done');
    expect(ok.localName).toBe('a');
    expect(ok.getAttribute('data-pptx-part')).toBe('apply');
    expect(ok.className).toBe('acme-ok');
    act(() => void fireEvent.click(ok));
    expect(closed).toBe(1);
  });

  test('preset={false} renders only the host arrangement over the same draft', () => {
    const themes: string[] = [];
    function Light() {
      const { values, setValue } = usePageSetupDialog();
      return (
        <button type="button" data-testid="light" onClick={() => setValue('theme', 'light')}>
          {values.theme}
        </button>
      );
    }
    render(
      <PageSetupDialog
        open
        preset={false}
        theme="dark"
        onClose={() => {}}
        onThemeChange={(t) => themes.push(t)}
      >
        <Light />
        <PageSetupDialog.Apply>Save</PageSetupDialog.Apply>
      </PageSetupDialog>
    );
    expect(part('body') === null).toBe(true);
    expect(screen.getByTestId('light').textContent).toBe('dark');
    act(() => void fireEvent.click(screen.getByTestId('light')));
    expect(themes).toEqual([]);
    act(() => void fireEvent.click(screen.getByText('Save')));
    expect(themes).toEqual(['light']);
  });

  test('Cancel discards the appearance draft', () => {
    const themes: string[] = [];
    let closed = 0;
    render(
      <PageSetupDialog
        open
        theme="light"
        onClose={() => (closed += 1)}
        onThemeChange={(t) => themes.push(t)}
      />
    );
    act(() => void fireEvent.click(screen.getByRole('radio', { name: 'Dark' })));
    act(() => void fireEvent.click(part('cancel')!));
    expect(themes).toEqual([]);
    expect(closed).toBe(1);
  });

  test('a part outside its dialog throws for its draft hook', () => {
    function Stray() {
      usePageSetupDialog();
      return null;
    }
    const error = console.error;
    console.error = () => {};
    try {
      expect(() => render(<Stray />)).toThrow(
        'PageSetupDialog parts must be rendered inside PageSetupDialog.'
      );
    } finally {
      console.error = error;
    }
  });
});

describe('Background dialog draft', () => {
  test('a color is a draft: OK applies it once, as one undo step', async () => {
    const editor = await open();
    let closed = 0;
    render(
      <PptxEditorContext.Provider value={editor}>
        <BackgroundDialog open onClose={() => (closed += 1)} />
      </PptxEditorContext.Provider>
    );
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Dark Blue' })));
    expect(editor.snapshot().slideBackground.inherited).toBe(true);
    act(() => void fireEvent.click(part('apply')!));
    expect(editor.snapshot().slideBackground.color).toBe('#002060');
    expect(closed).toBe(1);
    act(() => void editor.exec({ type: 'undo' }));
    expect(editor.snapshot().slideBackground.inherited).toBe(true);
  });

  test('Cancel leaves the slide as it was', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <BackgroundDialog open onClose={() => {}} />
      </PptxEditorContext.Provider>
    );
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Dark Blue' })));
    act(() => void fireEvent.click(part('cancel')!));
    expect(editor.snapshot().slideBackground.inherited).toBe(true);
    expect(editor.snapshot().canUndo).toBe(false);
  });

  test('OK with nothing changed adds no undo step', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <BackgroundDialog open onClose={() => {}} />
      </PptxEditorContext.Provider>
    );
    act(() => void fireEvent.click(part('apply')!));
    expect(editor.snapshot().canUndo).toBe(false);
  });

  test('a host field writes the same draft through useBackgroundDialog', async () => {
    const editor = await open();
    function Navy() {
      const { setValue } = useBackgroundDialog();
      return (
        <button
          type="button"
          data-testid="navy"
          onClick={() => setValue('background', { kind: 'solid', color: '#000080' })}
        />
      );
    }
    render(
      <PptxEditorContext.Provider value={editor}>
        <BackgroundDialog open onClose={() => {}}>
          <BackgroundDialog.Field name="solid">
            <Navy />
          </BackgroundDialog.Field>
        </BackgroundDialog>
      </PptxEditorContext.Provider>
    );
    act(() => void fireEvent.click(screen.getByTestId('navy')));
    act(() => void fireEvent.click(part('apply')!));
    expect(editor.snapshot().slideBackground.color).toBe('#000080');
  });

  test('viewing mode disables OK with the engine reason path', async () => {
    const editor = await open();
    act(() => editor.setMode('view'));
    render(
      <PptxEditorContext.Provider value={editor}>
        <BackgroundDialog open onClose={() => {}} />
      </PptxEditorContext.Provider>
    );
    expect((part('apply') as HTMLButtonElement).disabled).toBe(true);
  });
});
