import { afterEach, beforeEach, describe, expect, spyOn, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import { useState, type ReactNode } from 'react';
import {
  createPptxEditor,
  resetDevWarnings,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../context.ts';
import { BackgroundDialog, PageSetupDialog, useBackgroundDialog } from '../index.ts';

// The dialog frame's keys, focus, and backdrop, the Background dialog's paths, and the
// development warnings for parts a dialog cannot use.

async function open(background?: string): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({ slides: [{ layout: 'blank', background, shapes: '' }] });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

const part = (name: string) => document.querySelector<HTMLElement>(`[data-pptx-part="${name}"]`);
const dialog = (kind: string) =>
  document.querySelector<HTMLDialogElement>(`[data-pptx-dialog="${kind}"]`);

function withEditor(editor: PptxEditorInstance, ui: ReactNode) {
  return render(<PptxEditorContext.Provider value={editor}>{ui}</PptxEditorContext.Provider>);
}

let warn: ReturnType<typeof spyOn>;
beforeEach(() => {
  resetDevWarnings();
  warn = spyOn(console, 'warn').mockImplementation(() => {});
});
afterEach(() => {
  warn.mockRestore();
  cleanup();
});
const warnings = (): string[] => warn.mock.calls.map((c: unknown[]) => String(c[0]));

describe('DialogFrame', () => {
  test('Escape closes the dialog; Tab and Shift+Tab wrap inside it', async () => {
    const editor = await open();
    let closed = 0;
    withEditor(editor, <PageSetupDialog open onClose={() => (closed += 1)} />);
    const frame = dialog('pageSetup')!;
    // The header's close button is the first control, as in the other dialogs.
    const close = screen.getByRole('button', { name: 'Close' });
    const apply = part('apply')!;
    apply.focus();
    fireEvent.keyDown(apply, { key: 'Tab' });
    expect(document.activeElement).toBe(close);
    fireEvent.keyDown(close, { key: 'Tab', shiftKey: true });
    expect(document.activeElement).toBe(apply);
    act(() => void fireEvent.click(close));
    expect(closed).toBe(1);
    fireEvent.keyDown(frame, { key: 'Escape' });
    expect(closed).toBe(2);
  });

  test('a dialog without a field to type in starts on OK, so Return applies', async () => {
    const editor = await open();
    withEditor(editor, <PageSetupDialog open onClose={() => {}} />);
    expect(document.activeElement).toBe(part('apply'));
  });

  test('only a press and a release on the backdrop close it', async () => {
    const editor = await open();
    let closed = 0;
    withEditor(editor, <PageSetupDialog open onClose={() => (closed += 1)} />);
    const frame = dialog('pageSetup')!;
    // A press on a control, released on the dialog element: a drag, not a dismissal.
    fireEvent.mouseDown(part('title')!, { clientX: 5, clientY: 5 });
    fireEvent.click(frame, { clientX: -10, clientY: -10 });
    expect(closed).toBe(0);
    fireEvent.mouseDown(frame, { clientX: -10, clientY: -10 });
    fireEvent.click(frame, { clientX: -10, clientY: -10 });
    expect(closed).toBe(1);
  });

  test('Escape in the custom gradient editor closes only that editor and keeps the draft', async () => {
    const editor = await open();
    let closed = 0;
    withEditor(editor, <BackgroundDialog open onClose={() => (closed += 1)} />);
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Dark Blue' })));
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Custom gradient…' })));
    const inner = dialog('gradient')!;
    const control = inner.querySelector('select')!;
    fireEvent.keyDown(control, { key: 'Escape' });
    expect(closed).toBe(0);
    act(() => void fireEvent.click(part('apply')!));
    expect(editor.snapshot().slideBackground.color).toBe('#002060');
  });
});

describe('DialogFrame, round 3', () => {
  test("the gradient editor's own cancel event leaves the Background dialog open", async () => {
    const editor = await open();
    let closed = 0;
    withEditor(editor, <BackgroundDialog open onClose={() => (closed += 1)} />);
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Custom gradient…' })));
    const inner = dialog('gradient')!;
    act(() => void fireEvent(inner, new Event('cancel', { cancelable: true })));
    expect(closed).toBe(0);
    expect(dialog('background') !== null).toBe(true);
  });

  test('a host title on a part stays when the default sets none', async () => {
    const editor = await open();
    withEditor(
      editor,
      <BackgroundDialog open onClose={() => {}}>
        <BackgroundDialog.Apply title="Apply background" />
      </BackgroundDialog>
    );
    expect(part('apply')!.getAttribute('title')).toBe('Apply background');
  });
});

describe('Background dialog paths', () => {
  test('a theme swatch and a gradient preset show as chosen', async () => {
    const editor = await open();
    withEditor(editor, <BackgroundDialog open onClose={() => {}} />);
    const theme = document.querySelector<HTMLElement>('.pptx-swatches--theme button')!;
    act(() => void fireEvent.click(theme));
    expect(theme.getAttribute('aria-pressed')).toBe('true');
    const preset = screen.getAllByRole('button', { name: /gradient/i })[0]!;
    act(() => void fireEvent.click(preset));
    expect(preset.getAttribute('aria-pressed')).toBe('true');
    expect(theme.getAttribute('aria-pressed')).toBe('false');
  });

  test('Reset picks the layout background in the draft, and OK applies it', async () => {
    const editor = await open(
      '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
    );
    withEditor(editor, <BackgroundDialog open onClose={() => {}} />);
    expect(editor.snapshot().slideBackground.inherited).toBe(false);
    act(() => void fireEvent.click(part('reset')!));
    expect(editor.snapshot().slideBackground.inherited).toBe(false);
    act(() => void fireEvent.click(part('apply')!));
    expect(editor.snapshot().slideBackground.inherited).toBe(true);
  });

  test('Apply to all slides without a pick explains why, on the checkbox too', async () => {
    const editor = await open(
      '<p:bg><p:bgPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
    );
    withEditor(editor, <BackgroundDialog open onClose={() => {}} />);
    const box = screen.getByRole('checkbox', { name: 'Apply to all slides' });
    act(() => void fireEvent.click(box));
    act(() => void fireEvent.click(part('apply')!));
    const error = part('error')!;
    expect(error.textContent).toBe(
      "To apply a fill to all slides, pick a color, gradient, or picture. The slide's current fill cannot be copied."
    );
    expect(box.getAttribute('aria-invalid')).toBe('true');
    expect(box.getAttribute('aria-describedby')).toBe(error.id);
    expect(editor.snapshot().canUndo).toBe(false);
  });

  test('a chosen picture shows as chosen and OK applies it', async () => {
    const editor = await open();
    withEditor(editor, <BackgroundDialog open onClose={() => {}} />);
    const input = document.querySelector<HTMLInputElement>('input[type="file"]')!;
    const png = new File([new Uint8Array([0x89, 0x50, 0x4e, 0x47])], 'p.png', {
      type: 'image/png',
    });
    await act(async () => {
      fireEvent.change(input, { target: { files: [png] } });
      await new Promise((r) => setTimeout(r, 0));
    });
    expect(screen.getByText('Picture chosen')).toBeTruthy();
    act(() => void fireEvent.click(part('apply')!));
    expect(editor.snapshot().slideBackground.kind).toBe('picture');
  });

  test('a refused commit shows the engine reason, translated, through the hook too', async () => {
    const editor = await open();
    act(() => editor.setMode('view'));
    let errors: Readonly<Record<string, string | undefined>> = {};
    function HostApply() {
      const dialogState = useBackgroundDialog();
      errors = dialogState.errors;
      return <button type="button" data-testid="host-apply" onClick={dialogState.apply} />;
    }
    withEditor(
      editor,
      <BackgroundDialog open onClose={() => {}}>
        <HostApply />
      </BackgroundDialog>
    );
    expect(part('apply')!.getAttribute('title')).toBe('The presentation is open for viewing');
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Dark Blue' })));
    act(() => void fireEvent.click(screen.getByTestId('host-apply')));
    expect(part('error')!.textContent).toBe('The presentation is open for viewing');
    expect(errors.form).toBe('The presentation is open for viewing');
  });
});

describe('Background dialog, round 2', () => {
  const red =
    '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>';

  test('re-picking the slide fill applies it to every slide', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const bytes = buildPptx({
      slides: [
        { layout: 'blank', background: red, shapes: '' },
        { layout: 'blank', shapes: '' },
      ],
    });
    expect((await editor.load(bytes)).ok).toBe(true);
    withEditor(editor, <BackgroundDialog open onClose={() => {}} />);
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Red' })));
    act(() => void fireEvent.click(screen.getByRole('checkbox', { name: 'Apply to all slides' })));
    act(() => void fireEvent.click(part('apply')!));
    act(() => editor.selectSlide(1));
    expect(editor.snapshot().slideBackground).toMatchObject({ color: '#FF0000', inherited: false });
  });

  test('a pick applies even when another edit changed the slide after the dialog opened', async () => {
    const editor = await open(red);
    withEditor(editor, <BackgroundDialog open onClose={() => {}} />);
    act(
      () =>
        void editor.exec({
          type: 'setSlideBackground',
          background: { kind: 'solid', color: '#0000FF' },
        })
    );
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Red' })));
    act(() => void fireEvent.click(part('apply')!));
    expect(editor.snapshot().slideBackground.color).toBe('#FF0000');
  });

  test('a dialog the host renders closes when the editor opens another document', async () => {
    const editor = await open();
    let closed = 0;
    withEditor(editor, <BackgroundDialog open onClose={() => (closed += 1)} />);
    await act(async () => void (await editor.load(buildPptx({ slides: [{ shapes: '' }] }))));
    expect(closed).toBe(1);
  });

  test('Return applies from the checkbox; Reset keeps focus and stays enabled', async () => {
    const editor = await open(red);
    let closed = 0;
    withEditor(editor, <BackgroundDialog open onClose={() => (closed += 1)} />);
    const reset = part('reset') as HTMLButtonElement;
    reset.focus();
    act(() => void fireEvent.click(reset));
    expect(reset.disabled).toBe(false);
    expect(document.activeElement).toBe(reset);
    expect(screen.getByRole('status').textContent).toBe('Follows the layout background');
    const box = screen.getByRole('checkbox', { name: 'Apply to all slides' });
    act(() => void fireEvent.keyDown(box, { key: 'Enter' }));
    expect(closed).toBe(1);
    expect(editor.snapshot().slideBackground.inherited).toBe(true);
  });

  test('a picture that finishes reading after a later pick is dropped', async () => {
    const editor = await open();
    withEditor(editor, <BackgroundDialog open onClose={() => {}} />);
    let finish: (buffer: ArrayBuffer) => void = () => {};
    const slow = {
      type: 'image/png',
      arrayBuffer: () => new Promise<ArrayBuffer>((resolve) => (finish = resolve)),
    } as unknown as File;
    const input = document.querySelector<HTMLInputElement>('input[type="file"]')!;
    act(() => void fireEvent.change(input, { target: { files: [slow] } }));
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Dark Blue' })));
    await act(async () => {
      finish(new Uint8Array([0x89, 0x50]).buffer);
      await new Promise((r) => setTimeout(r, 0));
    });
    act(() => void fireEvent.click(part('apply')!));
    expect(editor.snapshot().slideBackground.color).toBe('#002060');
  });

  test('a hex fill presses one swatch only', async () => {
    const editor = await open(
      '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFC000"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
    );
    withEditor(editor, <BackgroundDialog open onClose={() => {}} />);
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Orange' })));
    expect(document.querySelectorAll('[aria-pressed="true"]').length).toBe(1);
  });
});

describe('Background dialog, round 3', () => {
  test('Reset applies on OK even when the slide followed the layout at opening', async () => {
    const editor = await open();
    withEditor(editor, <BackgroundDialog open onClose={() => {}} />);
    act(
      () =>
        void editor.exec({
          type: 'setSlideBackground',
          background: { kind: 'solid', color: '#FF0000' },
        })
    );
    act(() => void fireEvent.click(part('reset')!));
    act(() => void fireEvent.click(part('apply')!));
    expect(editor.snapshot().slideBackground.inherited).toBe(true);
  });

  test('the Apply to all slides error clears when the user picks a fill', async () => {
    const editor = await open(
      '<p:bg><p:bgPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
    );
    withEditor(editor, <BackgroundDialog open onClose={() => {}} />);
    const box = screen.getByRole('checkbox', { name: 'Apply to all slides' });
    act(() => void fireEvent.click(box));
    act(() => void fireEvent.click(part('apply')!));
    expect(box.getAttribute('aria-invalid')).toBe('true');
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Dark Blue' })));
    expect(box.getAttribute('aria-invalid')).toBeNull();
    expect(part('error')!.textContent).toBe('');
  });

  test('a part outside its dialog, and a part with children but no default, say so', async () => {
    const editor = await open();
    withEditor(
      editor,
      <>
        <PageSetupDialog.Title className="stray" />
        <PageSetupDialog open onClose={() => {}} preset={false}>
          <PageSetupDialog.Field name="theme">
            <span data-testid="picker" />
          </PageSetupDialog.Field>
        </PageSetupDialog>
      </>
    );
    expect(screen.getByTestId('picker')).toBeTruthy();
    const text = warnings().join('\n');
    expect(text).toContain('PageSetupDialog.Title is outside its dialog');
    expect(text).toContain(
      'PageSetupDialog.Field name="theme" matches no part the dialog renders now, so it renders only its children'
    );
  });
});

describe('Background dialog, round 4', () => {
  function Revert() {
    const s = useBackgroundDialog();
    const [opening] = useState(() => s.values.background);
    return (
      <button data-testid="revert" onClick={() => s.setValue('background', opening)}>
        revert
      </button>
    );
  }

  test('a host that sets the opening value back after a pick applies nothing', async () => {
    const editor = await open(
      '<p:bg><p:bgPr><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
    );
    withEditor(
      editor,
      <BackgroundDialog open onClose={() => {}}>
        <Revert />
      </BackgroundDialog>
    );
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Blue' })));
    act(() => void fireEvent.click(screen.getByTestId('revert')));
    act(() => void fireEvent.click(part('apply')!));
    expect(editor.snapshot().canUndo).toBe(false);
  });

  test('a swatch grid is one Tab stop, and the arrows move by column and row', async () => {
    const editor = await open();
    withEditor(editor, <BackgroundDialog open onClose={() => {}} />);
    const grid = document.querySelector<HTMLElement>('.pptx-swatches--theme')!;
    const all = [...grid.querySelectorAll<HTMLButtonElement>('.pptx-swatch')];
    expect(all.filter((b) => b.tabIndex === 0).length).toBe(1);
    all[0]!.focus();
    act(() => void fireEvent.keyDown(all[0]!, { key: 'ArrowDown' }));
    expect(document.activeElement).toBe(all[10]);
    act(() => void fireEvent.keyDown(all[10]!, { key: 'ArrowRight' }));
    expect(document.activeElement).toBe(all[11]);
    expect(all[11]!.tabIndex).toBe(0);
  });

  test('the focused swatch keeps the Tab stop when another swatch becomes the chosen one', async () => {
    const editor = await open();
    let setValue: ReturnType<typeof useBackgroundDialog>['setValue'] | null = null;
    function Grab() {
      setValue = useBackgroundDialog().setValue;
      return null;
    }
    withEditor(
      editor,
      <BackgroundDialog open onClose={() => {}}>
        <Grab />
      </BackgroundDialog>
    );
    const grid = document.querySelector<HTMLElement>('.pptx-swatches--theme')!;
    const all = [...grid.querySelectorAll<HTMLButtonElement>('.pptx-swatch')];
    all[0]!.focus();
    act(() => void fireEvent.keyDown(all[0]!, { key: 'ArrowRight' }));
    expect(document.activeElement).toBe(all[1]);
    act(() => setValue!('background', { kind: 'solid', color: { theme: 'accent6' } }));
    expect(all[1]!.tabIndex).toBe(0);
  });

  test('a misused asChild on the owner does not claim its parts render nothing', async () => {
    const editor = await open();
    withEditor(
      editor,
      <BackgroundDialog open onClose={() => {}}>
        <BackgroundDialog.Footer asChild>
          <b />
          <i />
        </BackgroundDialog.Footer>
        <BackgroundDialog.Apply className="host-apply" />
      </BackgroundDialog>
    );
    expect(document.querySelector('.host-apply') !== null).toBe(true);
    expect(warnings().join('\n')).not.toContain('replaces a default inside');
  });

  test('a part inside a part that brings its own children says it renders nothing', async () => {
    const editor = await open();
    withEditor(
      editor,
      <BackgroundDialog open onClose={() => {}}>
        <BackgroundDialog.Body>
          <p />
        </BackgroundDialog.Body>
        <BackgroundDialog.Field name="solid" className="x" />
      </BackgroundDialog>
    );
    expect(warnings().join('\n')).toContain(
      'BackgroundDialog.Field name="solid" replaces a default inside BackgroundDialog.Body'
    );
  });
});

describe('Page setup appearance', () => {
  test('arrow keys move the choice, one radio is a tab stop, and OK calls the latest callback', async () => {
    const editor = await open();
    const calls: string[] = [];
    function Host() {
      const [label, setLabel] = useState('first');
      return (
        <>
          <button type="button" data-testid="relabel" onClick={() => setLabel('second')} />
          <PageSetupDialog
            open
            onClose={() => {}}
            theme="light"
            onThemeChange={() => calls.push(label)}
          />
        </>
      );
    }
    withEditor(editor, <Host />);
    const light = screen.getByRole('radio', { name: 'Light' });
    const dark = screen.getByRole('radio', { name: 'Dark' });
    expect([light.tabIndex, dark.tabIndex]).toEqual([0, -1]);
    act(() => void fireEvent.keyDown(light, { key: 'ArrowRight' }));
    expect(dark.getAttribute('aria-checked')).toBe('true');
    expect(document.activeElement).toBe(dark);
    act(() => void fireEvent.click(screen.getByTestId('relabel')));
    act(() => void fireEvent.click(part('apply')!));
    expect(calls).toEqual(['second']);
  });

  test('Tab skips the unchecked radio and stays in the dialog; Return on a radio applies', async () => {
    const editor = await open();
    let closed = 0;
    withEditor(
      editor,
      <PageSetupDialog open onClose={() => (closed += 1)} theme="dark" onThemeChange={() => {}} />
    );
    const dark = screen.getByRole('radio', { name: 'Dark' });
    const close = screen.getByRole('button', { name: 'Close' });
    const apply = part('apply')!;
    // One Tab stop per radio group: the checked radio only.
    expect(screen.getByRole('radio', { name: 'Light' }).tabIndex).toBe(-1);
    expect(dark.tabIndex).toBe(0);
    apply.focus();
    fireEvent.keyDown(apply, { key: 'Tab' });
    expect(document.activeElement).toBe(close);
    fireEvent.keyDown(close, { key: 'Tab', shiftKey: true });
    expect(document.activeElement).toBe(apply);
    act(() => void fireEvent.keyDown(dark, { key: 'Enter' }));
    expect(closed).toBe(1);
  });
});

describe('dialog part warnings', () => {
  test('a part the dialog does not render, a duplicate, a nested part, and a bad asChild warn', async () => {
    const editor = await open();
    withEditor(
      editor,
      <PageSetupDialog open onClose={() => {}}>
        <PageSetupDialog.Field name="theme" className="x" />
        <PageSetupDialog.Title className="a" />
        <PageSetupDialog.Title className="b" />
        <div>
          <PageSetupDialog.Cancel />
        </div>
        <PageSetupDialog.Apply asChild>Save</PageSetupDialog.Apply>
      </PageSetupDialog>
    );
    const text = warnings().join('\n');
    expect(text).toContain('PageSetupDialog.Field name="theme" matches no part');
    expect(text).toContain('PageSetupDialog.Title appears more than once');
    expect(text).toContain('PageSetupDialog.Cancel is inside another element');
    expect(text).toContain('PageSetupDialog.Apply has asChild but not exactly one element child');
    // The misused asChild falls back to the packaged OK button.
    expect(part('apply')!.textContent).toBe('OK');
    expect(part('title')!.classList.contains('b')).toBe(true);
  });
});
