import {
  afterAll,
  afterEach,
  beforeAll,
  beforeEach,
  describe,
  expect,
  spyOn,
  test,
} from 'bun:test';
import { act, cleanup, fireEvent, render, screen, waitFor } from '@testing-library/react';
import { useState, type ReactNode } from 'react';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../context.ts';
import { PptxEditorRoot } from '../PptxEditorRoot.tsx';
import { SidePanel } from '../panes.tsx';
import { resetDevWarnings } from '@pptx-editor.dev/core/editor';
import {
  PageSetupDialog,
  PptxEditor,
  Toolbar,
  definePopup,
  useEditorDialog,
  type PageSetupDialogProps,
  type PptxEditorPopups,
} from '../index.ts';

// The customization ladder and the popups registry, as in DOCX Editor
// (docs/site/content/customization/toolbar.mdx and dialogs.mdx). DOCX Editor commit adf0e178.

const IN = 914_400;

const deckBytes = (text = 'Hi'): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run(text)) }),
      },
    ],
  });

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(deckBytes())).ok).toBe(true);
  // The text box is selected: the bar shows its text, shape, and arrange groups.
  editor.setSelection({
    kind: 'shapes',
    nodeIds: [editor.session!.layout(0)!.items[0]!.element.nodeId],
  });
  return editor;
}

/** Select the first two characters of the slide's text box. */
function selectText(editor: PptxEditorInstance): void {
  const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
  act(() =>
    editor.setSelection({
      kind: 'text',
      nodeId,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 2 },
    })
  );
}

// happy-dom has no layout: every element measures 9 pixels wide, and the toolbar has room
// so no group moves into the More menu unless a test narrows it.
let barWidth = 100_000;
const width = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth');
const offset = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'offsetWidth');
beforeAll(() => {
  Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
    configurable: true,
    get: () => barWidth,
  });
  Object.defineProperty(HTMLElement.prototype, 'offsetWidth', { configurable: true, get: () => 9 });
});
afterAll(() => {
  if (width) Object.defineProperty(HTMLElement.prototype, 'clientWidth', width);
  if (offset) Object.defineProperty(HTMLElement.prototype, 'offsetWidth', offset);
});
let warn: ReturnType<typeof spyOn<Console, 'warn'>>;
beforeEach(() => {
  resetDevWarnings();
  warn = spyOn(console, 'warn').mockImplementation(() => {});
});
afterEach(() => {
  cleanup();
  warn.mockRestore();
  barWidth = 100_000;
});

const warned = (fragment: string): boolean =>
  warn.mock.calls.some((call) => String(call[0]).includes(fragment));

function renderToolbar(editor: PptxEditorInstance, toolbar: ReactNode) {
  return render(<PptxEditorContext.Provider value={editor}>{toolbar}</PptxEditorContext.Provider>);
}

const slot = (container: HTMLElement, id: string) =>
  container.querySelector<HTMLElement>(`[data-slot="${id}"]`);

describe('Toolbar ladder, step 1: class names and state attributes', () => {
  test('className on the bar, and data-disabled on a disabled button', async () => {
    const editor = await open();
    const { container } = renderToolbar(editor, <Toolbar className="acme-bar" />);
    expect(container.querySelector('.pptx-toolbar.acme-bar') !== null).toBe(true);
    const undo = slot(container, 'history.undo')!;
    // aria-disabled, not disabled: the control stays reachable and gives its reason.
    expect(undo.getAttribute('aria-disabled')).toBe('true');
    expect(undo.hasAttribute('data-disabled')).toBe(true);
    const reason = document.getElementById(undo.getAttribute('aria-describedby')!);
    expect(reason?.textContent).toBe('Nothing to undo');
  });

  test('a toggle reports aria-pressed and data-active', async () => {
    const editor = await open();
    const { container } = renderToolbar(editor, <Toolbar />);
    selectText(editor);
    expect(slot(container, 'text.bold')!.getAttribute('aria-pressed')).toBe('false');
    expect(slot(container, 'text.bold')!.hasAttribute('data-active')).toBe(false);
    act(() => void editor.exec({ type: 'toggleBold' }));
    expect(slot(container, 'text.bold')!.getAttribute('aria-pressed')).toBe('true');
    expect(slot(container, 'text.bold')!.hasAttribute('data-active')).toBe(true);
  });
});

describe('Toolbar ladder, steps 2 and 3: icon and asChild', () => {
  test('Toolbar.Slot replaces the icon and adds a class name', async () => {
    const editor = await open();
    const { container } = renderToolbar(
      editor,
      <Toolbar>
        <Toolbar.Slot
          slotId="text.bold"
          className="acme-bold"
          icon={<span data-testid="bold-icon">B</span>}
        />
      </Toolbar>
    );
    const bold = slot(container, 'text.bold')!;
    expect(bold.classList.contains('acme-bold')).toBe(true);
    expect(bold.querySelector('[data-testid="bold-icon"]') !== null).toBe(true);
    expect(bold.querySelector('svg') === null).toBe(true);
  });

  test('asChild merges into the host element: classes join and both handlers run', async () => {
    const editor = await open();
    expect(editor.exec({ type: 'addSlide' }).ok).toBe(true);
    let hostClicks = 0;
    renderToolbar(
      editor,
      <Toolbar>
        <Toolbar.Slot slotId="history.undo" asChild>
          <a className="acme-link" data-testid="undo-link" onClick={() => (hostClicks += 1)}>
            Undo
          </a>
        </Toolbar.Slot>
      </Toolbar>
    );
    const link = screen.getByTestId('undo-link');
    expect(link.tagName).toBe('A');
    expect(link.getAttribute('role')).toBe('button');
    expect(link.classList.contains('acme-link')).toBe(true);
    expect(link.classList.contains('pptx-toolbar__button')).toBe(true);
    expect(link.getAttribute('data-slot')).toBe('history.undo');
    expect(link.getAttribute('aria-label')).toBeTruthy();
    const slides = editor.session!.info.slides.length;
    act(() => void fireEvent.click(link));
    expect(hostClicks).toBe(1);
    expect(editor.session!.info.slides.length).toBe(slides - 1);
  });

  test('a disabled asChild control says so and does not run its command', async () => {
    const editor = await open();
    renderToolbar(
      editor,
      <Toolbar>
        <Toolbar.Slot slotId="history.undo" asChild>
          <span data-testid="undo">Undo</span>
        </Toolbar.Slot>
      </Toolbar>
    );
    const undo = screen.getByTestId('undo');
    expect(undo.getAttribute('aria-disabled')).toBe('true');
    expect(undo.hasAttribute('data-disabled')).toBe(true);
    const revision = editor.snapshot().revision;
    act(() => void fireEvent.click(undo));
    expect(editor.snapshot().revision).toBe(revision);
  });

  test('picker controls take className and state, and warn on icon and asChild', async () => {
    const editor = await open();
    act(() => editor.setSelection({ kind: 'none' }));
    // A hand-built bar keeps the pickers while nothing is selected.
    const { container } = renderToolbar(
      editor,
      <Toolbar preset={false}>
        <Toolbar.Slot slotId="font.family" className="acme-family" icon={<span />} />
        <Toolbar.Slot slotId="font.size" className="acme-size" />
        <Toolbar.Slot slotId="paragraph.align" className="acme-align" asChild>
          <span />
        </Toolbar.Slot>
      </Toolbar>
    );
    for (const [id, cls] of [
      ['font.family', 'acme-family'],
      ['font.size', 'acme-size'],
      ['paragraph.align', 'acme-align'],
    ] as const) {
      const el = slot(container, id)!;
      expect(el.classList.contains(cls)).toBe(true);
      // Nothing is selected, so every one of them is disabled.
      expect(el.hasAttribute('data-disabled')).toBe(true);
    }
    expect(warned('"font.family" renders a picker')).toBe(true);
    expect(warned('"paragraph.align" renders a picker')).toBe(true);
  });
});

describe('Toolbar ladder, step 4: Slot, Group, and preset', () => {
  test('Toolbar.Slot children replace the control in its place', async () => {
    const editor = await open();
    const { container } = renderToolbar(
      editor,
      <Toolbar>
        <Toolbar.Slot slotId="text.italic">
          <button type="button" data-testid="my-italic" />
        </Toolbar.Slot>
      </Toolbar>
    );
    expect(slot(container, 'text.italic') === null).toBe(true);
    expect(screen.getByTestId('my-italic').closest('[data-group="text"]') !== null).toBe(true);
  });

  test('hidden removes a control, and its command keeps working', async () => {
    const editor = await open();
    const { container } = renderToolbar(
      editor,
      <Toolbar>
        <Toolbar.Slot slotId="text.bold" hidden />
      </Toolbar>
    );
    expect(slot(container, 'text.bold') === null).toBe(true);
    selectText(editor);
    expect(editor.can({ type: 'toggleBold' }).ok).toBe(true);
    let result: ReturnType<PptxEditorInstance['exec']> | undefined;
    act(() => void (result = editor.exec({ type: 'toggleBold' })));
    expect(result?.ok).toBe(true);
    expect(editor.snapshot().format?.bold).toBe(true);
  });

  test('Toolbar.Group hides a built-in group or adds controls to it', async () => {
    const editor = await open();
    const { container } = renderToolbar(
      editor,
      <Toolbar>
        <Toolbar.Group id="history" hidden />
        <Toolbar.Group id="text">
          <button type="button" data-testid="extra" />
        </Toolbar.Group>
      </Toolbar>
    );
    expect(container.querySelector('[data-group="history"]') === null).toBe(true);
    expect(screen.getByTestId('extra').closest('[data-group="text"]') !== null).toBe(true);
  });

  test('a host group comes after the built-in groups; label wins over labelKey', async () => {
    const editor = await open();
    const { container } = renderToolbar(
      editor,
      <Toolbar>
        <Toolbar.Group id="acme" label="Acme" labelKey="toolbar.groups.text">
          <button type="button" data-testid="ask" />
        </Toolbar.Group>
        <Toolbar.Group id="keyed" labelKey="toolbar.groups.text">
          <button type="button" />
        </Toolbar.Group>
        <Toolbar.Group id="plain">
          <button type="button" />
        </Toolbar.Group>
      </Toolbar>
    );
    const groups = [...container.querySelectorAll('[data-group]')].map((g) =>
      g.getAttribute('data-group')
    );
    expect(groups.slice(-3)).toEqual(['acme', 'keyed', 'plain']);
    const label = (id: string) =>
      container.querySelector(`[data-group="${id}"]`)!.getAttribute('aria-label');
    expect(label('acme')).toBe('Acme');
    expect(label('keyed')).toBe('Text');
    expect(label('plain')).toBe('plain');
    expect(
      container.querySelector('[data-group="acme"]')!.contains(screen.getByTestId('ask'))
    ).toBe(true);
  });

  test('a host group moves into the More menu when the bar is narrow', async () => {
    barWidth = 60;
    const editor = await open();
    const { container } = renderToolbar(
      editor,
      <Toolbar>
        <Toolbar.Group id="acme" label="Acme">
          <button type="button" data-testid="ask" />
        </Toolbar.Group>
      </Toolbar>
    );
    const bar = container.querySelector('.pptx-toolbar')!;
    expect(bar.querySelector('[data-group="acme"]') === null).toBe(true);
    const more = screen.getByRole('button', { name: 'More' });
    act(() => void fireEvent.click(more));
    const overflow = document.querySelector('.pptx-toolbar__overflow')!;
    expect(overflow.querySelector('[data-group="acme"] [data-testid="ask"]') !== null).toBe(true);
  });

  test('preset={false} renders only the children; hidden still applies', async () => {
    const editor = await open();
    const { container } = renderToolbar(
      editor,
      <Toolbar preset={false}>
        <Toolbar.Slot slotId="history.undo" />
        <Toolbar.Slot slotId="history.redo" hidden />
        <Toolbar.Group id="mine" label="Mine">
          <Toolbar.Slot slotId="text.bold" />
        </Toolbar.Group>
      </Toolbar>
    );
    const slots = [...container.querySelectorAll('[data-slot]')].map((b) =>
      b.getAttribute('data-slot')
    );
    expect(slots).toEqual(['history.undo', 'text.bold']);
    expect(container.querySelector('[data-group="mine"] [data-slot="text.bold"]') !== null).toBe(
      true
    );
  });
});

describe('Toolbar development warnings', () => {
  test('an unknown slot id renders nothing and warns', async () => {
    const editor = await open();
    renderToolbar(
      editor,
      <Toolbar preset={false}>
        <Toolbar.Slot slotId={'text.nope' as 'text.bold'} />
      </Toolbar>
    );
    expect(warned('"text.nope" is not a toolbar slot id')).toBe(true);
  });

  test('a duplicate slot warns and the last one wins', async () => {
    const editor = await open();
    const { container } = renderToolbar(
      editor,
      <Toolbar>
        <Toolbar.Slot slotId="text.bold" className="first" />
        <Toolbar.Slot slotId="text.bold" className="second" />
      </Toolbar>
    );
    expect(warned('Toolbar.Slot "text.bold" appears more than once')).toBe(true);
    expect(slot(container, 'text.bold')!.classList.contains('second')).toBe(true);
  });

  test('a slot that is not on the preset bar goes after the groups, as DOCX Editor', async () => {
    const editor = await open();
    const { container } = renderToolbar(
      editor,
      <Toolbar>
        <Toolbar.Slot slotId="view.smartGuides" />
      </Toolbar>
    );
    expect(container.querySelector('[data-fixed] [data-slot="view.smartGuides"]') !== null).toBe(
      true
    );
    expect(warned('Toolbar.Slot "view.smartGuides" is not part of the preset toolbar')).toBe(true);
  });

  test('asChild without exactly one element child warns and renders the packaged control', async () => {
    const editor = await open();
    const { container } = renderToolbar(
      editor,
      <Toolbar>
        {/* Two children: invalid for asChild. */}
        <Toolbar.Slot slotId="text.bold" asChild>
          {[<span key="a" />, <span key="b" />] as unknown as React.ReactElement}
        </Toolbar.Slot>
      </Toolbar>
    );
    expect(warned('"text.bold" has asChild but not exactly one element child')).toBe(true);
    expect(container.querySelector('button[data-slot="text.bold"]') !== null).toBe(true);
  });
});

function Opener({ dialog }: { dialog: 'pageSetup' | 'background' }) {
  const handle = useEditorDialog(dialog);
  return (
    <button
      type="button"
      data-testid="opener"
      data-available={String(handle.isAvailable)}
      data-open={String(handle.isOpen)}
      onClick={() => handle.open()}
    />
  );
}

function MyPageSetup({ open, onClose }: PageSetupDialogProps) {
  return open ? (
    <div data-testid="my-page-setup">
      <button type="button" data-testid="my-close" onClick={onClose} />
    </div>
  ) : null;
}

const dialogEl = (name: string) =>
  document.querySelector<HTMLDialogElement>(`[data-pptx-dialog="${name}"]`);

describe('popups registry', () => {
  const mount = (dialog: 'pageSetup' | 'background', popups?: PptxEditorPopups) =>
    render(
      <PptxEditorRoot {...(popups ? { popups } : {})}>
        <Opener dialog={dialog} />
      </PptxEditorRoot>
    );

  test('without an entry, the packaged dialog opens', () => {
    mount('pageSetup');
    expect(dialogEl('pageSetup') === null).toBe(true);
    act(() => void fireEvent.click(screen.getByTestId('opener')));
    expect(dialogEl('pageSetup') !== null).toBe(true);
    expect(screen.getByTestId('opener').getAttribute('data-open')).toBe('true');
  });

  test('a definePopup entry replaces the packaged dialog and gets open and onClose', () => {
    mount('pageSetup', { pageSetup: definePopup(MyPageSetup) });
    act(() => void fireEvent.click(screen.getByTestId('opener')));
    expect(screen.getByTestId('my-page-setup')).toBeTruthy();
    expect(dialogEl('pageSetup') === null).toBe(true);
    act(() => void fireEvent.click(screen.getByTestId('my-close')));
    expect(screen.queryByTestId('my-page-setup') === null).toBe(true);
    expect(screen.getByTestId('opener').getAttribute('data-open')).toBe('false');
  });

  test('a render callback forwards options to the packaged dialog', () => {
    mount('pageSetup', {
      pageSetup: (props) => <PageSetupDialog {...props} theme="dark" onThemeChange={() => {}} />,
    });
    act(() => void fireEvent.click(screen.getByTestId('opener')));
    expect(dialogEl('pageSetup')!.textContent).toContain('Appearance');
  });

  test('false turns off automatic rendering and reports the dialog unavailable', () => {
    mount('background', { background: false });
    const opener = screen.getByTestId('opener');
    expect(opener.getAttribute('data-available')).toBe('false');
    act(() => void fireEvent.click(opener));
    expect(document.querySelector('[data-pptx-dialog]') === null).toBe(true);
    expect(opener.getAttribute('data-open')).toBe('false');
  });

  test('Escape and a backdrop click close the dialog, and focus returns to the opener', () => {
    mount('pageSetup');
    const opener = screen.getByTestId('opener');
    for (const dismiss of [
      (d: HTMLDialogElement) => fireEvent(d, new Event('cancel', { cancelable: true })),
      // A press and a release on the backdrop: outside the dialog's box.
      (d: HTMLDialogElement) => {
        fireEvent.mouseDown(d, { clientX: -10, clientY: -10 });
        fireEvent.click(d, { clientX: -10, clientY: -10 });
      },
    ]) {
      opener.focus();
      act(() => void fireEvent.click(opener));
      const dialog = dialogEl('pageSetup')!;
      expect(dialog !== null).toBe(true);
      act(() => void dismiss(dialog));
      expect(dialogEl('pageSetup') === null).toBe(true);
      expect(document.activeElement).toBe(opener);
    }
  });

  test('a dialog closes when the editor opens another document', async () => {
    let editor: PptxEditorInstance | null = null;
    render(
      <PptxEditorRoot onReady={(e) => (editor = e)}>
        <Opener dialog="background" />
      </PptxEditorRoot>
    );
    await act(async () => void (await editor!.load(deckBytes())));
    act(() => void fireEvent.click(screen.getByTestId('opener')));
    expect(dialogEl('background') !== null).toBe(true);
    await act(async () => void (await editor!.load(deckBytes('Other'))));
    expect(dialogEl('background') === null).toBe(true);
  });

  test('outside PptxEditor.Root, a dialog is unavailable', () => {
    render(<Opener dialog="pageSetup" />);
    expect(screen.getByTestId('opener').getAttribute('data-available')).toBe('false');
  });

  test('popups.background = false removes the Format panel opener', async () => {
    const view = (popups?: PptxEditorPopups) => {
      let editor: PptxEditorInstance | null = null;
      const r = render(
        <PptxEditorRoot
          document={deckBytes()}
          onReady={(e) => (editor = e)}
          {...(popups ? { popups } : {})}
        >
          <SidePanel />
        </PptxEditorRoot>
      );
      return { ...r, editor: () => editor };
    };
    const on = view();
    await waitFor(() => expect(on.editor()?.snapshot().status).toBe('ready'));
    const button = await screen.findByRole('button', { name: 'Format Background…' });
    act(() => void fireEvent.click(button));
    expect(dialogEl('background') !== null).toBe(true);
    on.unmount();

    const off = view({ background: false });
    await waitFor(() => expect(off.editor()?.snapshot().status).toBe('ready'));
    expect(screen.queryByRole('button', { name: 'Format Background…' }) === null).toBe(true);
  });

  test('<PptxEditor> forwards popups to its root', async () => {
    function Probe() {
      const [shown, setShown] = useState(false);
      const pageSetup = useEditorDialog('pageSetup');
      return (
        <button
          type="button"
          data-testid="probe"
          data-shown={String(shown)}
          onClick={() => (setShown(true), pageSetup.open())}
        />
      );
    }
    render(
      <PptxEditor popups={{ pageSetup: definePopup(MyPageSetup) }}>
        <Probe />
      </PptxEditor>
    );
    act(() => void fireEvent.click(screen.getByTestId('probe')));
    expect(screen.getByTestId('my-page-setup')).toBeTruthy();
  });
});
