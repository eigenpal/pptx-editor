// The Paragraph dialog: the Line Spacing menu's last row opens it through the popups registry
// (there is no toolbar button), its fields seed from the selection with
// a blank for a mixed value, OK applies the changed fields as one undo step, a value out of
// range names the field, and `useParagraphFormat` reads and writes the same command.

import { afterAll, afterEach, beforeAll, describe, expect, spyOn, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import { resetDevWarnings, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { attr, child, elements, find, xmlPartOf, type Element } from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContent, PptxEditorRoot, PptxEditorViewport } from '../PptxEditorRoot.tsx';
import {
  ParagraphDialog,
  Toolbar,
  ToolbarGroup,
  ToolbarSlot,
  definePopup,
  useParagraphDialog,
  useParagraphFormat,
  type ParagraphDialogProps,
  type PptxEditorPopups,
  type UseParagraphFormatReturn,
} from '../index.ts';

const IN = 914_400;
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          x: IN,
          y: IN,
          cx: 4 * IN,
          cy: 2 * IN,
          text:
            para(run('one'), '<a:pPr><a:spcBef><a:spcPts val="600"/></a:spcBef></a:pPr>') +
            para(run('two')),
        }),
      },
    ],
  });

const width = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth');
beforeAll(() => {
  // happy-dom has no layout: give the toolbar room so no group moves into the More menu.
  Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
    configurable: true,
    get: () => 100_000,
  });
});
afterAll(() => {
  if (width) Object.defineProperty(HTMLElement.prototype, 'clientWidth', width);
});
afterEach(() => cleanup());

async function mount(popups?: PptxEditorPopups): Promise<PptxEditorInstance> {
  let editor: PptxEditorInstance | null = null;
  render(
    <PptxEditorRoot onReady={(e) => (editor = e)} {...(popups ? { popups } : {})}>
      <Toolbar />
    </PptxEditorRoot>
  );
  await act(async () => void (await editor!.load(deck())));
  return editor!;
}

function selectParagraphs(editor: PptxEditorInstance, from = 0, to = 1): void {
  const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
  act(() =>
    editor.setSelection({
      kind: 'text',
      nodeId,
      anchor: { paragraph: from, offset: 0 },
      focus: { paragraph: to, offset: 3 },
    })
  );
}
const selectAll = (editor: PptxEditorInstance) => selectParagraphs(editor, 0, 1);

const lineSpacing = () =>
  document.querySelector<HTMLButtonElement>('[data-slot="list.lineSpacing"]')!;
/** Open the dialog from the menu: Line Spacing > Line Spacing Options…. */
function openDialog(): void {
  act(() => void fireEvent.click(lineSpacing()));
  act(() => void fireEvent.click(screen.getByText('Line spacing options…')));
}
const dialog = () => document.querySelector<HTMLDialogElement>('[data-pptx-dialog="paragraph"]');
const input = (name: string) => dialog()!.querySelector<HTMLInputElement>(`[name="${name}"]`)!;
const select = (name: string) => dialog()!.querySelector<HTMLSelectElement>(`[name="${name}"]`)!;
const ok = () => dialog()!.querySelector<HTMLButtonElement>('[data-pptx-part="apply"]')!;
const change = (el: HTMLElement, value: string) =>
  act(() => void fireEvent.change(el, { target: { value } }));

function paragraphs(editor: PptxEditorInstance): Element[] {
  const part = editor.session!.info.slides[0]!.partName;
  const root = xmlPartOf(editor.session!.store.package, part)!.root;
  return elements(find(root, 'p:txBody')!).filter((c) => c.local === 'p');
}
const spacingVal = (p: Element, local: string): string | undefined => {
  const spacing = find(child(p, 'a:pPr')!, `a:${local}`);
  return spacing ? attr(elements(spacing)[0], 'val') : undefined;
};

describe('Paragraph dialog', () => {
  test('has no toolbar button; Line Spacing Options… opens it seeded from the selection', async () => {
    const editor = await mount();
    expect(document.querySelector('[data-slot="paragraph.dialog"]') === null).toBe(true);
    selectParagraphs(editor, 0, 0);
    openDialog();
    expect(dialog() !== null).toBe(true);
    expect(input('spaceBefore').value).toBe('6');
    expect(select('spaceBeforeUnit').value).toBe('pt');
    // The fixture master's text styles state 90 % line spacing, as a common default theme does.
    expect(select('lineRule').value).toBe('multiple');
    expect(input('lineAt').value).toBe('0.9');
    expect(select('alignment').value).toBe('left');
    expect([...select('alignment').options].map((o) => o.textContent)).toEqual([
      'Left',
      'Centered',
      'Right',
      'Justified',
      'Distributed',
    ]);
    // Header close and footer Cancel have different names.
    expect(screen.getByRole('button', { name: 'Close' })).toBeTruthy();
  });

  test('a mixed value opens blank, and a value typed there applies to every paragraph', async () => {
    const editor = await mount();
    selectAll(editor);
    openDialog();
    expect(input('spaceBefore').value).toBe('');
    expect(input('spaceBefore').placeholder).toBe('Mixed');
    change(input('spaceBefore'), '6');
    change(input('spaceAfter'), '0.5');
    change(select('spaceAfterUnit'), 'lines');
    // A number typed before the unit change is meant in the new unit: it stays.
    expect(input('spaceAfter').value).toBe('0.5');
    change(select('special'), 'hanging');
    // Hanging starts at half an inch.
    expect(input('specialBy').value).toBe('0.5');
    change(input('specialBy'), '0.25');
    act(() => void fireEvent.click(ok()));
    expect(dialog() === null).toBe(true);
    for (const p of paragraphs(editor)) {
      expect(spacingVal(p, 'spcBef')).toBe('600');
      expect(spacingVal(p, 'spcAft')).toBe('50000');
      expect(attr(child(p, 'a:pPr'), 'indent')).toBe(String(-IN / 4));
    }
    act(() => void editor.exec({ type: 'undo' }));
    expect(spacingVal(paragraphs(editor)[0]!, 'spcAft')).toBeUndefined();
    expect(editor.snapshot().canUndo).toBe(false);
  });

  test('Line Spacing choices set At; Multiple never takes points as lines', async () => {
    const editor = await mount();
    selectParagraphs(editor, 0, 0);
    openDialog();
    change(select('lineRule'), 'exact');
    expect(input('lineAt').value).toBe('12');
    change(select('lineRule'), 'double');
    // A disabled At shows no stale number.
    expect(input('lineAt').value).toBe('');
    change(select('lineRule'), 'multiple');
    expect(input('lineAt').value).toBe('2');
    act(() => void fireEvent.click(ok()));
    expect(spacingVal(paragraphs(editor)[0]!, 'lnSpc')).toBe('200000');
  });

  test('Cancel and an unchanged OK change nothing', async () => {
    const editor = await mount();
    selectParagraphs(editor, 0, 0);
    openDialog();
    change(input('spaceBefore'), '30');
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Cancel' })));
    expect(dialog() === null).toBe(true);
    openDialog();
    expect(input('spaceBefore').value).toBe('6');
    act(() => void fireEvent.click(ok()));
    expect(dialog() === null).toBe(true);
    expect(editor.snapshot().canUndo).toBe(false);
  });

  test('a value out of range names the field and its range, and marks the field', async () => {
    const editor = await mount();
    selectParagraphs(editor, 0, 0);
    openDialog();
    change(input('spaceBefore'), '5000');
    act(() => void fireEvent.click(ok()));
    expect(dialog() !== null).toBe(true);
    const alert = dialog()!.querySelector('[role="alert"]')!;
    expect(alert.textContent).toBe('Before must be a number from 0 to 1584 pt.');
    expect(input('spaceBefore').getAttribute('aria-invalid')).toBe('true');
    expect(input('spaceBefore').getAttribute('aria-describedby')).toBe(alert.id);
    expect(document.activeElement).toBe(input('spaceBefore'));
    expect(editor.snapshot().canUndo).toBe(false);
    // The dialog root carries the state for host styles.
    expect(dialog()!.hasAttribute('data-invalid')).toBe(true);
    // An edit clears the error, and the state with it.
    change(input('spaceBefore'), '12');
    expect(dialog()!.hasAttribute('data-invalid')).toBe(false);
    change(input('spaceBefore'), '5000');
    act(() => void fireEvent.click(ok()));
    // A cleared field where the paragraphs agreed is not a value either.
    change(input('spaceBefore'), '');
    act(() => void fireEvent.click(ok()));
    expect(dialog()!.querySelector('[role="alert"]')!.textContent).toContain('Before must be');
  });

  test('a press that starts in the dialog and ends on the backdrop does not close it', async () => {
    const editor = await mount();
    selectParagraphs(editor, 0, 0);
    openDialog();
    act(() => void fireEvent.mouseDown(input('spaceBefore'), { clientX: 5, clientY: 5 }));
    act(() => void fireEvent.click(dialog()!, { clientX: -10, clientY: -10 }));
    expect(dialog() !== null).toBe(true);
    act(() => void fireEvent.mouseDown(dialog()!, { clientX: -10, clientY: -10 }));
    act(() => void fireEvent.click(dialog()!, { clientX: -10, clientY: -10 }));
    expect(dialog() === null).toBe(true);
  });

  test('a host className goes on the dialog element, after pptx-dialog', async () => {
    const editor = await mount({
      paragraph: definePopup((props: ParagraphDialogProps) => (
        <ParagraphDialog {...props} className="host-dialog" />
      )),
    });
    selectAll(editor);
    openDialog();
    expect(dialog()!.classList.contains('pptx-dialog')).toBe(true);
    expect(dialog()!.classList.contains('host-dialog')).toBe(true);
  });

  test('named parts: a hidden field leaves, a host part restyles, and a host field reads the draft', async () => {
    let read = null as string | null;
    function SpaceBeforeEcho() {
      read = useParagraphDialog().values.spaceBefore;
      return <span data-testid="echo" />;
    }
    const editor = await mount({
      paragraph: definePopup((props: ParagraphDialogProps) => (
        <ParagraphDialog {...props}>
          <ParagraphDialog.Field name="indentation" hidden />
          <ParagraphDialog.Apply className="host-ok" />
          <SpaceBeforeEcho />
        </ParagraphDialog>
      )),
    });
    selectParagraphs(editor, 0, 0);
    openDialog();
    expect(dialog()!.querySelector('[name="indentLeft"]') === null).toBe(true);
    expect(input('spaceBefore') !== null).toBe(true);
    expect(ok().classList.contains('host-ok')).toBe(true);
    expect(screen.getByTestId('echo')).toBeTruthy();
    expect(read).toBe('6');
    change(input('spaceBefore'), '12');
    expect(read).toBe('12');
    act(() => void fireEvent.click(ok()));
    expect(dialog() === null).toBe(true);
    expect(spacingVal(paragraphs(editor)[0]!, 'spcBef')).toBe('1200');
  });

  test('a host control replaces one field and writes the draft; asChild there warns', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    function LineAtBox() {
      const { values, setValue } = useParagraphDialog();
      return (
        <input
          data-testid="host-line-at"
          value={values.lineAt}
          onChange={(e) => setValue('lineAt', e.target.value)}
        />
      );
    }
    const editor = await mount({
      paragraph: definePopup((props: ParagraphDialogProps) => (
        <ParagraphDialog {...props}>
          <ParagraphDialog.Field name="spaceBefore" id="host-before" />
          <ParagraphDialog.Field name="lineAt" asChild>
            <LineAtBox />
          </ParagraphDialog.Field>
        </ParagraphDialog>
      )),
    });
    selectParagraphs(editor, 0, 0);
    openDialog();
    expect(dialog()!.querySelector('#host-before') !== null).toBe(true);
    expect(dialog()!.querySelector('[name="lineAt"]') === null).toBe(true);
    expect(
      warn.mock.calls.some((c) => String(c[0]).includes('no single element for asChild'))
    ).toBe(true);
    warn.mockRestore();
    change(input('spaceBefore'), '12');
    act(() => void fireEvent.click(ok()));
    expect(spacingVal(paragraphs(editor)[0]!, 'spcBef')).toBe('1200');
  });

  test('in viewing mode OK is disabled and says why', async () => {
    const editor = await mount();
    selectParagraphs(editor, 0, 0);
    openDialog();
    act(() => editor.setMode('view'));
    expect(ok().disabled).toBe(true);
    expect(ok().title).toBe('The presentation is open for viewing');
  });

  test('named parts per field: one field hides, one field restyles, its section stays', async () => {
    const editor = await mount({
      paragraph: definePopup((props: ParagraphDialogProps) => (
        <ParagraphDialog {...props}>
          <ParagraphDialog.Field name="spaceAfter" hidden />
          <ParagraphDialog.Field name="lineAt" className="host-line-at" />
        </ParagraphDialog>
      )),
    });
    selectParagraphs(editor, 0, 0);
    openDialog();
    expect(input('spaceAfter') === null).toBe(true);
    expect(dialog()!.querySelector('[name="spaceAfterUnit"]') === null).toBe(true);
    expect(input('spaceBefore') !== null).toBe(true);
    expect(input('lineAt')!.closest('label')!.classList.contains('host-line-at')).toBe(true);
    expect(input('lineAt')!.closest('label')!.classList.contains('pptx-field')).toBe(true);
    change(input('spaceBefore'), '12');
    act(() => void fireEvent.click(ok()));
    expect(spacingVal(paragraphs(editor)[0]!, 'spcBef')).toBe('1200');
  });

  test('popups.paragraph replaces the dialog, and false removes its menu row', async () => {
    let seen: ParagraphDialogProps | null = null;
    function Mine(props: ParagraphDialogProps) {
      seen = props;
      return props.open ? <div data-testid="mine" /> : null;
    }
    const editor = await mount({ paragraph: definePopup(Mine) });
    selectAll(editor);
    openDialog();
    expect(screen.getByTestId('mine')).toBeTruthy();
    expect(dialog() === null).toBe(true);
    expect(seen!.open).toBe(true);
    cleanup();
    const again = await mount({ paragraph: false });
    selectAll(again);
    act(() => void fireEvent.click(lineSpacing()));
    // The row stays, disabled with the reason, as in the menus.
    expect(
      screen
        .getByText('Line spacing options…')
        .closest('[role="menuitem"]')!
        .getAttribute('aria-disabled')
    ).toBe('true');
    expect(screen.getByText('1.5')).toBeTruthy();
  });
});

describe('Paragraph dialog routes and focus', () => {
  async function mountCanvas(popups?: PptxEditorPopups): Promise<PptxEditorInstance> {
    let editor: PptxEditorInstance | null = null;
    render(
      <PptxEditorRoot onReady={(e) => (editor = e)} {...(popups ? { popups } : {})}>
        <PptxEditorViewport>
          <Toolbar />
          <PptxEditorContent />
        </PptxEditorViewport>
      </PptxEditorRoot>
    );
    await act(async () => void (await editor!.load(deck())));
    return editor!;
  }
  const textMenu = () =>
    act(
      () =>
        void fireEvent.keyDown(document.querySelector('.pptx-canvas')!, {
          key: 'F10',
          shiftKey: true,
        })
    );

  test('the text context menu opens it, and focus returns to the text on close', async () => {
    const editor = await mountCanvas();
    selectParagraphs(editor, 0, 0);
    textMenu();
    act(() => void fireEvent.click(screen.getByRole('menuitem', { name: 'Paragraph…' })));
    expect(dialog() !== null).toBe(true);
    // The dialog renders inside the editor's scope, so the theme tokens apply.
    expect(dialog()!.closest('.pptx-editor') !== null).toBe(true);
    act(() => void fireEvent.keyDown(dialog()!, { key: 'Escape' }));
    expect(dialog() === null).toBe(true);
    expect(document.activeElement?.classList.contains('pptx-surface__input')).toBe(true);
  });

  test('popups.paragraph false disables the context-menu row with the reason', async () => {
    const editor = await mountCanvas({ paragraph: false });
    selectParagraphs(editor, 0, 0);
    textMenu();
    const row = screen.getByText('Paragraph…').closest<HTMLElement>('[role="menuitem"]')!;
    expect(row.getAttribute('aria-disabled')).toBe('true');
    act(() => void fireEvent.click(row));
    expect(dialog() === null).toBe(true);
  });

  test('the first field takes focus, also when the dialog element held it first', async () => {
    // A browser's showModal focuses the <dialog> itself while its fields are still empty.
    const proto = HTMLDialogElement.prototype as { showModal?: () => void };
    const real = Object.getOwnPropertyDescriptor(proto, 'showModal');
    proto.showModal = function (this: HTMLDialogElement) {
      this.setAttribute('open', '');
      this.tabIndex = -1;
      this.focus();
    };
    try {
      const editor = await mount();
      selectParagraphs(editor, 0, 0);
      openDialog();
      const active = document.activeElement as HTMLElement | null;
      expect(active === dialog()).toBe(false);
      expect(dialog()!.contains(active)).toBe(true);
      expect(active?.matches('input, select')).toBe(true);
    } finally {
      if (real) Object.defineProperty(proto, 'showModal', real);
      else delete proto.showModal;
    }
  });

  test('Tab wraps inside the dialog in both directions', async () => {
    const editor = await mount();
    selectParagraphs(editor, 0, 0);
    openDialog();
    const focusable = [
      ...dialog()!.querySelectorAll<HTMLElement>('button, input:not([disabled]), select'),
    ].filter((el) => !el.hasAttribute('disabled'));
    focusable.at(-1)!.focus();
    act(() => void fireEvent.keyDown(focusable.at(-1)!, { key: 'Tab' }));
    expect(document.activeElement).toBe(focusable[0]!);
    act(() => void fireEvent.keyDown(focusable[0]!, { key: 'Tab', shiftKey: true }));
    expect(document.activeElement).toBe(focusable.at(-1)!);
  });

  test('a host places the slot in a group: it opens the dialog through the registry', async () => {
    let editor: PptxEditorInstance | null = null;
    render(
      <PptxEditorRoot onReady={(e) => (editor = e)}>
        <Toolbar preset={false}>
          <ToolbarGroup id="mine" label="Mine">
            <ToolbarSlot slotId="paragraph.dialog" className="host-paragraph" />
          </ToolbarGroup>
        </Toolbar>
      </PptxEditorRoot>
    );
    await act(async () => void (await editor!.load(deck())));
    selectParagraphs(editor!, 0, 0);
    const button = document.querySelector<HTMLButtonElement>('[data-slot="paragraph.dialog"]')!;
    expect(button.classList.contains('host-paragraph')).toBe(true);
    act(() => void fireEvent.click(button));
    expect(dialog() !== null).toBe(true);
  });

  test('a disabled field shows no "Mixed"; a unit change keeps a blank field blank', async () => {
    const editor = await mount();
    selectAll(editor);
    openDialog();
    // Special is (none) for both paragraphs: By is disabled and has no placeholder.
    expect(input('specialBy').disabled).toBe(true);
    expect(input('specialBy').placeholder).toBe('');
    // Before is mixed: a look at the unit list writes nothing.
    change(select('spaceBeforeUnit'), 'lines');
    change(select('spaceBeforeUnit'), 'pt');
    expect(input('spaceBefore').value).toBe('');
    act(() => void fireEvent.click(ok()));
    expect(spacingVal(paragraphs(editor)[1]!, 'spcBef')).toBeUndefined();
  });

  test('field names carry the unit for screen readers', async () => {
    const editor = await mount();
    selectParagraphs(editor, 0, 0);
    openDialog();
    expect(screen.getByRole('spinbutton', { name: 'Before text (in)' })).toBeTruthy();
  });
});

describe('useParagraphFormat', () => {
  test('reads the mixed selection and applies one command', async () => {
    let hook: UseParagraphFormatReturn | null = null;
    function Probe() {
      hook = useParagraphFormat();
      return null;
    }
    let editor: PptxEditorInstance | null = null;
    render(
      <PptxEditorRoot onReady={(e) => (editor = e)}>
        <Probe />
      </PptxEditorRoot>
    );
    await act(async () => void (await editor!.load(deck())));
    expect(hook!.format).toBeNull();
    expect(hook!.isEnabled).toBe(false);
    selectAll(editor!);
    expect(hook!.isEnabled).toBe(true);
    expect(hook!.format!.disagrees.spaceBefore).toBe(true);
    expect(hook!.format!.spaceBeforePt).toBeNull();
    let applied = false;
    act(() => void (applied = hook!.apply({ spaceBeforePt: 12, alignment: 'center' })));
    expect(applied).toBe(true);
    expect(hook!.format!.spaceBeforePt).toBe(12);
    expect(hook!.format!.alignment).toBe('center');
    expect(hook!.apply({ spaceBeforePt: 5000 })).toBe(false);
    act(() => void editor!.exec({ type: 'undo' }));
    expect(hook!.format!.disagrees.spaceBefore).toBe(true);
  });
});
