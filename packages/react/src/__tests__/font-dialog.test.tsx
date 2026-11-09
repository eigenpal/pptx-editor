// The Font group's new controls: superscript and subscript toggles, the Character Spacing and
// Change Case menus, and the Font dialog with its launcher.

import { afterAll, afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { attr, child, elements, find, textOf, xmlPartOf } from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContent, PptxEditorRoot, PptxEditorViewport } from '../PptxEditorRoot.tsx';
import {
  ColumnsDialog,
  FontDialog,
  Toolbar,
  definePopup,
  useFontDialog,
  useEditorDialog,
  useFontFormat,
  type ColumnsDialogProps,
  type FontDialogProps,
  type UseFontFormatReturn,
} from '../index.ts';
import { tooltipText } from './tooltip-text.ts';

/** A host button that opens the Font dialog, as Format > Font… does. */
function OpenFont() {
  const font = useEditorDialog('font');
  return (
    <button type="button" data-testid="open-font" onClick={() => font.open()}>
      Font…
    </button>
  );
}
const openFont = () => act(() => void fireEvent.click(screen.getByTestId('open-font')));

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
          text: para(run('make it so')),
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

/** The font and paragraph slots the default bar leaves out: the Font dialog and Format menu set them. */
const OFF_BAR = [
  'script.super',
  'script.sub',
  'text.characterSpacing',
  'text.changeCase',
  'paragraph.columns',
  'paragraph.textDirection',
] as const;

async function mount(
  popups?: { font?: false | ((p: FontDialogProps) => unknown); columns?: false },
  hostDialog = false
) {
  let editor: PptxEditorInstance | null = null;
  const root = (
    <PptxEditorRoot onReady={(e) => (editor = e)} {...(popups ? { popups: popups as never } : {})}>
      <PptxEditorViewport>
        <Toolbar>
          {/* Off the default bar; a host adds them back. */}
          {OFF_BAR.map((id) => (
            <Toolbar.Slot key={id} slotId={id} />
          ))}
        </Toolbar>
        <OpenFont />
        {/* The slide text's input, which the toolbar hands the caret back to. */}
        <textarea className="pptx-surface__input" aria-label="slide text" />
      </PptxEditorViewport>
    </PptxEditorRoot>
  );
  // A host that shows the whole editor in a modal.
  render(hostDialog ? <div role="dialog">{root}</div> : root);
  await act(async () => void (await editor!.load(deck())));
  const nodeId = editor!.session!.layout(0)!.items[0]!.element.nodeId;
  act(() =>
    editor!.setSelection({
      kind: 'text',
      nodeId,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 4 },
    })
  );
  return editor!;
}

/** Mount with the real slide surface, so `editor.focus()` picks the caret or the slide. */
async function mountWithSurface(bytes: Uint8Array = deck()) {
  let editor: PptxEditorInstance | null = null;
  render(
    <PptxEditorRoot onReady={(e) => (editor = e)}>
      <PptxEditorViewport>
        <Toolbar />
        <PptxEditorContent />
      </PptxEditorViewport>
    </PptxEditorRoot>
  );
  await act(async () => void (await editor!.load(bytes)));
  return editor!;
}

const slot = (id: string) => document.querySelector<HTMLButtonElement>(`[data-slot="${id}"]`)!;
const dialog = () => document.querySelector<HTMLDialogElement>('[data-pptx-dialog="font"]');
const field = <T extends Element>(name: string) =>
  dialog()!.querySelector<T & HTMLInputElement>(`[name="${name}"]`)!;
function firstRun(editor: PptxEditorInstance) {
  const part = editor.session!.info.slides[0]!.partName;
  const body = find(xmlPartOf(editor.session!.store.package, part)!.root, 'p:txBody')!;
  const p = elements(body).find((c) => c.local === 'p')!;
  return elements(p).filter((c) => c.local === 'r')[0]!;
}
const rPr = (editor: PptxEditorInstance, name: string) =>
  attr(child(firstRun(editor), 'a:rPr'), name);

describe('script toggles and Font group menus', () => {
  test('superscript and subscript are pressed toggles', async () => {
    const editor = await mount();
    expect(slot('script.super').getAttribute('aria-pressed')).toBe('false');
    act(() => void fireEvent.click(slot('script.super')));
    expect(rPr(editor, 'baseline')).toBe('30000');
    expect(slot('script.super').getAttribute('aria-pressed')).toBe('true');
    act(() => void fireEvent.click(slot('script.sub')));
    expect(rPr(editor, 'baseline')).toBe('-25000');
    expect(slot('script.super').getAttribute('aria-pressed')).toBe('false');
    expect(slot('script.sub').getAttribute('aria-pressed')).toBe('true');
  });

  test('Character Spacing and Change Case run their option', async () => {
    const editor = await mount();
    act(() => void fireEvent.click(slot('text.characterSpacing')));
    act(() => void fireEvent.click(screen.getByText('Very loose')));
    expect(rPr(editor, 'spc')).toBe('600');
    act(() => void fireEvent.click(slot('text.changeCase')));
    act(() => void fireEvent.click(screen.getByText('UPPERCASE')));
    expect(textOf(child(firstRun(editor), 'a:t'))).toBe('MAKE');
  });
});

describe('Paragraph group menus', () => {
  test('Text Direction and Columns set the shape and check the current option', async () => {
    const editor = await mount();
    const bodyPr = () => {
      const part = editor.session!.info.slides[0]!.partName;
      return find(xmlPartOf(editor.session!.store.package, part)!.root, 'a:bodyPr')!;
    };
    act(() => void fireEvent.click(slot('paragraph.textDirection')));
    act(() => void fireEvent.click(screen.getByText('Rotate all text 270°')));
    expect(attr(bodyPr(), 'vert')).toBe('vert270');
    act(() => void fireEvent.click(slot('paragraph.columns')));
    act(() => void fireEvent.click(screen.getByText('Two columns')));
    expect(attr(bodyPr(), 'numCol')).toBe('2');
    act(() => void fireEvent.click(slot('paragraph.columns')));
    const two = screen.getByText('Two columns').closest('[role]')!;
    expect(two.getAttribute('aria-checked')).toBe('true');
  });

  test('More columns sets the count and spacing as one undo step', async () => {
    const editor = await mount();
    const bodyPr = () => {
      const part = editor.session!.info.slides[0]!.partName;
      return find(xmlPartOf(editor.session!.store.package, part)!.root, 'a:bodyPr')!;
    };
    act(() => void fireEvent.click(slot('paragraph.columns')));
    act(() => void fireEvent.click(screen.getByText('More columns…')));
    const columns = document.querySelector('[data-pptx-dialog="columns"]')!;
    const input = (name: string) => columns.querySelector<HTMLInputElement>(`[name="${name}"]`)!;
    expect([input('count').value, input('spacing').value]).toEqual(['1', '0']);
    act(() => void fireEvent.change(input('count'), { target: { value: '4' } }));
    act(() => void fireEvent.change(input('spacing'), { target: { value: '0.5' } }));
    act(() => void fireEvent.click(columns.querySelector('[data-pptx-part="apply"]')!));
    expect(document.querySelector('[data-pptx-dialog="columns"]') === null).toBe(true);
    expect([attr(bodyPr(), 'numCol'), attr(bodyPr(), 'spcCol')]).toEqual(['4', '457200']);
    act(() => void editor.exec({ type: 'undo' }));
    expect(attr(bodyPr(), 'numCol')).toBeUndefined();
    expect(editor.snapshot().canUndo).toBe(false);
  });

  test('Columns dialog parts per field: the spacing hides, the count restyles', async () => {
    await mount({
      columns: definePopup((props: ColumnsDialogProps) => (
        <ColumnsDialog {...props}>
          <ColumnsDialog.Field name="spacing" hidden />
          <ColumnsDialog.Field name="count" className="host-count" />
        </ColumnsDialog>
      )),
    } as never);
    act(() => void fireEvent.click(slot('paragraph.columns')));
    act(() => void fireEvent.click(screen.getByText('More columns…')));
    const columns = document.querySelector('[data-pptx-dialog="columns"]')!;
    expect(columns.querySelector('[name="spacing"]') === null).toBe(true);
    const count = columns.querySelector<HTMLInputElement>('[name="count"]')!;
    expect(count.closest('label')!.classList.contains('host-count')).toBe(true);
  });

  test('More columns marks a count out of range, and false hides its row', async () => {
    await mount();
    act(() => void fireEvent.click(slot('paragraph.columns')));
    act(() => void fireEvent.click(screen.getByText('More columns…')));
    const columns = document.querySelector('[data-pptx-dialog="columns"]')!;
    const count = columns.querySelector<HTMLInputElement>('[name="count"]')!;
    act(() => void fireEvent.change(count, { target: { value: '17' } }));
    act(() => void fireEvent.click(columns.querySelector('[data-pptx-part="apply"]')!));
    expect(columns.querySelector('[role="alert"]')?.textContent).toContain('1 to 16');
    expect(count.getAttribute('aria-invalid')).toBe('true');
    cleanup();
    await mount({ columns: false });
    act(() => void fireEvent.click(slot('paragraph.columns')));
    // The row stays, disabled with the reason, as in the menus.
    expect(
      screen.getByText('More columns…').closest('[role="menuitem"]')!.getAttribute('aria-disabled')
    ).toBe('true');
    expect(screen.getByText('Two columns')).toBeTruthy();
  });
});

describe('toolbar menus', () => {
  test('a pick gives the caret back to the slide text', async () => {
    const editor = await mount();
    const text = document.querySelector<HTMLTextAreaElement>('.pptx-surface__input')!;
    act(() => void fireEvent.click(slot('text.characterSpacing')));
    act(() => void fireEvent.click(screen.getByText('Very loose')));
    expect(rPr(editor, 'spc')).toBe('600');
    expect(document.activeElement).toBe(text);
  });

  test('a menu opened with the pointer takes the keys; Escape closes it to its trigger', async () => {
    await mount();
    const text = document.querySelector<HTMLTextAreaElement>('.pptx-surface__input')!;
    act(() => text.focus());
    let reached = false;
    const listen = (e: KeyboardEvent) => void (e.key === 'Escape' && (reached = true));
    text.addEventListener('keydown', listen);
    act(() => void fireEvent.click(slot('text.characterSpacing')));
    const menu = document.querySelector<HTMLElement>('.pptx-menu')!;
    // The menu has focus, on its checked row, so a typed key or an arrow does not reach the
    // slide text.
    const checked = menu.querySelector('[role="menuitemradio"][aria-checked="true"]')!;
    expect(document.activeElement === checked).toBe(true);
    act(() => void fireEvent.keyDown(checked, { key: 'Escape' }));
    expect(document.querySelector('.pptx-menu') === null).toBe(true);
    // The slide text never sees the key, so text editing does not end.
    expect(reached).toBe(false);
    expect(document.activeElement).toBe(slot('text.characterSpacing'));
    text.removeEventListener('keydown', listen);
  });

  describe('focus that leaves a menu opened by the pointer', () => {
    /** Open Character Spacing with the pointer, move focus to `at`, and press Escape there. */
    function escapeAt(at: HTMLElement): { stopped: boolean; open: boolean } {
      act(() => void fireEvent.click(slot('text.characterSpacing')));
      expect(document.querySelector('.pptx-menu') !== null).toBe(true);
      act(() => at.focus());
      // Focus that leaves the menu closes it, as DOCX Editor's `listenForPopupFocusLeave`.
      const open = document.querySelector('.pptx-menu') !== null;
      let reached = false;
      const listen = (e: KeyboardEvent) => void (e.key === 'Escape' && (reached = true));
      at.addEventListener('keydown', listen);
      act(() => void fireEvent.keyDown(at, { key: 'Escape' }));
      at.removeEventListener('keydown', listen);
      return { stopped: !reached, open };
    }

    test("into the editor's Find bar: the menu closes, and the Find bar gets Escape", async () => {
      await mount();
      const find = document.createElement('div');
      find.className = 'pptx-find';
      const input = find.appendChild(document.createElement('input'));
      document.querySelector('.pptx-editor')!.append(find);
      expect(escapeAt(input)).toEqual({ stopped: false, open: false });
    });

    test('into a host field outside the editor: the menu closes, and the key is not stopped', async () => {
      await mount();
      const input = document.body.appendChild(document.createElement('input'));
      expect(escapeAt(input)).toEqual({ stopped: false, open: false });
      input.remove();
    });

    test('onto the slide, in a host dialog around the editor: the menu closes', async () => {
      await mount(undefined, true);
      const text = document.querySelector<HTMLTextAreaElement>('.pptx-surface__input')!;
      expect(escapeAt(text).open).toBe(false);
      expect(document.querySelector('.pptx-menu') === null).toBe(true);
    });
  });

  test('a color pick and the reset row give the caret back to the slide text', async () => {
    await mount();
    const text = document.querySelector<HTMLTextAreaElement>('.pptx-surface__input')!;
    const arrow = () =>
      slot('text.color').parentElement!.querySelector<HTMLButtonElement>('[aria-haspopup]')!;
    act(() => arrow().focus());
    act(() => void fireEvent.click(arrow()));
    act(() => void fireEvent.click(screen.getByRole('menuitem', { name: 'Red' })));
    expect(document.activeElement).toBe(text);
    act(() => arrow().focus());
    act(() => void fireEvent.click(arrow()));
    act(() => void fireEvent.click(screen.getAllByRole('menuitem').at(-1)!));
    expect(document.activeElement).toBe(text);
  });
});

describe('focus after a pick, with the real slide surface', () => {
  test('a fill pick on a shape gives the keyboard to the slide, a text pick to the caret', async () => {
    const editor = await mountWithSurface();
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] }));
    const arrow = (id: string) =>
      slot(id).parentElement!.querySelector<HTMLButtonElement>('[aria-haspopup]')!;
    act(() => arrow('shape.fill').focus());
    act(() => void fireEvent.click(arrow('shape.fill')));
    act(() => void fireEvent.click(screen.getByRole('menuitem', { name: 'Red' })));
    const active = document.activeElement as HTMLElement;
    expect(active.classList.contains('pptx-surface__input')).toBe(false);
    expect(active.closest('.pptx-surface') !== null).toBe(true);
    act(() =>
      editor.setSelection({
        kind: 'text',
        nodeId,
        anchor: { paragraph: 0, offset: 0 },
        focus: { paragraph: 0, offset: 4 },
      })
    );
    act(() => arrow('text.color').focus());
    act(() => void fireEvent.click(arrow('text.color')));
    act(() => void fireEvent.click(screen.getByRole('menuitem', { name: 'Red' })));
    expect((document.activeElement as HTMLElement).classList.contains('pptx-surface__input')).toBe(
      true
    );
  });

  test('the size steps are off over a mixed size, with the reason', async () => {
    const editor = await mountWithSurface(
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
              text: para(run('small ', 'sz="1200"') + run('large', 'sz="2400"')),
            }),
          },
        ],
      })
    );
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    const select = (to: number) =>
      act(() =>
        editor.setSelection({
          kind: 'text',
          nodeId,
          anchor: { paragraph: 0, offset: 0 },
          focus: { paragraph: 0, offset: to },
        })
      );
    const steps = () => [
      ...document.querySelectorAll<HTMLButtonElement>(
        '[data-slot="font.size"] .pptx-toolbar__step'
      ),
    ];
    select(11);
    expect(document.querySelector<HTMLInputElement>('.pptx-toolbar__size')!.value).toBe('');
    expect(steps()).toHaveLength(2);
    for (const b of steps()) {
      expect(b.getAttribute('aria-disabled')).toBe('true');
      expect(tooltipText(b)).toContain('more than one size');
    }
    select(3);
    for (const b of steps()) expect(b.hasAttribute('aria-disabled')).toBe(false);
    expect(document.querySelector<HTMLInputElement>('.pptx-toolbar__size')!.value).toBe('12');
  });
});

describe('font size box', () => {
  test('Enter applies the size, stops the key, and gives the caret back; Escape keeps the size', async () => {
    const editor = await mount();
    const box = document.querySelector<HTMLInputElement>('.pptx-toolbar__size')!;
    const text = document.querySelector<HTMLTextAreaElement>('.pptx-surface__input')!;
    act(() => box.focus());
    act(() => void fireEvent.change(box, { target: { value: '40' } }));
    let passed = true;
    act(() => void (passed = fireEvent.keyDown(box, { key: 'Enter' })));
    // The Enter stops at the box, so it never reaches the slide text.
    expect(passed).toBe(false);
    expect(document.activeElement).toBe(text);
    expect(rPr(editor, 'sz')).toBe('4000');
    act(() => box.focus());
    act(() => void fireEvent.change(box, { target: { value: '60' } }));
    act(() => void fireEvent.keyDown(box, { key: 'Escape' }));
    expect(document.activeElement).toBe(text);
    act(() => void fireEvent.blur(box));
    expect(rPr(editor, 'sz')).toBe('4000');
  });
});

describe('Font dialog', () => {
  test('Current color shows the text color, and Offset follows both script boxes', async () => {
    const editor = await mount();
    openFont();
    expect(field<HTMLSelectElement>('color').value).toBe('keep');
    const swatch = field('color')
      .closest('.pptx-color-field')!
      .querySelector<HTMLElement>('.pptx-color-field__button')!;
    const color = editor.snapshot().format?.color;
    expect(typeof color).toBe('string');
    expect(swatch.getAttribute('data-color')).toBe(color!);
    const order = [...dialog()!.querySelectorAll('[name]')].map((el) => el.getAttribute('name'));
    expect(order.indexOf('superscript')).toBeLessThan(order.indexOf('subscript'));
    expect(order.indexOf('subscript')).toBeLessThan(order.indexOf('offset'));
  });

  test('sets the font color and the underline color', async () => {
    const editor = await mount();
    openFont();
    expect(field<HTMLSelectElement>('color').value).toBe('keep');
    expect(field<HTMLSelectElement>('underlineColor').value).toBe('auto');
    // No underline, no underline color.
    expect(field<HTMLSelectElement>('underlineColor').disabled).toBe(true);
    act(() => void fireEvent.change(field('color'), { target: { value: 'accent1' } }));
    act(() => void fireEvent.change(field('underline'), { target: { value: 'single' } }));
    act(() => void fireEvent.change(field('underlineColor'), { target: { value: 'accent2' } }));
    act(() => void fireEvent.click(dialog()!.querySelector('[data-pptx-part="apply"]')!));
    const props = child(firstRun(editor), 'a:rPr')!;
    expect(elements(props).map((c) => c.local)).toEqual(['solidFill', 'uFill']);
    expect(attr(find(child(props, 'a:solidFill')!, 'a:schemeClr'), 'val')).toBe('accent1');
    expect(attr(find(child(props, 'a:uFill')!, 'a:schemeClr'), 'val')).toBe('accent2');
    // The dialog opens again at the run's own underline color.
    openFont();
    expect(field<HTMLSelectElement>('underlineColor').value).toBe('keep');
    act(() => void fireEvent.change(field('underlineColor'), { target: { value: 'auto' } }));
    act(() => void fireEvent.click(dialog()!.querySelector('[data-pptx-part="apply"]')!));
    expect(child(child(firstRun(editor), 'a:rPr')!, 'a:uFill')).toBeUndefined();
  });

  test('seeds from the selection and applies the changed fields as one undo step', async () => {
    const editor = await mount();
    // No toolbar button: menus open the dialog.
    expect(document.querySelector('[data-slot="text.dialog"]') === null).toBe(true);
    openFont();
    expect(dialog() !== null).toBe(true);
    expect(field<HTMLSelectElement>('underline').value).toBe('none');
    act(() => void fireEvent.click(field('superscript')));
    expect(field('offset').value).toBe('30');
    act(() => void fireEvent.change(field('offset'), { target: { value: '45' } }));
    act(() => void fireEvent.click(field('allCaps')));
    act(() => void fireEvent.click(field('doubleStrike')));
    act(() => void fireEvent.click(screen.getByRole('tab', { name: 'Character spacing' })));
    expect(screen.getByRole('tabpanel').getAttribute('aria-labelledby')).toBe(
      screen.getByRole('tab', { name: 'Character spacing' }).id
    );
    act(() => void fireEvent.change(field('spacing'), { target: { value: 'condensed' } }));
    // Condensed starts By at 1 pt.
    expect(field('by').value).toBe('1');
    act(() => void fireEvent.change(field('by'), { target: { value: '0.5' } }));
    act(() => void fireEvent.click(dialog()!.querySelector('[data-pptx-part="apply"]')!));
    expect(dialog() === null).toBe(true);
    expect([
      rPr(editor, 'baseline'),
      rPr(editor, 'cap'),
      rPr(editor, 'strike'),
      rPr(editor, 'spc'),
    ]).toEqual(['45000', 'all', 'dblStrike', '-50']);
    // Unchanged fields are not written.
    expect(rPr(editor, 'u')).toBeUndefined();
    act(() => void editor.exec({ type: 'undo' }));
    expect(rPr(editor, 'baseline')).toBeUndefined();
    expect(editor.snapshot().canUndo).toBe(false);
  });

  test('popups.font replaces the dialog, and false removes the More spacing row', async () => {
    let opened = false;
    function Mine(props: FontDialogProps) {
      if (props.open) opened = true;
      return null;
    }
    await mount({ font: definePopup(Mine) });
    openFont();
    expect(opened).toBe(true);
    expect(dialog() === null).toBe(true);
    cleanup();
    await mount({ font: false });
    act(() => void fireEvent.click(slot('text.characterSpacing')));
    // The row stays, disabled with the reason, as in the menus.
    expect(
      screen.getByText('More spacing…').closest('[role="menuitem"]')!.getAttribute('aria-disabled')
    ).toBe('true');
  });

  test('named parts: a hidden field leaves, a host part restyles, and a host field reads the draft', async () => {
    let read = null as string | null;
    function SizeEcho() {
      read = useFontDialog().values.size;
      return null;
    }
    await mount({
      font: definePopup((props: FontDialogProps) => (
        <FontDialog {...props}>
          <FontDialog.Field name="effects" hidden />
          <FontDialog.Cancel className="host-cancel" />
          <SizeEcho />
        </FontDialog>
      )),
    });
    openFont();
    expect(dialog()!.querySelector('[name="superscript"]') === null).toBe(true);
    expect(dialog()!.querySelector('[name="underline"]') !== null).toBe(true);
    const cancel = dialog()!.querySelector('[data-pptx-part="cancel"]')!;
    expect(cancel.classList.contains('host-cancel')).toBe(true);
    act(() => void fireEvent.change(field('size'), { target: { value: '20' } }));
    expect(read).toBe('20');
    act(() => void fireEvent.click(cancel));
    expect(dialog() === null).toBe(true);
  });

  test('named parts per field: the font row and the spacing tab fields hide or restyle', async () => {
    await mount({
      font: definePopup((props: FontDialogProps) => (
        <FontDialog {...props}>
          <FontDialog.Field name="style" hidden />
          <FontDialog.Field name="kerningAt" hidden />
          <FontDialog.Field name="fontFamily" className="host-family" />
          <FontDialog.Field name="kerning" className="host-kerning" />
        </FontDialog>
      )),
    });
    openFont();
    expect(dialog()!.querySelector('[name="style"]') === null).toBe(true);
    expect(dialog()!.querySelector('[name="kerningAt"]') === null).toBe(true);
    expect(dialog()!.querySelector('[name="size"]') !== null).toBe(true);
    expect(field('fontFamily').closest('label')!.classList.contains('host-family')).toBe(true);
    const kerning = field('kerning').closest('label')!;
    expect([...kerning.classList]).toEqual(['pptx-radio', 'host-kerning']);
  });

  test('More spacing… opens the Character spacing tab', async () => {
    await mount();
    act(() => void fireEvent.click(slot('text.characterSpacing')));
    act(() => void fireEvent.click(screen.getByText('More spacing…')));
    expect(
      screen.getByRole('tab', { name: 'Character spacing' }).getAttribute('aria-selected')
    ).toBe('true');
  });

  test('a size out of range names the field, and a mixed size opens blank', async () => {
    const editor = await mount();
    openFont();
    act(() => void fireEvent.change(field('size'), { target: { value: '5000' } }));
    act(() => void fireEvent.click(dialog()!.querySelector('[data-pptx-part="apply"]')!));
    expect(dialog()!.querySelector('[role="alert"]')!.textContent).toBe(
      'Size must be a number from 1 to 4000 pt.'
    );
    expect(field('size').getAttribute('aria-invalid')).toBe('true');
    expect(editor.snapshot().canUndo).toBe(false);
  });

  test('the arrow keys move between the tabs', async () => {
    await mount();
    openFont();
    const tab = screen.getByRole('tab', { name: 'Font' });
    expect(tab.tabIndex).toBe(0);
    act(() => void fireEvent.keyDown(tab, { key: 'ArrowRight' }));
    expect(
      screen.getByRole('tab', { name: 'Character spacing' }).getAttribute('aria-selected')
    ).toBe('true');
  });

  test('useFontFormat reads and applies the same fields', async () => {
    let hook: UseFontFormatReturn | null = null;
    function Probe() {
      hook = useFontFormat();
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
    const nodeId = editor!.session!.layout(0)!.items[0]!.element.nodeId;
    act(() =>
      editor!.setSelection({
        kind: 'text',
        nodeId,
        anchor: { paragraph: 0, offset: 0 },
        focus: { paragraph: 0, offset: 4 },
      })
    );
    expect(hook!.isEnabled).toBe(true);
    let ok = false;
    act(() => void (ok = hook!.apply({ characterSpacingPt: 3 })));
    expect(ok).toBe(true);
    expect(hook!.format!.characterSpacingPt).toBe(3);
  });
});
