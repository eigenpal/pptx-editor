// The Bullets and Numbering galleries, the menus' dialog rows ("Line spacing options…",
// "More spacing…", "Bullets and numbering…"), and the Bullets and Numbering dialog.

import { afterAll, afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { attr, child, elements, find, xmlPartOf } from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorRoot } from '../PptxEditorRoot.tsx';
import {
  BulletsAndNumberingDialog,
  Toolbar,
  definePopup,
  useBulletFormat,
  useBulletsAndNumberingDialog,
  type BulletsAndNumberingDialogProps,
  type PptxEditorPopups,
  type UseBulletFormatReturn,
} from '../index.ts';

const IN = 914_400;
const deck = (text = para(run('one')) + para(run('two'))): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({ id: 2, x: IN, y: IN, cx: 4 * IN, cy: 2 * IN, text }),
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

async function mount(popups?: PptxEditorPopups, bytes = deck()): Promise<PptxEditorInstance> {
  let editor: PptxEditorInstance | null = null;
  render(
    <PptxEditorRoot onReady={(e) => (editor = e)} {...(popups ? { popups } : {})}>
      <Toolbar>
        {/* Off the default bar: the Font dialog sets it too. */}
        <Toolbar.Slot slotId="text.characterSpacing" />
      </Toolbar>
    </PptxEditorRoot>
  );
  await act(async () => void (await editor!.load(bytes)));
  const nodeId = editor!.session!.layout(0)!.items[0]!.element.nodeId;
  const el = editor!.session!.layout(0)!.items[0]!.element;
  const last = el.kind === 'shape' && el.text ? el.text.paragraphs.length - 1 : 0;
  act(() =>
    editor!.setSelection({
      kind: 'text',
      nodeId,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: last, offset: 1 },
    })
  );
  return editor!;
}

const slot = (id: string) => document.querySelector<HTMLButtonElement>(`[data-slot="${id}"]`)!;
const dialog = (name: string) =>
  document.querySelector<HTMLDialogElement>(`[data-pptx-dialog="${name}"]`);
const field = <T extends HTMLElement>(name: string) =>
  dialog('bulletsAndNumbering')!.querySelector<T>(`[name="${name}"]`)!;
const pick = (label: string) => act(() => void fireEvent.click(screen.getByText(label)));
function pPr(editor: PptxEditorInstance, i = 0) {
  const part = editor.session!.info.slides[0]!.partName;
  const body = find(xmlPartOf(editor.session!.store.package, part)!.root, 'p:txBody')!;
  return child(elements(body).filter((c) => c.local === 'p')[i]!, 'a:pPr');
}

describe('galleries', () => {
  test('a bullet pick writes the character and shows as checked', async () => {
    const editor = await mount();
    act(() => void fireEvent.click(slot('list.bulletStyle')));
    pick('Star bullets');
    expect(attr(find(pPr(editor)!, 'a:buChar'), 'char')).toBe('v');
    expect(attr(find(pPr(editor, 1)!, 'a:buChar'), 'char')).toBe('v');
    act(() => void fireEvent.click(slot('list.bulletStyle')));
    expect(screen.getByText('Star bullets').closest('[role]')!.getAttribute('aria-checked')).toBe(
      'true'
    );
    pick('None');
    expect(find(pPr(editor)!, 'a:buChar')).toBeUndefined();
  });

  test('a numbering pick writes the scheme', async () => {
    const editor = await mount();
    act(() => void fireEvent.click(slot('list.numberStyle')));
    pick('I. II. III.');
    expect(attr(find(pPr(editor)!, 'a:buAutoNum'), 'type')).toBe('romanUcPeriod');
  });
});

describe('dialog rows', () => {
  test('menu rows open the Paragraph and Font dialogs', async () => {
    await mount();
    act(() => void fireEvent.click(slot('list.lineSpacing')));
    pick('Line spacing options…');
    expect(dialog('paragraph') !== null).toBe(true);
    cleanup();
    await mount();
    act(() => void fireEvent.click(slot('text.characterSpacing')));
    pick('More spacing…');
    expect(dialog('font') !== null).toBe(true);
  });

  test('a row whose dialog is turned off is disabled', async () => {
    await mount({ paragraph: false });
    act(() => void fireEvent.click(slot('list.lineSpacing')));
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

describe('Bullets and Numbering dialog', () => {
  test('sets the scheme, start number, size, and color as one undo step', async () => {
    const editor = await mount();
    act(() => void fireEvent.click(slot('list.numberStyle')));
    pick('Bullets and numbering…');
    expect(dialog('bulletsAndNumbering') !== null).toBe(true);
    // The Numbering gallery's row opens the Numbered tab.
    expect(screen.getByRole('tab', { name: 'Numbered' }).getAttribute('aria-selected')).toBe(
      'true'
    );
    act(
      () =>
        void fireEvent.change(field<HTMLSelectElement>('numberStyle'), {
          target: { value: 'romanLcPeriod' },
        })
    );
    act(() => void fireEvent.change(field('startAt'), { target: { value: '3' } }));
    act(() => void fireEvent.change(field('numberSizePercent'), { target: { value: '120' } }));
    act(
      () =>
        void fireEvent.change(field<HTMLSelectElement>('numberColor'), {
          target: { value: 'accent1' },
        })
    );
    act(
      () =>
        void fireEvent.click(
          dialog('bulletsAndNumbering')!.querySelector('[data-pptx-part="apply"]')!
        )
    );
    expect(dialog('bulletsAndNumbering') === null).toBe(true);
    const p = pPr(editor)!;
    expect(elements(p).map((c) => c.local)).toEqual(['buClr', 'buSzPct', 'buFont', 'buAutoNum']);
    expect(attr(find(p, 'a:buAutoNum'), 'startAt')).toBe('3');
    expect(attr(find(p, 'a:buSzPct'), 'val')).toBe('120000');
    act(() => void editor.exec({ type: 'undo' }));
    expect(pPr(editor)).toBeUndefined();
    expect(editor.snapshot().canUndo).toBe(false);
  });

  test('a size out of range is marked on its field, and the tabs move with the arrow keys', async () => {
    await mount();
    act(() => void fireEvent.click(slot('list.bulletStyle')));
    pick('Bullets and numbering…');
    const bulleted = screen.getByRole('tab', { name: 'Bulleted' });
    expect(bulleted.getAttribute('aria-selected')).toBe('true');
    act(() => void fireEvent.keyDown(bulleted, { key: 'ArrowRight' }));
    expect(screen.getByRole('tab', { name: 'Numbered' }).getAttribute('aria-selected')).toBe(
      'true'
    );
    act(
      () =>
        void fireEvent.keyDown(screen.getByRole('tab', { name: 'Numbered' }), { key: 'ArrowLeft' })
    );
    act(
      () =>
        void fireEvent.change(field<HTMLSelectElement>('style'), {
          target: { value: 'filledRound' },
        })
    );
    act(() => void fireEvent.change(field('sizePercent'), { target: { value: '500' } }));
    act(
      () =>
        void fireEvent.click(
          dialog('bulletsAndNumbering')!.querySelector('[data-pptx-part="apply"]')!
        )
    );
    expect(dialog('bulletsAndNumbering')!.querySelector('[role="alert"]')?.textContent).toContain(
      '25 to 400'
    );
    expect(field('sizePercent').getAttribute('aria-invalid')).toBe('true');
  });

  test('a size change alone on mixed paragraphs keeps each its own bullet or number', async () => {
    const list = (child: string) => `<a:pPr marL="342900" indent="-342900">${child}</a:pPr>`;
    const editor = await mount(
      undefined,
      deck(
        para(run('dot'), list('<a:buChar char="•"/>')) +
          para(run('roman'), list('<a:buAutoNum type="romanUcPeriod"/>')) +
          para(run('plain'))
      )
    );
    act(() => void fireEvent.click(slot('list.numberStyle')));
    pick('Bullets and numbering…');
    expect(field<HTMLSelectElement>('numberStyle').value).toBe('');
    act(() => void fireEvent.change(field('numberSizePercent'), { target: { value: '150' } }));
    act(
      () =>
        void fireEvent.click(
          dialog('bulletsAndNumbering')!.querySelector('[data-pptx-part="apply"]')!
        )
    );
    expect(dialog('bulletsAndNumbering') === null).toBe(true);
    expect(elements(pPr(editor, 0)).map((c) => c.local)).toEqual(['buSzPct', 'buChar']);
    expect(elements(pPr(editor, 1)).map((c) => c.local)).toEqual(['buSzPct', 'buAutoNum']);
    expect(pPr(editor, 2)).toBeUndefined();
  });

  test('bullets in different fonts open the Font field at "Mixed", in a fixed order', async () => {
    const list = (child: string) => `<a:pPr marL="342900" indent="-342900">${child}</a:pPr>`;
    await mount(
      undefined,
      deck(
        para(run('dot'), list('<a:buFont typeface="Arial"/><a:buChar char="•"/>')) +
          para(run('star'), list('<a:buFont typeface="Wingdings"/><a:buChar char="v"/>'))
      )
    );
    act(() => void fireEvent.click(slot('list.bulletStyle')));
    pick('Bullets and numbering…');
    const font = field<HTMLSelectElement>('font');
    expect(font.value).toBe('');
    expect(font.selectedOptions[0]!.textContent).toBe('Mixed');
    const order = () =>
      [...field<HTMLSelectElement>('font').options]
        .map((o) => o.textContent)
        .filter((l) => l !== 'Mixed');
    const before = order();
    const symbol = [...font.options].find((o) => o.value === 'Symbol')!.value;
    act(() => void fireEvent.change(font, { target: { value: symbol } }));
    expect(field<HTMLSelectElement>('font').value).toBe('Symbol');
    expect(order()).toEqual(before);
  });

  test('size, color, and start are dimmed while the style is None', async () => {
    await mount();
    act(() => void fireEvent.click(slot('list.numberStyle')));
    pick('Bullets and numbering…');
    expect(field<HTMLSelectElement>('numberStyle').value).toBe('none');
    for (const name of ['numberSizePercent', 'numberColor', 'startAt'])
      expect(field<HTMLInputElement>(name).disabled).toBe(true);
    act(
      () =>
        void fireEvent.change(field<HTMLSelectElement>('numberStyle'), {
          target: { value: 'arabicPeriod' },
        })
    );
    for (const name of ['numberSizePercent', 'numberColor', 'startAt'])
      expect(field<HTMLInputElement>(name).disabled).toBe(false);
  });

  test('paragraphs with no list open on None, and OK with no change applies nothing', async () => {
    const editor = await mount();
    act(() => void fireEvent.click(slot('list.bulletStyle')));
    pick('Bullets and numbering…');
    expect(field<HTMLSelectElement>('style').value).toBe('none');
    act(
      () =>
        void fireEvent.click(
          dialog('bulletsAndNumbering')!.querySelector('[data-pptx-part="apply"]')!
        )
    );
    expect(editor.snapshot().canUndo).toBe(false);
  });
});

describe('the bullet color grid', () => {
  test('the swatch button opens the grid, and a standard color goes through OK', async () => {
    const editor = await mount();
    act(() => void fireEvent.click(slot('list.bulletStyle')));
    pick('Bullets and numbering…');
    const d = dialog('bulletsAndNumbering')!;
    act(
      () => void fireEvent.change(field<HTMLSelectElement>('style'), { target: { value: 'star' } })
    );
    const grid = () => d.querySelector('.pptx-color-field__panel');
    const button = d.querySelector<HTMLButtonElement>('.pptx-color-field__button')!;
    expect(button.getAttribute('aria-label')).toBe('Color: color grid');
    act(() => void fireEvent.click(button));
    expect(grid() !== null).toBe(true);
    // Escape closes the grid, not the dialog.
    act(() => void fireEvent.keyDown(grid()!, { key: 'Escape' }));
    expect(grid() === null).toBe(true);
    expect(dialog('bulletsAndNumbering') !== null).toBe(true);
    act(() => void fireEvent.click(button));
    act(() => void fireEvent.click(grid()!.querySelector('[aria-label="Dark Red"]')!));
    expect(grid() === null).toBe(true);
    expect(field<HTMLSelectElement>('color').value).toBe('#C00000');
    act(() => void fireEvent.click(d.querySelector('[data-pptx-part="apply"]')!));
    const p = pPr(editor)!;
    expect(attr(find(find(p, 'a:buClr')!, 'a:srgbClr'), 'val')).toBe('C00000');
  });
});

describe('Bullets and Numbering dialog parts', () => {
  test('a hidden field leaves, and a host field reads the draft', async () => {
    let kind = null as string | null;
    function KindEcho() {
      kind = useBulletsAndNumberingDialog().values.kind;
      return null;
    }
    await mount({
      bulletsAndNumbering: definePopup((props: BulletsAndNumberingDialogProps) => (
        <BulletsAndNumberingDialog {...props}>
          <BulletsAndNumberingDialog.Field name="bullet" hidden />
          <KindEcho />
        </BulletsAndNumberingDialog>
      )),
    });
    act(() => void fireEvent.click(slot('list.numberStyle')));
    pick('Bullets and numbering…');
    expect(dialog('bulletsAndNumbering')!.querySelector('[name="style"]') === null).toBe(true);
    expect(field('startAt') !== null).toBe(true);
    expect(kind).toBe('number');
  });

  test('one field hides and one restyles, on its own tab', async () => {
    await mount({
      bulletsAndNumbering: definePopup((props: BulletsAndNumberingDialogProps) => (
        <BulletsAndNumberingDialog {...props}>
          <BulletsAndNumberingDialog.Field name="numberColor" hidden />
          <BulletsAndNumberingDialog.Field name="startAt" className="host-start" />
        </BulletsAndNumberingDialog>
      )),
    });
    act(() => void fireEvent.click(slot('list.numberStyle')));
    pick('Bullets and numbering…');
    const d = dialog('bulletsAndNumbering')!;
    expect(d.querySelector('[name="numberColor"]') === null).toBe(true);
    expect(d.querySelector('[name="color"]') !== null).toBe(true);
    expect(field('startAt').closest('label')!.classList.contains('host-start')).toBe(true);
  });
});

describe('useBulletFormat', () => {
  test('reads and applies the same bullet format as the dialog', async () => {
    let hook: UseBulletFormatReturn | null = null;
    function Probe() {
      hook = useBulletFormat();
      return null;
    }
    let editor: PptxEditorInstance | null = null;
    render(
      <PptxEditorRoot onReady={(e) => (editor = e)}>
        <Probe />
      </PptxEditorRoot>
    );
    await act(async () => void (await editor!.load(deck())));
    expect(hook!.isEnabled).toBe(false);
    const nodeId = editor!.session!.layout(0)!.items[0]!.element.nodeId;
    act(() =>
      editor!.setSelection({
        kind: 'text',
        nodeId,
        anchor: { paragraph: 0, offset: 0 },
        focus: { paragraph: 1, offset: 1 },
      })
    );
    expect(hook!.isEnabled).toBe(true);
    expect(hook!.format).toBeNull();
    let ok = false;
    act(() => void (ok = hook!.apply({ kind: 'bullet', char: '–', sizePercent: 150 })));
    expect(ok).toBe(true);
    expect(hook!.format).toMatchObject({ kind: 'bullet', char: '–', sizePercent: 150 });
    act(() => void (ok = hook!.apply({ kind: 'bullet', sizePercent: 999 })));
    expect(ok).toBe(false);
  });
});

describe('custom bullet', () => {
  test('a custom character and font, as Customize… sets them', async () => {
    const editor = await mount();
    act(() => void fireEvent.click(slot('list.bulletStyle')));
    pick('Bullets and numbering…');
    const style = field<HTMLSelectElement>('style');
    act(() => void fireEvent.change(style, { target: { value: 'star' } }));
    // A gallery pick shows its character and font.
    expect([field<HTMLInputElement>('char').value, field<HTMLSelectElement>('font').value]).toEqual(
      ['v', 'Wingdings']
    );
    // The character shows in its bullet font.
    expect(field<HTMLInputElement>('char').style.fontFamily).toContain('Wingdings');
    // And as its Unicode twin, which reads the same where Wingdings is not installed.
    const preview = () =>
      field<HTMLInputElement>('char').closest('label')!.querySelector('[data-preview]');
    expect(preview()!.textContent).toBe('❖');
    act(() => void fireEvent.change(field('char'), { target: { value: '→' } }));
    const textFont = [...field<HTMLSelectElement>('font').options].find(
      (o) => o.textContent === 'Text font'
    )!.value;
    act(() => void fireEvent.change(field('font'), { target: { value: textFont } }));
    // The symbol fonts stay on offer after the font changes, so the user can go back.
    const offered = [...field<HTMLSelectElement>('font').options].map((o) => o.value);
    expect(offered).toEqual(expect.arrayContaining(['Wingdings', 'Symbol', 'Courier New']));
    expect(style.value).toBe('current');
    expect(screen.getByText('Custom: →')).toBeTruthy();
    act(
      () =>
        void fireEvent.click(
          dialog('bulletsAndNumbering')!.querySelector('[data-pptx-part="apply"]')!
        )
    );
    const p = pPr(editor)!;
    // No font: the bullet follows the text font.
    expect(elements(p).map((c) => c.local)).toEqual(['buFontTx', 'buChar']);
    expect(attr(find(p, 'a:buChar'), 'char')).toBe('→');
    // The dialog opens again at the custom bullet; two characters are refused on the field.
    act(() => void fireEvent.click(slot('list.bulletStyle')));
    pick('Bullets and numbering…');
    expect(field<HTMLSelectElement>('style').value).toBe('current');
    expect(field<HTMLInputElement>('char').value).toBe('→');
    act(() => void fireEvent.change(field('char'), { target: { value: '->' } }));
    act(
      () =>
        void fireEvent.click(
          dialog('bulletsAndNumbering')!.querySelector('[data-pptx-part="apply"]')!
        )
    );
    expect(dialog('bulletsAndNumbering')!.querySelector('[role="alert"]')?.textContent).toContain(
      'one character'
    );
    expect(field('char').getAttribute('aria-invalid')).toBe('true');
    expect(attr(find(pPr(editor)!, 'a:buChar'), 'char')).toBe('→');
  });
});
