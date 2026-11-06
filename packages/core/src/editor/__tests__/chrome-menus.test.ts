import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { buildSlideStructureDeck } from '../../../../../scripts/fixtures/slide-structure.ts';
import {
  CANVAS_CONTEXT_MENUS,
  canvasMenuKind,
  CHROME_MENUS,
  chromeCommandState,
  chromeControlForSlot,
  chromeRowState,
  createPptxEditor,
  POINTER_OPEN_ATTRIBUTE,
  focusMenuRowAtPointer,
  focusPopoverOnOpen,
  listenPopoverDismiss,
  panelKeyPressed,
  restorePopoverFocus,
  resolveChromeMenuRow,
  runChromeRow,
  toolbarCommandState,
  type ChromeActionHost,
  type ChromeMenuEntry,
  type PptxEditorInstance,
} from '../index.ts';
import { ICON_PATHS } from '../icon-paths.ts';
import { GlobalWindow } from 'happy-dom';

// A DOM of this file's own: the core run has no global DOM (`bun run test`).
const win = new GlobalWindow();
const document = win.document as unknown as Document;
const window = win as unknown as Window;
const Event = win.Event as unknown as typeof globalThis.Event;

// The menu registry (the menu bar and the canvas context menus) and how a row runs.

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

function items(entries: readonly ChromeMenuEntry[]): ChromeMenuEntry[] {
  return entries.flatMap((e) => (e.kind === 'submenu' ? items(e.items) : [e]));
}

const host = (): ChromeActionHost & { calls: string[] } => {
  const calls: string[] = [];
  return {
    calls,
    openDialog: (key) => calls.push(`dialog:${key}`),
    startSlideshow: (from) => calls.push(`show:${from}`),
    openFind: (replace) => calls.push(replace ? 'replace' : 'find'),
    openLink: () => calls.push('link'),
    pickImage: (target) => calls.push(target ? `image:${target}` : 'image'),
  };
};

describe('menu registry', () => {
  test('the menu bar follows the menu design and every row names a registered slot', () => {
    expect(CHROME_MENUS.map((m) => m.id)).toEqual([
      'file',
      'edit',
      'view',
      'insert',
      'format',
      'arrange',
      'slideShow',
    ]);
    const all = [
      ...CHROME_MENUS.flatMap((m) => m.entries),
      ...Object.values(CANVAS_CONTEXT_MENUS).flat(),
    ];
    for (const entry of items(all)) {
      if (entry.kind !== 'item') continue;
      expect(chromeControlForSlot(entry.slot)).not.toBeNull();
      expect(resolveChromeMenuRow(entry)).not.toBeNull();
    }
  });

  test('Insert starts with New Slide and Slide Show ends with Hide Slide', () => {
    const menu = (id: string) => CHROME_MENUS.find((m) => m.id === id)!.entries;
    expect(menu('insert')[0]).toMatchObject({
      kind: 'item',
      slot: 'slide.add',
      direct: true,
      shortcutKey: 'contextMenu.shortcuts.newSlide',
    });
    expect(menu('slideShow').at(-1)).toMatchObject({ kind: 'item', slot: 'slide.hide' });
    expect(resolveChromeMenuRow({ kind: 'item', slot: 'slide.hide' })?.toggle).toBe(true);
  });

  test('a value row takes its label from the option, and a toggle row shows a check', () => {
    const front = resolveChromeMenuRow({
      kind: 'item',
      slot: 'arrange.order',
      value: 'front',
      labelKey: 'toolbar.bringToFront',
    });
    expect(front?.key).toBe('arrange.order.front');
    expect(front?.toggle).toBe(false);
    expect(resolveChromeMenuRow({ kind: 'item', slot: 'text.bold' })?.toggle).toBe(true);
  });
});

describe('Format menu labels', () => {
  test('the Format menu names the picture row for what it does: Format picture', () => {
    const format = CHROME_MENUS.find((m) => m.id === 'format')!;
    const row = items(format.entries).find(
      (e) => e.kind === 'item' && e.slot === 'image.properties'
    );
    expect(row?.kind === 'item' ? resolveChromeMenuRow(row)?.labelKey : null).toBe(
      'pictureFormat.properties'
    );
  });
});

describe('rows the adapter runs', () => {
  test('dialog, show, find, link, and picture rows go to the adapter', async () => {
    const editor = await open();
    const h = host();
    for (const slot of [
      'file.pageSetup',
      'slide.background',
      'view.playFromCurrent',
      'find.replace',
      'text.link',
      'insert.image',
      'image.change',
      'image.altText',
      'image.properties',
    ] as const)
      runChromeRow(editor, h, slot);
    // Change Picture asks for a file that replaces the selected picture's image.
    expect(h.calls).toEqual([
      'dialog:pageSetup',
      'dialog:background',
      'show:current',
      'replace',
      'link',
      'image',
      'image:replaceImage',
      'dialog:imageAltText',
      'dialog:imageProperties',
    ]);
  });

  test("Format menu's dialog rows come first and open their dialogs", async () => {
    const format = CHROME_MENUS.find((m) => m.id === 'format')!;
    const dialogs = items(format.entries)
      .filter((e) => e.kind === 'item')
      .map((e) => (e.kind === 'item' ? `${e.slot}${e.value ? `=${e.value}` : ''}` : ''))
      .filter((k) => /dialog/.test(k));
    expect(dialogs).toEqual([
      'text.dialog',
      'paragraph.dialog',
      'list.bulletStyle=dialog',
      'paragraph.columns=dialog',
    ]);
    // They open the menu, before the text toggles.
    const slots = items(format.entries).map((e) => (e.kind === 'item' ? e.slot : ''));
    expect(slots.indexOf('text.dialog')).toBe(0);
    expect(slots.indexOf('paragraph.columns')).toBeLessThan(slots.indexOf('text.bold'));
    const editor = await open();
    const h = host();
    runChromeRow(editor, h, 'paragraph.columns', 'dialog');
    expect(h.calls).toEqual(['dialog:columns']);
  });

  test('a table selected as an object has the shape rows and the table submenus', () => {
    const shapes = { kind: 'shapes', nodeIds: ['7'] } as const;
    expect(canvasMenuKind({ selection: shapes, image: null })).toBe('shapes');
    const table = {} as NonNullable<Parameters<typeof canvasMenuKind>[0]['table']>;
    expect(canvasMenuKind({ selection: shapes, image: null, table })).toBe('tableFrame');
    const ids = CANVAS_CONTEXT_MENUS.tableFrame.map((e) =>
      e.kind === 'item' ? e.slot : e.kind === 'submenu' ? e.id : ''
    );
    expect(ids.slice(0, 7)).toEqual([
      'edit.cut',
      'edit.copy',
      'edit.paste',
      'more.duplicate',
      'more.delete',
      '',
      'tableInsert',
    ]);
    // On the whole table, Delete removes it: no Delete submenu beside the Delete row.
    expect(ids).not.toContain('tableDelete');
    expect(ids).not.toContain('tableSelect');
  });

  test('clipboard rows: copy and paste shapes through the editor clipboard', async () => {
    const editor = await open();
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    expect(toolbarCommandState(editor, 'edit.copy').enabled).toBe(false);
    editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] });
    expect(toolbarCommandState(editor, 'edit.paste').disabledReason).toBe(
      'contextMenu.nothingCopied'
    );
    runChromeRow(editor, host(), 'edit.copy');
    expect(toolbarCommandState(editor, 'edit.paste').enabled).toBe(true);
    runChromeRow(editor, host(), 'edit.paste');
    expect(editor.session!.layout(0)!.items.length).toBe(2);
  });

  test('a value row asks the engine about its own value', async () => {
    const editor = await open();
    expect(chromeRowState(editor, 'insert.line', 'arrow').isEnabled).toBe(true);
    // A value the slot does not take has no command: the row is disabled, not dropped.
    expect(chromeRowState(editor, 'insert.line', 'zigzag')).toMatchObject({
      isEnabled: false,
      disabledReason: 'disabled.notWired',
    });
  });

  test('viewing mode refuses Cut and Paste and the background dialog', async () => {
    const editor = await open();
    editor.setMode('view');
    expect(toolbarCommandState(editor, 'edit.paste').disabledReason).toBe('disabled.viewing');
    expect(toolbarCommandState(editor, 'slide.background').enabled).toBe(false);
    expect(toolbarCommandState(editor, 'file.pageSetup').enabled).toBe(true);
    expect(toolbarCommandState(editor, 'view.playFromStart').enabled).toBe(true);
  });
});

describe('clipboard rows and frames the clipboard cannot carry', () => {
  test('Cut and Copy are off, with the reason, while a chart frame is selected', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(buildSlideStructureDeck())).ok).toBe(true);
    const items = editor.session!.layout(0)!.items;
    const chart = items.find((i) => i.element.kind === 'graphicFrame')!;
    editor.setSelection({ kind: 'shapes', nodeIds: [chart.element.nodeId] });
    for (const slot of ['edit.cut', 'edit.copy'] as const) {
      const state = chromeCommandState(editor, slot);
      expect(state.isEnabled).toBe(false);
      expect(state.disabledReason).toBe('contextMenu.cannotCopyFrame');
    }
  });

  test('a chart inside a selected group turns Cut and Copy off too', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(buildSlideStructureDeck())).ok).toBe(true);
    const chart = editor.session!.layout(0)!.items.find((i) => i.element.kind === 'graphicFrame')!;
    expect(editor.exec({ type: 'insertShape', preset: 'rect' }).ok).toBe(true);
    const sel = editor.snapshot().selection;
    const rect = sel.kind === 'shapes' ? sel.nodeIds[0]! : '';
    editor.setSelection({ kind: 'shapes', nodeIds: [chart.element.nodeId, rect] });
    expect(editor.exec({ type: 'group' }).ok).toBe(true);
    expect(editor.snapshot().selection.kind).toBe('shapes');
    for (const slot of ['edit.cut', 'edit.copy'] as const)
      expect(chromeCommandState(editor, slot).disabledReason).toBe('contextMenu.cannotCopyFrame');
  });

  test('a chart picked inside a group turns Cut and Copy off', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(buildSlideStructureDeck())).ok).toBe(true);
    const chart = editor.session!.layout(0)!.items.find((i) => i.element.kind === 'graphicFrame')!;
    const chartId = chart.element.nodeId;
    expect(editor.exec({ type: 'insertShape', preset: 'rect' }).ok).toBe(true);
    const sel = editor.snapshot().selection;
    const rect = sel.kind === 'shapes' ? sel.nodeIds[0]! : '';
    editor.setSelection({ kind: 'shapes', nodeIds: [chartId, rect] });
    expect(editor.exec({ type: 'group' }).ok).toBe(true);
    const grouped = editor.snapshot().selection;
    const group = grouped.kind === 'shapes' ? grouped.nodeIds[0]! : '';
    editor.setSelection({ kind: 'shapes', nodeIds: [chartId], group });
    expect(editor.snapshot().selection).toMatchObject({ nodeIds: [chartId], group });
    for (const slot of ['edit.cut', 'edit.copy'] as const)
      expect(chromeCommandState(editor, slot).disabledReason).toBe('contextMenu.cannotCopyFrame');
  });
});

describe('menu and popover dismissal', () => {
  test('a popover closes on a pointer press outside, a wheel outside, a blur, and a resize', () => {
    const panel = document.createElement('div');
    const anchor = document.createElement('button');
    const outside = document.createElement('div');
    document.body.append(panel, anchor, outside);
    let closed = 0;
    const stop = listenPopoverDismiss({
      panel: () => panel,
      anchor,
      focusInside: () => false,
      onClose: () => (closed += 1),
    });
    // Presses and wheels on the panel or its trigger keep it.
    for (const el of [panel, anchor]) el.dispatchEvent(new Event('pointerdown', { bubbles: true }));
    panel.dispatchEvent(new Event('wheel', { bubbles: true }));
    expect(closed).toBe(0);
    // A press the slide surface prevents fires `pointerdown` and no `mousedown`.
    outside.dispatchEvent(new Event('pointerdown', { bubbles: true }));
    expect(closed).toBe(1);
    outside.dispatchEvent(new Event('wheel', { bubbles: true }));
    expect(closed).toBe(2);
    window.dispatchEvent(new Event('blur'));
    window.dispatchEvent(new Event('resize'));
    expect(closed).toBe(4);
    stop();
    outside.dispatchEvent(new Event('pointerdown', { bubbles: true }));
    window.dispatchEvent(new Event('blur'));
    expect(closed).toBe(4);
    panel.remove();
    anchor.remove();
    outside.remove();
  });
});

describe('menu row icons and the pointer', () => {
  test('Format > Bullets and numbering… has an icon of its own', () => {
    const entry = CHROME_MENUS.find((m) => m.id === 'format')!.entries.find(
      (e) => e.kind === 'item' && e.slot === 'list.bulletStyle' && e.value === 'dialog'
    );
    const row = resolveChromeMenuRow(entry as Extract<ChromeMenuEntry, { kind: 'item' }>)!;
    expect(row.paths).toEqual(ICON_PATHS['list_alt']!);
  });

  test('a row under a moving pointer takes focus; a nested panel row and a non-row do not', () => {
    const panel = document.createElement('div');
    panel.setAttribute('role', 'menu');
    const rowOf = (parent: HTMLElement, role = 'menuitem'): HTMLElement => {
      const b = document.createElement('button');
      b.setAttribute('role', role);
      b.tabIndex = -1;
      parent.append(b);
      return b;
    };
    const first = rowOf(panel);
    const second = rowOf(panel, 'menuitemcheckbox');
    const label = document.createElement('span');
    second.append(label);
    const nested = document.createElement('div');
    nested.setAttribute('role', 'menu');
    panel.append(nested);
    const inner = rowOf(nested);
    document.body.append(panel);
    first.focus();
    let x = 0;
    const move = (target: EventTarget, still = false) => {
      if (!still) x += 5;
      return { target, clientX: x, clientY: 10 };
    };
    // The first event only records where the pointer is: a menu that opens under a resting
    // pointer keeps its focus.
    focusMenuRowAtPointer(panel, move(label));
    expect(document.activeElement === first).toBe(true);
    // An event with no move (the browser's event after a layout change) does nothing either.
    focusMenuRowAtPointer(panel, move(label, true));
    expect(document.activeElement === first).toBe(true);
    // A real move over a row's label focuses the row.
    focusMenuRowAtPointer(panel, move(label));
    expect(document.activeElement === second).toBe(true);
    // A row of a nested panel belongs to that panel; the panel itself is not a row.
    focusMenuRowAtPointer(panel, move(inner));
    focusMenuRowAtPointer(panel, move(panel));
    focusMenuRowAtPointer(null, move(first));
    expect(document.activeElement === second).toBe(true);
    focusMenuRowAtPointer(panel, move(first));
    expect(document.activeElement === first).toBe(true);
    panel.remove();
  });

  test('a menu opens on its checked radio row, from the keyboard and the pointer', () => {
    const slide = document.createElement('textarea');
    const trigger = document.createElement('button');
    const panel = document.createElement('div');
    panel.setAttribute('role', 'menu');
    const rows = [false, true, false].map((on) => {
      const row = document.createElement('button');
      row.setAttribute('role', 'menuitemradio');
      row.setAttribute('aria-checked', String(on));
      row.tabIndex = -1;
      panel.append(row);
      return row;
    });
    document.body.append(slide, trigger, panel);
    trigger.focus();
    focusPopoverOnOpen(panel, trigger, { role: 'menu', focusFirst: false });
    expect(document.activeElement === rows[1]).toBe(true);
    slide.focus();
    const back = focusPopoverOnOpen(panel, trigger, { role: 'menu', focusFirst: false });
    expect(document.activeElement === rows[1]).toBe(true);
    expect(panel.hasAttribute(POINTER_OPEN_ATTRIBUTE)).toBe(true);
    expect(back === slide).toBe(true);
    slide.remove();
    trigger.remove();
    panel.remove();
  });

  test('a pointer open focuses the panel; closing returns focus, and Escape goes to the trigger', () => {
    const slide = document.createElement('textarea');
    const trigger = document.createElement('button');
    const panel = document.createElement('div');
    const row = document.createElement('button');
    row.setAttribute('role', 'menuitem');
    row.tabIndex = -1;
    panel.append(row);
    document.body.append(slide, trigger, panel);
    // The slide text has focus: the trigger was clicked without taking focus.
    slide.focus();
    const back = focusPopoverOnOpen(panel, trigger, { role: 'menu', focusFirst: false });
    expect(document.activeElement === panel).toBe(true);
    expect(panel.hasAttribute(POINTER_OPEN_ATTRIBUTE)).toBe(true);
    expect(back === slide).toBe(true);
    // A key shows the focus ring from then on.
    panelKeyPressed(panel);
    expect(panel.hasAttribute(POINTER_OPEN_ATTRIBUTE)).toBe(false);
    // Closed after a pick, with focus dropped on <body>: focus goes back to the slide text.
    row.focus();
    panel.remove();
    restorePopoverFocus(trigger, true, back);
    expect(document.activeElement === slide).toBe(true);
    // A keyboard open (the trigger has focus) focuses the first row and returns nothing.
    document.body.append(panel);
    trigger.focus();
    expect(focusPopoverOnOpen(panel, trigger, { role: 'menu', focusFirst: false }) === null).toBe(
      true
    );
    expect(document.activeElement === row).toBe(true);
    // Escape inside the popover returns focus to the trigger.
    let closed = 0;
    const stop = listenPopoverDismiss({
      panel: () => panel,
      anchor: trigger,
      focusInside: () => true,
      onClose: () => (closed += 1),
    });
    row.dispatchEvent(new win.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }) as never);
    stop();
    expect(closed).toBe(1);
    expect(document.activeElement === trigger).toBe(true);
    slide.remove();
    trigger.remove();
    panel.remove();
  });
});
