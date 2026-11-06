// The slide canvas: Tab through the shapes, redo that keeps the caret, a press on the first
// glyphs of a text box with no inset, the caret kept in view, the IME input on the caret,
// shapes outside the slide, the wheel zoom step and the one zoom range, the table context menu,
// and the triple-click paragraph.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import {
  buildPptx,
  cell,
  para,
  run,
  shape,
  table,
} from '../../../../../scripts/fixtures/builder.ts';
import { attr, elements, find, textOf, xmlPartOf, type Element } from '../../store/index.ts';
import type { EditorSelection } from '../../contracts/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { CANVAS_CONTEXT_MENUS, canvasMenuKind } from '../chrome-menus.ts';
import { canvasMenuLabelKey } from '../menu-arrangement.ts';
import { contextMenuSelection } from '../surface-context-menu.ts';
import { caretAt } from '../surface-geometry.ts';
import { onBorder, onGlyph } from '../surface-hit.ts';
import { scrollBoxIntoView } from '../surface-view.ts';
import { shapeKeyAction, tabSelection } from '../shape-keys.ts';
import { editNoticeTone, elementKindLabelKey } from '../canvas-support.ts';
import {
  ZOOM_MAX,
  ZOOM_MIN,
  clampZoom,
  wheelZoomFactor,
  zoomKeyAction,
  zoomLimitReason,
  zoomStep,
} from '../zoom.ts';

const IN = 914_400;
const NO_INSET =
  '<a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"><a:noAutofit/></a:bodyPr>';

/** Two loose shapes (one text box with no inset), a group of two, and a 2 × 2 table. */
function deck(): Uint8Array {
  const group =
    '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="914400" y="4572000"/><a:ext cx="1828800" cy="914400"/>' +
    '<a:chOff x="914400" y="4572000"/><a:chExt cx="1828800" cy="914400"/></a:xfrm></p:grpSpPr>' +
    shape({
      id: 11,
      name: 'G1',
      x: IN,
      y: 5 * IN,
      cx: IN,
      cy: IN,
      bodyPr: NO_INSET,
      text: para(run('Alpha beta')),
    }) +
    shape({ id: 12, name: 'G2', x: 2 * IN, y: 5 * IN, cx: IN, cy: IN }) +
    '</p:grpSp>';
  return buildPptx({
    slides: [
      {
        shapes:
          shape({
            id: 2,
            name: 'T',
            x: IN,
            y: IN,
            cx: 6 * IN,
            cy: IN,
            bodyPr: NO_INSET,
            text: para(run('Hello world')) + para(run('Second line')),
          }) +
          shape({ id: 3, name: 'B', x: 8 * IN, y: IN, cx: IN, cy: IN }) +
          group +
          table({
            id: 4,
            x: 6 * IN,
            y: 4 * IN,
            cx: 2 * IN,
            cy: 2 * 370_840,
            columns: [IN, IN],
            tblPr: '<a:tblPr/>',
            rows: [0, 1]
              .map((r) => `<a:tr h="370840">${cell(`r${r}a`)}${cell(`r${r}b`)}</a:tr>`)
              .join(''),
          }),
      },
    ],
  });
}

async function open(): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(deck())).ok).toBe(true);
  return e;
}

const rootOf = (e: PptxEditorInstance): Element =>
  xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!.root;
const tree = (e: PptxEditorInstance): Element => find(rootOf(e), 'p:spTree')!;
const idOf = (e: PptxEditorInstance, name: string): string => {
  const walk = (el: Element): string | undefined => {
    for (const c of elements(el)) {
      if (attr(find(c, 'p:cNvPr'), 'name') === name) return c.id;
      if (c.local === 'grpSp') {
        const inner = walk(c);
        if (inner) return inner;
      }
    }
    return undefined;
  };
  return walk(tree(e))!;
};
const tableId = (e: PptxEditorInstance): string =>
  elements(tree(e)).find((c) => c.local === 'graphicFrame')!.id;
const textItem = (e: PptxEditorInstance) => {
  const item = e.session!.layout(0)!.items.find((i) => i.element.nodeId === idOf(e, 'T'))!;
  if (item.kind !== 'shape') throw new Error('no shape');
  return item;
};
const caret = (paragraph: number, offset: number) => ({ paragraph, offset });

describe('Tab through the shapes', () => {
  test('Tab steps in z-order from no selection; past the last shape it leaves the slide', async () => {
    const e = await open();
    const layout = e.session!.layout(0);
    const order = ['T', 'B', 'G', '4'].map((n) => (n === '4' ? tableId(e) : idOf(e, n)));
    let sel: EditorSelection = { kind: 'none' };
    const seen: string[] = [];
    for (;;) {
      const next = tabSelection(layout, sel as Extract<EditorSelection, { kind: 'none' }>, false);
      if (!next) break;
      expect(next.kind).toBe('shapes');
      if (next.kind !== 'shapes') break;
      seen.push(next.nodeIds[0]);
      sel = next;
    }
    expect(seen).toEqual(order);
    // Shift+Tab from no selection starts at the frontmost shape.
    expect(tabSelection(layout, { kind: 'none' }, true)).toEqual({
      kind: 'shapes',
      nodeIds: [order[3]],
    });
    // Shift+Tab on the first shape leaves.
    expect(tabSelection(layout, { kind: 'shapes', nodeIds: [order[0]] }, true)).toBeNull();
    e.destroy();
  });

  test('inside an entered group, Tab cycles through its children', async () => {
    const e = await open();
    const layout = e.session!.layout(0);
    const [g, g1, g2] = ['G', 'G1', 'G2'].map((n) => idOf(e, n));
    const inGroup = { kind: 'shapes' as const, nodeIds: [g1], group: g };
    expect(tabSelection(layout, inGroup, false)).toEqual({ ...inGroup, nodeIds: [g2] });
    expect(tabSelection(layout, { ...inGroup, nodeIds: [g2] }, false)).toEqual(inGroup);
    expect(tabSelection(layout, inGroup, true)).toEqual({ ...inGroup, nodeIds: [g2] });
    e.destroy();
  });

  test('Ctrl+Tab and Alt+Tab are not ours, and a slide with no shapes lets Tab go', () => {
    const none = { kind: 'none' } as const;
    expect(shapeKeyAction('Tab', { shift: false, mod: true }, none, undefined, 'edit')).toBe(null);
    expect(
      shapeKeyAction('Tab', { shift: false, mod: false, alt: true }, none, undefined, 'edit')
    ).toBe(null);
    expect(shapeKeyAction('Tab', { shift: false, mod: false }, none, undefined, 'edit')).toEqual({
      kind: 'leave',
    });
    expect(shapeKeyAction('Delete', { shift: false, mod: false }, none, undefined, 'edit')).toBe(
      null
    );
  });
});

describe('redo keeps the caret', () => {
  test('type, Enter, type, Backspace; undo four times; redo restores a caret that types', async () => {
    const e = await open();
    const id = idOf(e, 'T');
    e.setSelection({ kind: 'text', nodeId: id, anchor: caret(0, 0), focus: caret(0, 0) });
    for (const ch of 'Big ') expect(e.input('text', ch).ok).toBe(true);
    expect(e.input('paragraph').ok).toBe(true);
    for (const ch of 'xy') expect(e.input('text', ch).ok).toBe(true);
    expect(e.input('deleteBackward').ok).toBe(true);
    for (let i = 0; i < 4; i += 1) expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(textOf(find(find(tree(e), 'p:sp')!, 'p:txBody'))).toBe('Hello worldSecond line');
    expect(e.exec({ type: 'redo' }).ok).toBe(true);
    expect(e.snapshot().selection).toEqual({
      kind: 'text',
      nodeId: id,
      anchor: caret(0, 4),
      focus: caret(0, 4),
    });
    expect(e.input('text', 'Q').ok).toBe(true);
    e.destroy();
  });

  test('text typed over a selection is one undo step with the replace', async () => {
    const e = await open();
    const id = idOf(e, 'T');
    const body = () => textOf(find(find(tree(e), 'p:sp')!, 'p:txBody'));
    e.setSelection({ kind: 'text', nodeId: id, anchor: caret(0, 0), focus: caret(0, 5) });
    for (const ch of ' Zx') expect(e.input('text', ch).ok).toBe(true);
    expect(body()).toBe(' Zx worldSecond line');
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(body()).toBe('Hello worldSecond line');
    expect(e.exec({ type: 'redo' }).ok).toBe(true);
    expect(body()).toBe(' Zx worldSecond line');
    // A second replace after the burst is a step of its own.
    e.setSelection({ kind: 'text', nodeId: id, anchor: caret(0, 1), focus: caret(0, 3) });
    expect(e.input('text', 'Q').ok).toBe(true);
    expect(body()).toBe(' Q worldSecond line');
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(body()).toBe(' Zx worldSecond line');
    e.destroy();
  });
});

describe('zoom', () => {
  test('one wheel notch is 10 %; a trackpad pinch moves smoothly; a big delta is one notch', () => {
    expect(wheelZoomFactor(-100, 0)).toBeCloseTo(1.1, 6);
    expect(wheelZoomFactor(100, 0)).toBeCloseTo(1 / 1.1, 6);
    expect(wheelZoomFactor(-3, 1)).toBeCloseTo(1.1, 2);
    expect(wheelZoomFactor(-1000, 0)).toBeCloseTo(1.1, 6);
    const pinch = wheelZoomFactor(-4, 0);
    expect(pinch).toBeGreaterThan(1);
    expect(pinch).toBeLessThan(1.01);
  });

  test('the buttons, keys, and setZoom share one range, and the buttons stop at its ends', async () => {
    expect(clampZoom(9)).toBe(ZOOM_MAX);
    expect(clampZoom(0.01)).toBe(ZOOM_MIN);
    expect(zoomStep(3, 1)).toBe(ZOOM_MAX);
    expect(zoomStep(ZOOM_MAX, 1)).toBeNull();
    expect(zoomStep(0.25, -1)).toBe(ZOOM_MIN);
    expect(zoomLimitReason(ZOOM_MAX, 1)).toBe('disabled.zoomMax');
    expect(zoomLimitReason(ZOOM_MIN, -1)).toBe('disabled.zoomMin');
    expect(zoomLimitReason(1, 1)).toBeNull();
    expect(zoomKeyAction('=', true)).toBe(1.25);
    expect(zoomKeyAction('0', true)).toBe('fit');
    expect(zoomKeyAction('=', false)).toBeNull();
    const e = await open();
    e.setZoom(8);
    expect(e.snapshot().zoom).toBe(ZOOM_MAX);
    e.destroy();
  });
});

describe('the table context menu', () => {
  test('a caret in a cell or a cell range opens the table menu with row and column rows', () => {
    const cellCaret: EditorSelection = {
      kind: 'text',
      nodeId: 'f',
      cell: { row: 0, column: 0 },
      anchor: caret(0, 0),
      focus: caret(0, 0),
    };
    expect(canvasMenuKind({ selection: cellCaret })).toBe('table');
    expect(
      canvasMenuKind({
        selection: {
          kind: 'cells',
          nodeId: 'f',
          anchor: { row: 0, column: 0 },
          focus: { row: 1, column: 1 },
        },
      })
    ).toBe('table');
    expect(canvasMenuLabelKey('table')).toBe('contextMenu.tableLabel');
    const rows = CANVAS_CONTEXT_MENUS.table.flatMap((r) =>
      r.kind === 'item'
        ? [r.value ? `${r.slot}.${r.value}` : r.slot]
        : r.kind === 'submenu'
          ? r.items.map((i) =>
              i.kind === 'item' ? (i.value ? `${i.slot}.${i.value}` : i.slot) : ''
            )
          : []
    );
    for (const slot of [
      'table.insertRowAbove',
      'table.insertRowBelow',
      'table.insertColumnLeft',
      'table.insertColumnRight',
      'table.delete.rows',
      'table.delete.columns',
      'table.delete.table',
      'table.mergeCells',
      'table.splitCell.cols2',
      'table.select.table',
      'table.select.row',
      'table.select.column',
    ])
      expect(rows).toContain(slot);
  });

  test('a right-click on a table cell places the caret there; a range or a selected table stays', () => {
    const at: Extract<EditorSelection, { kind: 'text' }> = {
      kind: 'text',
      nodeId: 'f',
      cell: { row: 1, column: 0 },
      anchor: caret(0, 1),
      focus: caret(0, 1),
    };
    expect(contextMenuSelection({ kind: 'none' }, 'f', undefined, at)).toEqual(at);
    // A table selected as an object keeps that selection: its menu is the object menu.
    expect(contextMenuSelection({ kind: 'shapes', nodeIds: ['f'] }, 'f', undefined, at)).toBeNull();
    expect(canvasMenuKind({ selection: { kind: 'shapes', nodeIds: ['f'] } })).toBe('shapes');
    // Another shape selected: the cell under the pointer gets the caret.
    expect(contextMenuSelection({ kind: 'shapes', nodeIds: ['g'] }, 'f', undefined, at)).toEqual(
      at
    );
    expect(contextMenuSelection(at, 'f', undefined, { ...at, anchor: caret(0, 0) })).toBeNull();
    const range: EditorSelection = {
      kind: 'cells',
      nodeId: 'f',
      anchor: { row: 0, column: 0 },
      focus: { row: 1, column: 1 },
    };
    expect(contextMenuSelection(range, 'f', undefined, at)).toBeNull();
  });

  test('the Format panel calls a table a table, and a refusal in view mode is information', () => {
    const geometry = { kind: 'graphicFrame' as const, nodeId: 'f' };
    expect(elementKindLabelKey(geometry, null)).toBe('panel.format.elementKind.graphicFrame');
    expect(elementKindLabelKey(geometry, { nodeId: 'f' } as never)).toBe(
      'panel.format.elementKind.table'
    );
    expect(editNoticeTone({ reason: 'disabled.viewing' })).toBe('info');
    expect(editNoticeTone({ reason: 'errors.unsupported' })).toBe('error');
  });
});

describe('text hits', () => {
  test('a point on the first glyph of a box with no inset is text, not the border band', async () => {
    const e = await open();
    const item = textItem(e);
    const c = caretAt(item, caret(0, 0))!;
    // Inside the 8 px band at zoom 1 (76 200 EMU), a quarter of the way into the first glyph.
    const next = caretAt(item, caret(0, 1))!;
    const p = { x: (c.x1 * 3 + next.x1) / 4, y: (c.y1 + c.y2) / 2 };
    expect(p.x - IN).toBeLessThan(76_200);
    expect(onGlyph(item, p)).toBe(true);
    expect(onBorder(item, p, 1)).toBe(false);
    // Left of the text, in the band, the press still moves the box.
    expect(onBorder(item, { x: IN - 20_000, y: p.y }, 1)).toBe(true);
    e.destroy();
  });
});

describe('the view', () => {
  test('a caret box outside the scroller scrolls in by the least distance', () => {
    const scroller = {
      scrollLeft: 100,
      scrollTop: 100,
      getBoundingClientRect: () => ({
        left: 0,
        top: 0,
        right: 400,
        bottom: 300,
        width: 400,
        height: 300,
      }),
    } as unknown as HTMLElement;
    scrollBoxIntoView(scroller, { left: 500, right: 501, top: 10, bottom: 30 }, 32);
    expect(scroller.scrollLeft).toBe(100 + 501 - (400 - 32));
    expect(scroller.scrollTop).toBe(100 - (32 - 10));
    scrollBoxIntoView(scroller, { left: 50, right: 51, top: 100, bottom: 120 }, 32);
    expect([scroller.scrollLeft, scroller.scrollTop]).toEqual([233, 78]);
  });
});

// A DOM for this file only: the engine test run registers no global DOM.
const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

describe('the surface', () => {
  /** Attach the editor; happy-dom lays nothing out, so the stage is 1280 px wide. */
  async function attached() {
    const g = globalThis as { ResizeObserver?: unknown };
    g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
    const e = await open();
    const host = doc.createElement('div') as unknown as HTMLElement;
    doc.body.append(host as unknown as Node);
    e.attach(host);
    const root = host.querySelector('.pptx-surface') as HTMLElement;
    const stage = host.querySelector('.pptx-surface__stage') as HTMLElement;
    stage.getBoundingClientRect = () =>
      ({ left: 0, top: 0, right: 1280, bottom: 720, width: 1280, height: 720 }) as DOMRect;
    const input = host.querySelector('textarea') as HTMLTextAreaElement;
    const width = e.session!.layout(0)!.slide.width;
    const px = (emu: number) => (emu / width) * 1280;
    const key = (target: HTMLElement, k: string, shiftKey = false): boolean => {
      const ev = new win.KeyboardEvent('keydown', {
        key: k,
        shiftKey,
        bubbles: true,
        cancelable: true,
      });
      target.dispatchEvent(ev as unknown as Event);
      return ev.defaultPrevented;
    };
    const press = (x: number, y: number) =>
      stage.dispatchEvent(
        new win.MouseEvent('pointerdown', {
          clientX: px(x),
          clientY: px(y),
          button: 0,
          bubbles: true,
          cancelable: true,
        }) as unknown as Event
      );
    const release = () =>
      doc.dispatchEvent(new win.MouseEvent('pointerup', { bubbles: true }) as unknown as Event);
    const move = (x: number, y: number) =>
      doc.dispatchEvent(
        new win.MouseEvent('pointermove', {
          clientX: px(x),
          clientY: px(y),
          bubbles: true,
        }) as unknown as Event
      );
    const dblclick = (x: number, y: number) =>
      stage.dispatchEvent(
        new win.MouseEvent('dblclick', {
          clientX: px(x),
          clientY: px(y),
          bubbles: true,
        }) as unknown as Event
      );
    return { e, host, root, stage, input, key, press, release, move, dblclick };
  }

  test('Tab on the slide selects shapes in turn; past the last one it lets the key go', async () => {
    const { e, root, key } = await attached();
    expect(key(root, 'Tab')).toBe(true);
    expect(e.snapshot().selection).toEqual({ kind: 'shapes', nodeIds: [idOf(e, 'T')] });
    expect(root.querySelector('[aria-live]')?.textContent).toBe('canvas.shapeSelected');
    expect(key(root, 'Tab')).toBe(true);
    expect(e.snapshot().selection).toEqual({ kind: 'shapes', nodeIds: [idOf(e, 'B')] });
    expect(key(root, 'Tab', true)).toBe(true);
    // Enter on a selected text shape starts text editing; Escape returns to the shape.
    expect(key(root, 'Enter')).toBe(true);
    expect(e.snapshot().selection.kind).toBe('text');
    e.setSelection({ kind: 'shapes', nodeIds: [tableId(e)] });
    expect(key(root, 'Tab')).toBe(false);
    expect(e.snapshot().selection).toEqual({ kind: 'none' });
    // Escape clears a selection.
    e.setSelection({ kind: 'shapes', nodeIds: [idOf(e, 'B')] });
    expect(key(root, 'Escape')).toBe(true);
    expect(e.snapshot().selection).toEqual({ kind: 'none' });
    e.destroy();
  });

  test('a press on the first glyph of a text box being edited places the caret', async () => {
    const { e, press, release } = await attached();
    const id = idOf(e, 'T');
    e.setSelection({ kind: 'text', nodeId: id, anchor: caret(1, 3), focus: caret(1, 3) });
    const item = textItem(e);
    const c = caretAt(item, caret(0, 0))!;
    const next = caretAt(item, caret(0, 1))!;
    press((c.x1 * 3 + next.x1) / 4, (c.y1 + c.y2) / 2);
    release();
    expect(e.snapshot().selection).toEqual({
      kind: 'text',
      nodeId: id,
      anchor: caret(0, 0),
      focus: caret(0, 0),
    });
    e.destroy();
  });

  test('a triple-click selects the paragraph under the pointer', async () => {
    const { e, press, release } = await attached();
    const id = idOf(e, 'T');
    const item = textItem(e);
    const c = caretAt(item, caret(1, 3))!;
    for (let i = 0; i < 3; i += 1) {
      press(c.x1, (c.y1 + c.y2) / 2);
      release();
    }
    expect(e.snapshot().selection).toEqual({
      kind: 'text',
      nodeId: id,
      anchor: caret(1, 0),
      focus: caret(1, 11),
    });
    e.destroy();
  });

  test('a double-click drag extends by whole words, and the late dblclick keeps it', async () => {
    const { e, press, release, move, dblclick } = await attached();
    const id = idOf(e, 'T');
    const item = textItem(e);
    const mid = (c: { y1: number; y2: number }) => (c.y1 + c.y2) / 2;
    // Press twice on "Hello", then drag into "world": both whole words.
    const start = caretAt(item, caret(0, 2))!;
    press(start.x1, mid(start));
    release();
    press(start.x1, mid(start));
    expect(e.snapshot().selection).toMatchObject({ anchor: caret(0, 0), focus: caret(0, 5) });
    const end = caretAt(item, caret(0, 8))!;
    move(end.x1, mid(end));
    release();
    dblclick(end.x1, mid(end));
    expect(e.snapshot().selection).toEqual({
      kind: 'text',
      nodeId: id,
      anchor: caret(0, 0),
      focus: caret(0, 11),
    });
    // Back over the start: the pressed word stays whole, now from its end.
    press(end.x1, mid(end));
    release();
    press(end.x1, mid(end));
    move(start.x1, mid(start));
    release();
    expect(e.snapshot().selection).toMatchObject({ anchor: caret(0, 11), focus: caret(0, 0) });
    e.destroy();
  });

  test('in a child of an entered group, the second press takes a word, a fourth a caret', async () => {
    const { e, press, release } = await attached();
    const [g, g1] = ['G', 'G1'].map((n) => idOf(e, n));
    e.setSelection({
      kind: 'text',
      nodeId: g1!,
      anchor: caret(0, 2),
      focus: caret(0, 2),
    });
    expect(e.snapshot().selection.kind).toBe('text');
    const child = e.session!.layout(0)!.items.find((i) => i.element.nodeId === g)!;
    const item =
      child.kind === 'group' ? child.children.find((c) => c.element.nodeId === g1)! : child;
    const c = caretAt(item, caret(0, 2))!;
    const at = (): void => void press(c.x1, (c.y1 + c.y2) / 2);
    at();
    release();
    at();
    release();
    expect(e.snapshot().selection).toMatchObject({ anchor: caret(0, 0), focus: caret(0, 5) });
    at();
    release();
    expect(e.snapshot().selection).toMatchObject({ anchor: caret(0, 0), focus: caret(0, 10) });
    at();
    release();
    // The fourth press goes round to a caret, as DOCX Editor.
    const s = e.snapshot().selection;
    expect(s.kind === 'text' && s.anchor.offset === s.focus.offset).toBe(true);
    e.destroy();
  });

  test('a double-click on a table grid line selects no word', async () => {
    const { e, press, release, dblclick } = await attached();
    e.setZoom(1);
    const id = tableId(e);
    const at = { paragraph: 0, offset: 1 };
    e.setSelection({
      kind: 'text',
      nodeId: id,
      cell: { row: 0, column: 0 },
      anchor: at,
      focus: at,
    });
    // The line between the two columns, halfway down the first row.
    const x = 7 * IN;
    const y = 4 * IN + 370_840 / 2;
    press(x, y);
    release();
    press(x, y);
    release();
    dblclick(x, y);
    const s = e.snapshot().selection;
    expect(s.kind === 'text' && s.anchor.offset === s.focus.offset).toBe(s.kind === 'text');
    e.destroy();
  });

  test('a triple-click drag extends by whole paragraphs', async () => {
    const { e, press, release, move } = await attached();
    const item = textItem(e);
    const c = caretAt(item, caret(0, 3))!;
    for (let i = 0; i < 2; i += 1) {
      press(c.x1, (c.y1 + c.y2) / 2);
      release();
    }
    press(c.x1, (c.y1 + c.y2) / 2);
    const below = caretAt(item, caret(1, 2))!;
    move(below.x1, (below.y1 + below.y2) / 2);
    release();
    expect(e.snapshot().selection).toMatchObject({ anchor: caret(0, 0), focus: caret(1, 11) });
    e.destroy();
  });

  test('the IME input sits in the stage on the caret, as tall as the caret line', async () => {
    const { e, stage, input } = await attached();
    expect(input.parentElement === stage).toBe(true);
    const id = idOf(e, 'T');
    e.setSelection({ kind: 'text', nodeId: id, anchor: caret(0, 5), focus: caret(0, 5) });
    const c = caretAt(textItem(e), caret(0, 5))!;
    const scale = e.snapshot().zoom / 9525;
    expect(parseFloat(input.style.left)).toBeCloseTo(c.x1 * scale, 3);
    expect(parseFloat(input.style.top)).toBeCloseTo(c.y1 * scale, 3);
    expect(parseFloat(input.style.height)).toBeCloseTo((c.y2 - c.y1) * scale, 3);
    e.destroy();
  });

  test('the overlay dims the area outside the slide, with the slide as a hole', async () => {
    const { e, host } = await attached();
    const veil = host.querySelector('.pptx-surface__outside');
    expect(veil?.getAttribute('fill-rule')).toBe('evenodd');
    const { width, height } = e.session!.layout(0)!.slide;
    expect(veil?.getAttribute('d')).toContain(`M0 0V${height}H${width}V0Z`);
    e.destroy();
  });
});
