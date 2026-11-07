// A shape drag previews while the pointer moves and writes once on release: the store, the
// layout, and the slide paint stay as they were until then, the release is one undo step, and
// Escape writes nothing. The preview puts each shape where the release does.

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
import { buildSlideStructureDeck } from '../../../../../scripts/fixtures/slide-structure.ts';
import { landscapePng } from '../../../../../scripts/sample-deck/showcase-image.ts';
import { apply, type LayoutItem } from '../../layout/index.ts';
import { relayoutItem } from '../../layout/item-layout.ts';
import { cornersOf, findItem, lineEnds } from '../../automation/commands/placement.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { dragStep, startDrag, type Handle } from '../surface-drag.ts';
import { preview, previewLayout } from '../drag-preview.ts';

const IN = 914_400;

/**
 * A box, a text box, a group of two with a scaled child space, a horizontal line, a vertical
 * line, and a flipped diagonal line.
 */
function deck(): Uint8Array {
  const group =
    '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm rot="1800000"><a:off x="914400" y="4572000"/><a:ext cx="1828800" cy="914400"/>' +
    '<a:chOff x="0" y="0"/><a:chExt cx="3657600" cy="1828800"/></a:xfrm></p:grpSpPr>' +
    shape({ id: 11, name: 'G1', x: 0, y: 0, cx: 2 * IN, cy: 2 * IN }) +
    shape({ id: 12, name: 'G2', x: 2 * IN, y: 0, cx: 2 * IN, cy: 2 * IN }) +
    '</p:grpSp>';
  // A group drawn at twice its child size, with a red group fill (`a:grpFill`): its child S
  // takes the group's fill and has a 1 pt outline.
  const scaledGroup =
    '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="30" name="S2"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="5486400" y="2743200"/><a:ext cx="1828800" cy="914400"/>' +
    '<a:chOff x="0" y="0"/><a:chExt cx="914400" cy="457200"/></a:xfrm>' +
    '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:grpSpPr>' +
    shape({
      id: 31,
      name: 'S',
      x: 0,
      y: 0,
      cx: IN / 2,
      cy: IN / 2,
      fill: '<a:grpFill/>',
      line: '<a:ln w="12700"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln>',
    }) +
    '</p:grpSp>';
  return buildPptx({
    slides: [
      {
        shapes:
          shape({ id: 3, name: 'B', x: 8 * IN, y: IN, cx: IN, cy: IN }) +
          shape({
            id: 2,
            name: 'T',
            x: IN,
            y: IN,
            cx: 4 * IN,
            cy: IN,
            bodyPr: '<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>',
            text: para(run('Hello world')),
          }) +
          group +
          shape({ id: 20, name: 'H', x: 6 * IN, y: 6.8 * IN, cx: 2 * IN, cy: 0, prst: 'line' }) +
          shape({ id: 21, name: 'V', x: 9 * IN, y: 5.5 * IN, cx: 0, cy: IN, prst: 'line' }) +
          shape({
            id: 22,
            name: 'D',
            x: 10.5 * IN,
            y: 5.5 * IN,
            cx: IN,
            cy: IN,
            prst: 'line',
            flipH: true,
          }) +
          scaledGroup +
          table({
            id: 40,
            x: 9.5 * IN,
            y: 3 * IN,
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

const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const LINES = ['One line of text', 'and another line', 'and a third line', 'and one more'];

/**
 * Autofit text boxes (`spAutoFit`, `normAutofit`, a title placeholder with `normAutofit`), an
 * empty body placeholder (it paints its prompt), a cropped picture, and a table in a group
 * drawn at twice its child size.
 */
function autofitDeck(): Uint8Array {
  const text = LINES.map((l) => para(run(l, 'sz="2400"'))).join('');
  const pic =
    '<p:pic><p:nvPicPr><p:cNvPr id="7" name="P"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
    '<p:blipFill><a:blip r:embed="rIdImg"/><a:srcRect l="25000" t="10000" r="10000" b="20000"/>' +
    '<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="8229600" y="914400"/>' +
    '<a:ext cx="1828800" cy="1371600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
  const grouped =
    '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="50" name="TG"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="8229600" y="3657600"/><a:ext cx="3657600" cy="1483360"/>' +
    '<a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="741680"/></a:xfrm></p:grpSpPr>' +
    table({
      id: 51,
      x: 0,
      y: 0,
      cx: 2 * IN,
      cy: 2 * 370_840,
      columns: [IN, IN],
      tblPr: '<a:tblPr/>',
      rows: [0, 1]
        .map((r) => `<a:tr h="370840">${cell(`r${r}a`)}${cell(`r${r}b`)}</a:tr>`)
        .join(''),
    }) +
    '</p:grpSp>';
  return buildPptx({
    media: [{ name: 'image1.png', bytes: landscapePng(40, 25), contentType: 'image/png' }],
    slides: [
      {
        layout: 'blank',
        rels: [['rIdImg', R_IMAGE, '../media/image1.png']],
        shapes:
          shape({ id: 2, name: 'A', x: IN, y: IN, cx: 3 * IN, cy: IN, text }) +
          shape({
            id: 3,
            name: 'N',
            x: IN,
            y: 4 * IN,
            cx: 3 * IN,
            cy: 2 * IN,
            bodyPr: '<a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr>',
            text,
          }) +
          shape({
            id: 4,
            name: 'Ti',
            x: 4.5 * IN,
            y: IN,
            cx: 3 * IN,
            cy: 1.2 * IN,
            ph: '<p:ph type="title"/>',
            bodyPr: '<a:bodyPr><a:normAutofit/></a:bodyPr>',
            text: para(run('A title that wraps', 'sz="4000"')),
          }) +
          shape({
            id: 5,
            name: 'Pr',
            x: 4.5 * IN,
            y: 4 * IN,
            cx: 3 * IN,
            cy: 2 * IN,
            ph: '<p:ph type="body" idx="1"/>',
            bodyPr: '<a:bodyPr/>',
            text: '<a:p><a:endParaRPr lang="en-US"/></a:p>',
          }) +
          pic +
          grouped,
      },
      {
        // A group drawn at half its child size: `normAutofit` and `spAutoFit` children fit at
        // the size the group draws them.
        layout: 'blank',
        shapes:
          '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="60" name="HG"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
          '<p:grpSpPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="1371600"/>' +
          '<a:chOff x="0" y="0"/><a:chExt cx="10972800" cy="2743200"/></a:xfrm></p:grpSpPr>' +
          shape({
            id: 61,
            name: 'HN',
            x: 0,
            y: 0,
            cx: 6 * IN,
            cy: 3 * IN,
            bodyPr: '<a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr>',
            text,
          }) +
          shape({ id: 62, name: 'HA', x: 6.2 * IN, y: 0, cx: 5.8 * IN, cy: 3 * IN, text }) +
          '</p:grpSp>',
      },
    ],
  });
}

// A DOM for this file only: the engine test run registers no global DOM.
const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());
const frame = (): Promise<void> => new Promise((r) => setTimeout(r, 40));

const idOf = (e: PptxEditorInstance, name: string): string => {
  const walk = (items: readonly LayoutItem[]): string | undefined => {
    for (const i of items) {
      if (i.element.name === name) return i.element.nodeId;
      if (i.kind === 'group') {
        const inner = walk(i.children);
        if (inner) return inner;
      }
    }
    return undefined;
  };
  return walk(e.session!.layout(0)!.items)!;
};
const itemOf = (e: PptxEditorInstance, name: string): LayoutItem =>
  findItem(e.session!.layout(0)!.items, idOf(e, name))!;

async function attached(bytes: Uint8Array = deck()) {
  const g = globalThis as { ResizeObserver?: unknown };
  g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  const host = doc.createElement('div') as unknown as HTMLElement;
  doc.body.append(host as unknown as Node);
  e.attach(host);
  // happy-dom lays nothing out: without a zoom, the fit zoom (and the snap distance) is off.
  e.setZoom(1);
  const stage = host.querySelector('.pptx-surface__stage') as HTMLElement;
  stage.getBoundingClientRect = () =>
    ({ left: 0, top: 0, right: 1280, bottom: 720, width: 1280, height: 720 }) as DOMRect;
  const width = e.session!.layout(0)!.slide.width;
  const px = (emu: number) => (emu / width) * 1280;
  const pointer = (type: string, target: EventTarget, x: number, y: number) =>
    target.dispatchEvent(
      new win.MouseEvent(type, {
        clientX: px(x),
        clientY: px(y),
        button: 0,
        bubbles: true,
        cancelable: true,
      }) as unknown as Event
    );
  const press = (x: number, y: number, target: EventTarget = stage) =>
    pointer('pointerdown', target, x, y);
  const move = (x: number, y: number) => pointer('pointermove', doc, x, y);
  const release = (x: number, y: number) => pointer('pointerup', doc, x, y);
  const escape = () =>
    host.querySelector('.pptx-surface')!.dispatchEvent(
      new win.KeyboardEvent('keydown', {
        key: 'Escape',
        bubbles: true,
        cancelable: true,
      }) as unknown as Event
    );
  const painted = (name: string) =>
    host.querySelector(`.pptx-surface__slide [data-node-id="${idOf(e, name)}"]`)!;
  /** The preview painted for a shape, or null. */
  const shown = (name: string) => host.querySelector(`[data-preview-of="${idOf(e, name)}"]`);
  const handle = (h: string) => host.querySelector(`[data-handle="${h}"]`)!;
  const cancel = () =>
    doc.dispatchEvent(new win.MouseEvent('pointercancel', { bubbles: true }) as unknown as Event);
  return { e, host, press, move, release, escape, painted, shown, handle, cancel };
}

describe('a shape drag', () => {
  test('moves write nothing until the release, which is one undo step', async () => {
    const { e, host, press, move, release, painted, shown } = await attached();
    const before = e.session!.store.package;
    const slide = host.querySelector('.pptx-surface__slide');
    let emits = 0;
    e.subscribe(() => void (emits += 1));
    press(8.5 * IN, 1.5 * IN);
    emits = 0;
    const layouts = e.session!.layout(0);
    for (let i = 1; i <= 5; i += 1) move(8.5 * IN + i * 0.2 * IN, 1.5 * IN);
    // Nothing is written or laid out, and the slide is not painted again: the shape is hidden,
    // and its preview is painted after it.
    expect(e.session!.store.package).toBe(before);
    expect(e.session!.layout(0)).toBe(layouts);
    expect(host.querySelector('.pptx-surface__slide')).toBe(slide);
    expect(emits).toBe(0);
    expect(painted('B').getAttribute('display')).toBe('none');
    expect(painted('B').nextElementSibling?.getAttribute('data-kind')).toBe('drag-preview');
    // At the new box: its placement names the new left edge.
    expect(shown('B')?.outerHTML).toContain(` ${9 * IN} `);
    // The selection box follows the preview.
    const outline = host.querySelector('.pptx-surface__overlay .pptx-surface__outline')!;
    const at = outline.parentElement!.getAttribute('transform')!;
    expect(at).toContain(`${9 * IN}`);
    // The box's top stays on the text box's top: the smart guide shows.
    expect(host.querySelector('.pptx-surface__overlay .pptx-surface__guide')).not.toBeNull();

    const hidden = painted('B');
    release(9.5 * IN, 1.5 * IN);
    expect(itemOf(e, 'B').element.xfrm).toMatchObject({ x: 9 * IN, y: IN });
    expect(host.querySelector('.pptx-surface__overlay .pptx-surface__guide')).toBeNull();
    // The moved shape paints again at its new place, and the preview is gone.
    expect(painted('B')).not.toBe(hidden);
    expect(painted('B').getAttribute('display')).toBeNull();
    expect(shown('B')).toBeNull();
    e.exec({ type: 'undo' });
    expect(itemOf(e, 'B').element.xfrm).toMatchObject({ x: 8 * IN, y: IN });
    expect(e.snapshot().canUndo).toBe(false);
    e.destroy();
  });

  test('Escape drops the preview, writes nothing, and keeps the selection', async () => {
    const { e, host, press, move, release, escape, painted, shown } = await attached();
    press(8.5 * IN, 1.5 * IN);
    move(9 * IN, 2 * IN);
    expect(shown('B')).not.toBeNull();
    escape();
    expect(shown('B')).toBeNull();
    // The region is cleared, then says the cancel on the next frame.
    await frame();
    expect(host.querySelector('[aria-live]')?.textContent).toBe('canvas.dragCanceled');
    expect(painted('B').hasAttribute('display')).toBe(false);
    expect(host.querySelector('.pptx-surface__overlay .pptx-surface__guide')).toBeNull();
    expect(e.snapshot().selection).toEqual({ kind: 'shapes', nodeIds: [idOf(e, 'B')] });
    // The release after Escape moves nothing.
    move(9.5 * IN, 2.5 * IN);
    release(9.5 * IN, 2.5 * IN);
    expect(itemOf(e, 'B').element.xfrm).toMatchObject({ x: 8 * IN, y: IN });
    expect(e.snapshot().canUndo).toBe(false);
    e.destroy();
  });

  test('Escape after a press with no move says nothing about a drag', async () => {
    const { e, host, press, escape, release } = await attached();
    press(8.5 * IN, 1.5 * IN);
    escape();
    await frame();
    expect(host.querySelector('[aria-live]')?.textContent).not.toBe('canvas.dragCanceled');
    release(8.5 * IN, 1.5 * IN);
    e.destroy();
  });

  test('an undo during the drag ends it, and the release writes nothing', async () => {
    const { e, host, press, move, release, shown } = await attached();
    // One edit to undo, on another shape: T moves down.
    e.setSelection({ kind: 'shapes', nodeIds: [idOf(e, 'T')] });
    expect(e.exec({ type: 'nudge', dx: 0, dy: IN / 4 }).ok).toBe(true);
    press(8.5 * IN, 1.5 * IN);
    move(9.75 * IN, 2 * IN);
    expect(shown('B')).not.toBeNull();
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(shown('B')).toBeNull();
    // The region is cleared, then says the cancel on the next frame.
    await frame();
    expect(host.querySelector('[aria-live]')?.textContent).toBe('canvas.dragCanceled');
    release(9.75 * IN, 2 * IN);
    expect(itemOf(e, 'B').element.xfrm).toMatchObject({ x: 8 * IN, y: IN });
    expect(itemOf(e, 'T').element.xfrm).toMatchObject({ x: IN, y: IN });
    expect(e.snapshot().canRedo).toBe(true);
    e.destroy();
  });

  test('a host command on another shape, or on the notes, keeps the drag', async () => {
    const { e, host, press, move, release, shown } = await attached();
    press(8.5 * IN, 1.5 * IN);
    move(9 * IN, 1.5 * IN);
    expect(e.exec({ type: 'setNotes', index: 0, text: 'Said aloud' }).ok).toBe(true);
    e.setSelection({ kind: 'shapes', nodeIds: [idOf(e, 'T')] });
    expect(e.exec({ type: 'setShapeFill', color: '#70AD47' }).ok).toBe(true);
    expect(shown('B')).not.toBeNull();
    expect(host.querySelector('[aria-live]')?.textContent).not.toBe('canvas.dragCanceled');
    release(9 * IN, 1.5 * IN);
    expect(itemOf(e, 'B').element.xfrm).toMatchObject({ x: 8.5 * IN, y: IN });
    e.destroy();
  });

  test('an empty placeholder previews with its prompt', async () => {
    const { e, press, move, release, painted, shown, handle } = await attached(autofitDeck());
    const prompt = painted('Pr').textContent ?? '';
    expect(prompt.length).toBeGreaterThan(0);
    e.setSelection({ kind: 'shapes', nodeIds: [idOf(e, 'Pr')] });
    press(7.5 * IN, 6 * IN, handle('se'));
    move(8 * IN, 6.4 * IN);
    expect(shown('Pr')?.textContent).toBe(prompt);
    release(8 * IN, 6.4 * IN);
    e.destroy();
  });

  test("a change to a dragged child's group cancels the drag", async () => {
    const { e, press, move, release, shown } = await attached();
    const g1 = cornersOf(itemOf(e, 'G1'));
    const center = { x: (g1[0]!.x + g1[2]!.x) / 2, y: (g1[0]!.y + g1[2]!.y) / 2 };
    e.setSelection({ kind: 'shapes', nodeIds: [idOf(e, 'G1')], group: idOf(e, 'G') });
    const before = itemOf(e, 'G1').element.xfrm;
    press(center.x, center.y);
    move(center.x + 0.4 * IN, center.y);
    expect(shown('G1')).not.toBeNull();
    // The host moves the group; the child's own element does not change.
    e.setSelection({ kind: 'shapes', nodeIds: [idOf(e, 'G')] });
    expect(e.exec({ type: 'nudge', dx: 0, dy: IN }).ok).toBe(true);
    expect(itemOf(e, 'G1').element.xfrm).toEqual(before);
    expect(shown('G1')).toBeNull();
    release(center.x + 0.4 * IN, center.y);
    expect(itemOf(e, 'G1').element.xfrm).toEqual(before);
    e.destroy();
  });

  test('each cancel is said, also the same words twice', async () => {
    const { e, host, press, move, release, escape } = await attached();
    const live = () => host.querySelector('[aria-live]')?.textContent;
    for (let i = 0; i < 2; i += 1) {
      press(8.5 * IN, 1.5 * IN);
      move(9 * IN, 2 * IN);
      escape();
      // Cleared at once, said on the next frame.
      expect(live()).toBe('');
      await frame();
      expect(live()).toBe('canvas.dragCanceled');
      release(9 * IN, 2 * IN);
    }
    e.destroy();
  });

  test('the window losing focus cancels as Escape does', async () => {
    const { e, press, move, release, shown } = await attached();
    press(8.5 * IN, 1.5 * IN);
    move(9 * IN, 2 * IN);
    (win as unknown as EventTarget).dispatchEvent(new win.Event('blur') as unknown as Event);
    expect(shown('B')).toBeNull();
    release(9 * IN, 2 * IN);
    expect(itemOf(e, 'B').element.xfrm).toMatchObject({ x: 8 * IN, y: IN });
    e.destroy();
  });

  test('pointercancel writes the preview as one undo step', async () => {
    const { e, press, move, cancel, shown } = await attached();
    press(8.5 * IN, 1.5 * IN);
    move(8.9 * IN, 1.5 * IN);
    move(9.3 * IN, 1.5 * IN);
    cancel();
    expect(shown('B')).toBeNull();
    expect(itemOf(e, 'B').element.xfrm.x).toBeCloseTo(8.8 * IN, -1);
    e.exec({ type: 'undo' });
    expect(itemOf(e, 'B').element.xfrm).toMatchObject({ x: 8 * IN, y: IN });
    e.destroy();
  });

  test('a child of a scaled group with a group fill previews with its fill and outline', async () => {
    const { e, host, press, move, release, handle } = await attached();
    e.setSelection({ kind: 'shapes', nodeIds: [idOf(e, 'S')], group: idOf(e, 'S2') });
    const paints = (root: Element | null) =>
      [...(root?.querySelectorAll('path') ?? [])].map((p) => [
        p.getAttribute('fill'),
        p.getAttribute('stroke'),
        p.getAttribute('stroke-width'),
      ]);
    // S covers 6 to 7 inches across and 3 to 4 down on the slide.
    press(7 * IN, 4 * IN, handle('se'));
    move(7.3 * IN, 4.2 * IN);
    move(7.5 * IN, 4.3 * IN);
    const during = paints(host.querySelector(`[data-preview-of="${idOf(e, 'S')}"]`));
    expect(during.length).toBeGreaterThan(0);
    expect(during.some(([fill]) => fill && fill !== 'none')).toBe(true);
    // No hit targets in the preview: the slide keeps one element for each shape.
    expect(host.querySelectorAll(`[data-node-id="${idOf(e, 'S')}"]`)).toHaveLength(1);
    release(7.5 * IN, 4.3 * IN);
    const after = paints(
      host.querySelector(`.pptx-surface__slide [data-node-id="${idOf(e, 'S')}"]`)
    );
    expect(during).toEqual(after);
    e.destroy();
  });

  test('a resize and a rotation by their handles write once on release', async () => {
    const { e, press, move, release, handle } = await attached();
    e.setSelection({ kind: 'shapes', nodeIds: [idOf(e, 'B')] });
    press(9 * IN, 2 * IN, handle('se'));
    move(9.5 * IN, 2.25 * IN);
    move(10 * IN, 2.5 * IN);
    expect(itemOf(e, 'B').element.xfrm).toMatchObject({ cx: IN, cy: IN });
    release(10 * IN, 2.5 * IN);
    expect(itemOf(e, 'B').element.xfrm).toMatchObject({ cx: 2 * IN, cy: 1.5 * IN });

    const top = itemOf(e, 'B');
    const c = apply(top.matrix, IN, 0.75 * IN);
    press(c.x, c.y - IN, handle('rot'));
    move(c.x + IN, c.y);
    expect(itemOf(e, 'B').element.xfrm.rot).toBe(0);
    release(c.x + IN, c.y);
    expect(itemOf(e, 'B').element.xfrm.rot).toBeCloseTo(90, 5);
    e.exec({ type: 'undo' });
    e.exec({ type: 'undo' });
    expect(itemOf(e, 'B').element.xfrm).toMatchObject({ cx: IN, cy: IN, rot: 0 });
    e.destroy();
  });
});

describe('the preview', () => {
  /** Drag `names` by `handle` from `from` to `to`; compare the preview with the write. */
  async function check(
    names: string[],
    handle: Handle | 'move',
    from: number[],
    to: number[],
    opts: { bytes?: Uint8Array; slide?: number; ids?: (e: PptxEditorInstance) => string[] } = {}
  ) {
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await e.load(opts.bytes ?? deck())).ok).toBe(true);
    const slide = opts.slide ?? 0;
    e.goToSlide(slide);
    const layout = e.session!.layout(slide)!;
    const ids = opts.ids?.(e) ?? names.map((n) => idOf(e, n));
    const d = startDrag(layout, handle, ids, { x: from[0], y: from[1] }, 'g')!;
    const step = dragStep(
      d,
      { x: to[0], y: to[1] },
      { shift: false, snapTo: null, aspectLocked: () => false }
    );
    preview(d, step.items, (item, xfrm) => relayoutItem(item, xfrm, e.session!.measurer));
    const shown = previewLayout(layout, d);
    const r = e.exec({ type: 'setTransforms', items: step.items });
    expect(r.ok).toBe(true);
    const after = e.session!.layout(slide)!;
    // The file rounds to whole EMU.
    const near = (a: { x: number; y: number }, b: { x: number; y: number }) => {
      expect(Math.abs(a.x - b.x)).toBeLessThan(2);
      expect(Math.abs(a.y - b.y)).toBeLessThan(2);
    };
    for (const id of ids) {
      const want = findItem(after.items, id)!;
      const got = findItem(shown.items, id)!;
      const corners = cornersOf(got);
      cornersOf(want).forEach((c, k) => near(corners[k]!, c));
      // A line's end points, where the end handles are drawn.
      const ends = lineEnds(want);
      if (ends) {
        near(lineEnds(got)!.start, ends.start);
        near(lineEnds(got)!.end, ends.end);
      }
      // An object's picture is drawn in the frame's box.
      if (want.kind === 'frame' && want.poster && got.kind === 'frame') {
        const drawn = cornersOf(got.poster!);
        cornersOf(want.poster).forEach((c, k) => near(drawn[k]!, c));
      }
      // A table takes the columns and rows the release writes.
      if (want.kind === 'frame' && want.table && got.kind === 'frame') {
        want.table.columnX.forEach((x, k) =>
          expect(Math.abs(got.table!.columnX[k]! - x)).toBeLessThan(2)
        );
        want.table.rowY.forEach((y, k) =>
          expect(Math.abs(got.table!.rowY[k]! - y)).toBeLessThan(2)
        );
      }
      // Autofit text takes the scale the release writes.
      if (want.kind === 'shape' && got.kind === 'shape')
        expect(got.element.text?.body.autofit).toEqual(want.element.text?.body.autofit);
      // A picture keeps its crop.
      if (want.kind === 'picture' && got.kind === 'picture')
        expect(got.element.fill).toEqual(want.element.fill);
      // Text reflows in the new box as the release lays it out.
      if (want.kind === 'shape' && want.text && got.kind === 'shape') {
        expect(got.text?.lines.length).toBe(want.text.lines.length);
        expect(Math.abs(got.text!.box.width - want.text.box.width)).toBeLessThan(2);
      }
    }
    e.destroy();
  }

  test('a move, a resize, and a rotation land where the release writes them', async () => {
    await check(['B'], 'move', [8.5 * IN, 1.5 * IN], [9.1 * IN, 2.3 * IN]);
    await check(['B', 'T'], 'move', [8.5 * IN, 1.5 * IN], [7.9 * IN, 2.3 * IN]);
    await check(['B'], 'se', [9 * IN, 2 * IN], [10.2 * IN, 2.7 * IN]);
    await check(['B'], 'nw', [8 * IN, IN], [7.5 * IN, 0.4 * IN]);
    await check(['B', 'T'], 'e', [9 * IN, 1.5 * IN], [10 * IN, 1.5 * IN]);
    await check(['B'], 'rot', [8.5 * IN, 0.5 * IN], [9.7 * IN, 1.5 * IN]);
  });

  test('a table resize lays the table out as the release does', async () => {
    await check(['Table 39'], 'move', [10 * IN, 3.2 * IN], [10.4 * IN, 3.5 * IN]);
    await check(['Table 39'], 'se', [11.5 * IN, 3.8 * IN], [12.3 * IN, 4.6 * IN]);
    await check(['Table 39'], 'e', [11.5 * IN, 3.4 * IN], [10.8 * IN, 3.4 * IN]);
    await check(['Table 39'], 's', [10.5 * IN, 3.8 * IN], [10.5 * IN, 3.5 * IN]);
  });

  test("an object's picture moves and resizes with its frame", async () => {
    const ole = (e: PptxEditorInstance) => [
      e.session!.layout(3)!.items.find((i) => i.kind === 'frame' && i.element.frameKind === 'ole')!
        .element.nodeId,
    ];
    const opts = { bytes: buildSlideStructureDeck(), slide: 3, ids: ole };
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await e.load(opts.bytes)).ok).toBe(true);
    const c = cornersOf(findItem(e.session!.layout(3)!.items, ole(e)[0]!)!);
    e.destroy();
    await check([], 'move', [c[0]!.x, c[0]!.y], [c[0]!.x + 0.5 * IN, c[0]!.y + 0.25 * IN], opts);
    await check([], 'se', [c[2]!.x, c[2]!.y], [c[2]!.x + 0.6 * IN, c[2]!.y + 0.4 * IN], opts);
  });

  test('autofit text sizes or scales in the preview as the release writes it', async () => {
    const opts = { bytes: autofitDeck() };
    // spAutoFit: narrower, the box grows taller; wider, it shrinks.
    await check(['A'], 'e', [4 * IN, 1.5 * IN], [2.4 * IN, 1.5 * IN], opts);
    await check(['A'], 'e', [4 * IN, 1.5 * IN], [6 * IN, 1.5 * IN], opts);
    // normAutofit: a smaller box scales the text down; a title takes its own steps.
    await check(['N'], 's', [2.5 * IN, 6 * IN], [2.5 * IN, 5 * IN], opts);
    await check(['N'], 'se', [4 * IN, 6 * IN], [3 * IN, 5.2 * IN], opts);
    await check(['Ti'], 'e', [7.5 * IN, 1.6 * IN], [6 * IN, 1.6 * IN], opts);
  });

  test('autofit children of a scaled group fit at the size the group draws them', async () => {
    const named = (name: string) => (e: PptxEditorInstance) => {
      const group = e.session!.layout(1)!.items[0]!;
      const child = group.kind === 'group' ? group.children : [];
      return [child.find((c) => c.element.name === name)!.element.nodeId];
    };
    const at = (name: string) => ({ bytes: autofitDeck(), slide: 1, ids: named(name) });
    // HN is drawn from (1 in, 1 in) to (4 in, 2.5 in); HA from (4.1 in, 1 in) to (7 in, 2.5 in).
    await check([], 'se', [4 * IN, 2.5 * IN], [3.4 * IN, 2.5 * IN], at('HN'));
    await check([], 's', [2.5 * IN, 2.5 * IN], [2.5 * IN, 2 * IN], at('HN'));
    await check([], 'e', [7 * IN, 1.75 * IN], [6.2 * IN, 1.75 * IN], at('HA'));
    await check([], 'e', [7 * IN, 1.75 * IN], [7.6 * IN, 1.75 * IN], at('HA'));
  });

  test('a cropped picture and a table in a scaled group land where the release writes them', async () => {
    const opts = { bytes: autofitDeck() };
    await check(['P'], 'se', [10.2 * IN, 2.5 * IN], [11 * IN, 3 * IN], opts);
    await check(['P'], 'w', [9 * IN, 1.75 * IN], [8.4 * IN, 1.75 * IN], opts);
    await check(['Table 50'], 'se', [12 * IN, 5.27 * IN], [12.6 * IN, 5.8 * IN], opts);
    await check(['Table 50'], 'e', [12 * IN, 4.4 * IN], [11.2 * IN, 4.4 * IN], opts);
  });

  test('a text box resize reflows its text in the preview', async () => {
    await check(['T'], 'e', [5 * IN, 1.5 * IN], [2.2 * IN, 1.5 * IN]);
    await check(['T'], 'w', [IN, 1.5 * IN], [3.5 * IN, 1.5 * IN]);
  });

  test('a resize past the opposite edge flips the box as the release does', async () => {
    await check(['B'], 'e', [9 * IN, 1.5 * IN], [7.4 * IN, 1.5 * IN]);
    await check(['B'], 'se', [9 * IN, 2 * IN], [7.6 * IN, 0.5 * IN]);
  });

  test('a line end follows the pointer: horizontal, vertical, and flipped lines', async () => {
    await check(['H'], 'end', [8 * IN, 6.8 * IN], [8.5 * IN, 7.1 * IN]);
    await check(['H'], 'start', [6 * IN, 6.8 * IN], [5.5 * IN, 6.8 * IN]);
    await check(['V'], 'start', [9 * IN, 5.5 * IN], [8.6 * IN, 5.2 * IN]);
    await check(['V'], 'end', [9 * IN, 6.5 * IN], [9 * IN, 7 * IN]);
    await check(['D'], 'end', [10.5 * IN, 6.5 * IN], [10.2 * IN, 6.9 * IN]);
    await check(['D'], 'end', [10.5 * IN, 6.5 * IN], [12 * IN, 6.9 * IN]);
    await check(['H', 'V'], 'move', [7 * IN, 6.8 * IN], [7.2 * IN, 6.5 * IN]);
  });

  test('a child of a rotated, scaled group lands where the release writes it', async () => {
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await e.load(deck())).ok).toBe(true);
    const g1 = cornersOf(itemOf(e, 'G1'));
    e.destroy();
    await check(['G1'], 'move', [g1[0].x, g1[0].y], [g1[0].x + 0.3 * IN, g1[0].y + 0.2 * IN]);
    await check(['G1'], 'se', [g1[2].x, g1[2].y], [g1[2].x + 0.4 * IN, g1[2].y + 0.1 * IN]);
    await check(['G1'], 'rot', [g1[0].x, g1[0].y - IN], [g1[0].x + IN, g1[0].y]);
    await check(['G'], 'move', [g1[0].x, g1[0].y], [g1[0].x - 0.3 * IN, g1[0].y + 0.4 * IN]);
    await check(['G'], 'se', [g1[2].x, g1[2].y], [g1[2].x + 0.3 * IN, g1[2].y + 0.3 * IN]);
  });
});
