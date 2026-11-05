// Regression tests for the editing defects E11, E13, E16, and E17,
// and for refusals that only say a command had nothing to do. Each test failed
// before its fix.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { attr, elements, find, xmlPartOf, type Element } from '../../store/index.ts';
import type { LayoutItem } from '../../layout/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { MAX_COORDINATE, MIN_COORDINATE } from '../../automation/commands/package-edits.ts';
import {
  boxOf,
  findItem,
  rotationAfterDrag,
  rotationHandle,
  slideCenter,
} from '../../automation/commands/placement.ts';

const IN = 914_400;
const PX = 9_525;

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const partOf = (e: PptxEditorInstance, i = 0): string => e.session!.info.slides[i].partName;
const tree = (e: PptxEditorInstance, i = 0): Element =>
  find(xmlPartOf(e.session!.store.package, partOf(e, i))!.root, 'p:spTree')!;
const nameOf = (el: Element): string | undefined => attr(find(el, 'p:cNvPr'), 'name');
const idByName = (e: PptxEditorInstance, name: string): string =>
  elements(tree(e)).find((c) => nameOf(c) === name)!.id;
/** Shape names bottom to top, and the `p:extLst` of the tree where it is. */
const stack = (e: PptxEditorInstance): string[] =>
  elements(tree(e))
    .filter((c) => c.local === 'sp' || c.local === 'extLst')
    .map((c) => (c.local === 'extLst' ? 'extLst' : nameOf(c)!));
const itemOf = (e: PptxEditorInstance, name: string): LayoutItem =>
  findItem(e.session!.layout(0)!.items, idByName(e, name))!;
const rotOf = (e: PptxEditorInstance, name: string): number =>
  Number(
    attr(
      find(
        elements(tree(e)).find((c) => nameOf(c) === name),
        'a:xfrm'
      ),
      'rot'
    ) ?? 0
  ) / 60_000;

describe('E11: rotating a flipped shape', () => {
  const deck = (): Uint8Array =>
    buildPptx({
      slides: [
        {
          shapes:
            shape({ id: 2, name: 'flipV', x: IN, y: IN, cx: 2 * IN, cy: IN, flipV: true }) +
            shape({ id: 3, name: 'plain', x: 4 * IN, y: IN, cx: 2 * IN, cy: IN }) +
            shape({
              id: 4,
              name: 'turned',
              x: 7 * IN,
              y: IN,
              cx: 2 * IN,
              cy: IN,
              rot: 180 * 60_000,
            }) +
            shape({
              id: 5,
              name: 'flipV30',
              x: IN,
              y: 4 * IN,
              cx: 2 * IN,
              cy: IN,
              rot: 30 * 60_000,
              flipV: true,
            }),
        },
      ],
    });

  test('the handle is outside the box, above it in the shape frame', async () => {
    const e = await open(deck());
    for (const name of ['flipV', 'plain', 'turned', 'flipV30']) {
      const item = itemOf(e, name);
      const h = rotationHandle(item, 22 * PX);
      const c = slideCenter(item);
      const r = (item.element.xfrm.rot * Math.PI) / 180;
      // Distance from the center along the shape's own vertical axis.
      const along = Math.abs(-(h.x - c.x) * Math.sin(r) + (h.y - c.y) * Math.cos(r));
      expect(along).toBeGreaterThan(item.element.xfrm.cy / 2);
    }
  });

  test('the first move turns a flipV shape by the pointer angle change, not to 180°', async () => {
    const e = await open(deck());
    const item = itemOf(e, 'flipV');
    const h = rotationHandle(item, 22 * PX);
    const rot = rotationAfterDrag(item, boxOf(item), h, { x: h.x + 2 * PX, y: h.y }, false);
    expect(Math.min(rot, 360 - rot)).toBeLessThan(2);
    // Through the editor: the shape keeps its flip and turns by the small angle only.
    e.setSelection({ kind: 'shapes', nodeIds: [idByName(e, 'flipV')] });
    expect(
      e.exec({ type: 'setTransform', nodeId: item.element.nodeId, ...boxOf(item), rot }).ok
    ).toBe(true);
    expect(Math.min(rotOf(e, 'flipV'), 360 - rotOf(e, 'flipV'))).toBeLessThan(2);
  });

  test('a quarter turn of the pointer adds 90° to a rotated flipV shape', async () => {
    const e = await open(deck());
    const item = itemOf(e, 'flipV30');
    const c = slideCenter(item);
    const from = { x: c.x + IN, y: c.y };
    const to = { x: c.x, y: c.y + IN };
    expect(rotationAfterDrag(item, boxOf(item), from, to, false)).toBeCloseTo(120, 6);
    // Shift snaps the result to 15° steps.
    const near = { x: c.x + IN * Math.cos(1.6), y: c.y + IN * Math.sin(1.6) };
    expect(rotationAfterDrag(item, boxOf(item), from, near, true)).toBe(120);
  });
});

describe('E13: z-order with several shapes selected', () => {
  const ext = '<p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"/></p:extLst>';
  const deck = (): Uint8Array =>
    buildPptx({
      slides: [
        {
          shapes:
            ['A', 'B', 'C', 'D']
              .map((n, i) => shape({ id: i + 2, name: n, x: (i + 1) * IN, y: IN, cx: IN, cy: IN }))
              .join('') + ext,
        },
      ],
    });
  const order = async (names: string[], move: 'front' | 'back' | 'forward' | 'backward') => {
    const e = await open(deck());
    e.setSelection({ kind: 'shapes', nodeIds: names.map((n) => idByName(e, n)) });
    expect(e.exec({ type: 'order', move }).ok).toBe(true);
    return { e, stack: stack(e) };
  };

  test('Bring Forward moves adjacent selected shapes together', async () => {
    expect((await order(['A', 'B'], 'forward')).stack).toEqual(['C', 'A', 'B', 'D', 'extLst']);
    expect((await order(['A', 'C'], 'forward')).stack).toEqual(['B', 'A', 'D', 'C', 'extLst']);
    // The top shape cannot move; the one under it stays under it.
    expect((await order(['C', 'D'], 'forward')).stack).toEqual(['A', 'B', 'C', 'D', 'extLst']);
  });

  test('Send Backward moves adjacent selected shapes together', async () => {
    expect((await order(['C', 'D'], 'backward')).stack).toEqual(['A', 'C', 'D', 'B', 'extLst']);
    expect((await order(['B', 'D'], 'backward')).stack).toEqual(['B', 'A', 'D', 'C', 'extLst']);
  });

  test('Bring to Front and Send to Back keep the stacking order, not the click order', async () => {
    expect((await order(['B', 'A'], 'front')).stack).toEqual(['C', 'D', 'A', 'B', 'extLst']);
    expect((await order(['D', 'C'], 'back')).stack).toEqual(['C', 'D', 'A', 'B', 'extLst']);
    expect((await order(['D', 'A'], 'front')).stack).toEqual(['B', 'C', 'A', 'D', 'extLst']);
  });

  test('one undo step restores the order', async () => {
    const { e } = await order(['D', 'A'], 'back');
    expect(stack(e)).toEqual(['A', 'D', 'B', 'C', 'extLst']);
    e.exec({ type: 'undo' });
    expect(stack(e)).toEqual(['A', 'B', 'C', 'D', 'extLst']);
  });
});

describe('E16: transforms out of the ST_Coordinate range', () => {
  const deck = (): Uint8Array =>
    buildPptx({ slides: [{ shapes: shape({ id: 2, name: 'S', x: IN, y: IN, cx: IN, cy: IN }) }] });
  const transform = (
    e: PptxEditorInstance,
    box: Partial<Record<'x' | 'y' | 'cx' | 'cy', number>>
  ) =>
    e.exec({
      type: 'setTransform',
      nodeId: idByName(e, 'S'),
      x: IN,
      y: IN,
      cx: IN,
      cy: IN,
      rot: 0,
      ...box,
    });
  /** Every attribute value of the shape's `a:off` and `a:ext`. */
  const values = (e: PptxEditorInstance): string =>
    ['a:off', 'a:ext']
      .flatMap(
        (n) =>
          find(
            elements(tree(e)).find((c) => nameOf(c) === 'S'),
            n
          )!.attrs
      )
      .map((a) => a.value)
      .join(' ');

  test('values out of the range are refused with a typed reason', async () => {
    const e = await open(deck());
    for (const box of [
      { x: 1e21 },
      { y: -1e21 },
      { x: MAX_COORDINATE + 1 },
      { y: MIN_COORDINATE - 1 },
      { cx: MAX_COORDINATE + 1 },
      { cy: 1e22 },
    ]) {
      const r = transform(e, box);
      expect(r).toMatchObject({ ok: false, reason: 'errors.coordinateRange' });
      expect(e.snapshot().lastRejection?.reason).toBe('errors.coordinateRange');
    }
    expect(values(e)).toBe('914400 914400 914400 914400');
    expect(e.snapshot().canUndo).toBe(false);
  });

  test('the ends of the range are written as plain integers', async () => {
    const e = await open(deck());
    expect(transform(e, { x: MAX_COORDINATE, y: MIN_COORDINATE, cx: MAX_COORDINATE }).ok).toBe(
      true
    );
    const off = find(
      elements(tree(e)).find((c) => nameOf(c) === 'S'),
      'a:off'
    );
    expect(attr(off, 'x')).toBe('27273042316900');
    expect(attr(off, 'y')).toBe('-27273042329600');
  });

  test('a duplicate that would leave the range is refused', async () => {
    const e = await open(deck());
    expect(transform(e, { x: MAX_COORDINATE - 100 }).ok).toBe(true);
    e.setSelection({ kind: 'shapes', nodeIds: [idByName(e, 'S')] });
    expect(e.exec({ type: 'duplicate' })).toMatchObject({
      ok: false,
      reason: 'errors.coordinateRange',
    });
  });

  test('a rotation that is not a number is refused', async () => {
    const e = await open(deck());
    expect(transform(e, {}).ok).toBe(true);
    const r = e.exec({
      type: 'setTransform',
      nodeId: idByName(e, 'S'),
      x: IN,
      y: IN,
      cx: IN,
      cy: IN,
      rot: Number.NaN,
    });
    expect(r.ok).toBe(false);
  });
});

describe('E17: the font list follows the open document', () => {
  const deck = (family: string): Uint8Array =>
    buildPptx({
      slides: [
        {
          shapes: shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 4 * IN,
            cy: IN,
            text: para(run('Text', '', `<a:latin typeface="${family}"/>`)),
          }),
        },
      ],
    });

  test('opening another deck at the same revision lists its fonts, not the old ones', async () => {
    const e = await open(deck('Alpha Sans'));
    expect(e.fontFamilies()).toContain('Alpha Sans');
    expect((await e.load(deck('Beta Serif'))).ok).toBe(true);
    expect(e.session!.store.revision).toBe(0);
    expect(e.fontFamilies()).toContain('Beta Serif');
    expect(e.fontFamilies()).not.toContain('Alpha Sans');
  });
});

describe('refusals that only say there was nothing to do', () => {
  const deck = (): Uint8Array =>
    buildPptx({ slides: [{ shapes: shape({ id: 2, name: 'S', x: IN, y: IN, cx: IN, cy: IN }) }] });

  test('undo and redo with no history carry noop', async () => {
    const e = await open(deck());
    expect(e.exec({ type: 'undo' })).toEqual({
      ok: false,
      reason: 'disabled.nothingToUndo',
      noop: true,
    });
    expect(e.snapshot().lastRejection).toEqual({ reason: 'disabled.nothingToUndo', noop: true });
    expect(e.exec({ type: 'redo' })).toMatchObject({
      reason: 'disabled.nothingToRedo',
      noop: true,
    });
    expect(e.snapshot().lastRejection?.noop).toBe(true);
  });

  test('a command with nothing selected carries noop', async () => {
    const e = await open(deck());
    expect(e.exec({ type: 'toggleBold' })).toMatchObject({
      reason: 'disabled.noSelection',
      noop: true,
    });
  });

  test('a real refusal has no noop flag', async () => {
    const e = await open(deck());
    const r = e.exec({ type: 'deleteSlide', index: 0 });
    expect(r).toEqual({ ok: false, reason: 'disabled.lastSlide' });
    expect(e.snapshot().lastRejection).toEqual({ reason: 'disabled.lastSlide' });
    e.setMode('view');
    expect(e.exec({ type: 'undo' })).toEqual({ ok: false, reason: 'disabled.viewing' });
  });
});

describe('focus', () => {
  test('focus() before a surface is attached, or after destroy, does nothing', async () => {
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect(() => e.focus()).not.toThrow();
    expect((await e.load(buildPptx({ slides: [{ shapes: '' }] }))).ok).toBe(true);
    expect(() => e.focus()).not.toThrow();
    e.destroy();
    expect(() => e.focus()).not.toThrow();
  });
});
