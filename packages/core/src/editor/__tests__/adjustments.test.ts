// Adjust values of preset shapes (`setAdjustments`): the write, the handle range, the
// refusals, undo, and save and reopen. The handles that drag them are in surface-adjust.test.ts.

import { describe, expect, test } from 'bun:test';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import {
  attr,
  child,
  elements,
  find,
  presentationInfo,
  readPptx,
  shapeTreeOf,
  type PptxPackage,
} from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

const shape = (id: number, name: string, prst: string, av = ''): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="1371600"/></a:xfrm>` +
  `<a:prstGeom prst="${prst}"><a:avLst>${av}</a:avLst></a:prstGeom></p:spPr></p:sp>`;

async function open(): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const deck = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          shape(2, 'Round', 'roundRect') +
          shape(3, 'Arrow', 'rightArrow', '<a:gd name="adj1" fmla="val 30000"/>') +
          shape(4, 'Box', 'rect'),
      },
    ],
  });
  expect((await e.load(deck)).ok).toBe(true);
  return e;
}

function select(e: PptxEditorInstance, name: string): void {
  const id = e.snapshot().shapes!.find((s) => s.name === name)!.nodeId;
  e.setSelection({ kind: 'shapes', nodeIds: [id] });
}

/** The `a:gd` values of shape `name`, by guide name. */
function guides(pkg: PptxPackage, name: string): Record<string, string> {
  const tree = shapeTreeOf(pkg, presentationInfo(pkg).slides[0]!.partName);
  const sp = elements(tree).find((el) => attr(find(el, 'p:cNvPr'), 'name') === name)!;
  const out: Record<string, string> = {};
  for (const gd of elements(child(find(sp, 'a:prstGeom'), 'a:avLst')))
    out[attr(gd, 'name')!] = attr(gd, 'fmla')!;
  return out;
}

const pkgOf = (e: PptxEditorInstance): PptxPackage => e.session!.store.package;
const reopen = (e: PptxEditorInstance): PptxPackage => {
  const r = readPptx(e.save()!);
  if (!r.ok) throw new Error('reopen');
  return r.pkg;
};

describe('setAdjustments', () => {
  test('writes the value in one undo step and keeps it after save and reopen', async () => {
    const e = await open();
    select(e, 'Round');
    expect(e.exec({ type: 'setAdjustments', values: { adj: 40_000 } }).ok).toBe(true);
    expect(guides(reopen(e), 'Round')).toEqual({ adj: 'val 40000' });
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(guides(pkgOf(e), 'Round')).toEqual({});
  });

  test('bounds a value to its handle range and keeps the other values', async () => {
    const e = await open();
    select(e, 'Round');
    // roundRect: `adj` is pinned to 0..50000.
    expect(e.exec({ type: 'setAdjustments', values: { adj: 90_000 } }).ok).toBe(true);
    expect(guides(pkgOf(e), 'Round')).toEqual({ adj: 'val 50000' });
    select(e, 'Arrow');
    expect(e.exec({ type: 'setAdjustments', values: { adj2: 20_000 } }).ok).toBe(true);
    expect(guides(pkgOf(e), 'Arrow')).toEqual({ adj1: 'val 30000', adj2: 'val 20000' });
  });

  test('refuses a shape with no handles, an unknown name, and a value with no change', async () => {
    const e = await open();
    select(e, 'Box');
    const box = e.exec({ type: 'setAdjustments', values: { adj: 1 } });
    expect(box.ok ? '' : box.reason).toBe('disabled.noAdjustHandles');
    select(e, 'Round');
    const name = e.exec({ type: 'setAdjustments', values: { 'x" y="1': 1 } });
    expect(name.ok ? '' : name.reason).toBe('errors.unsupported');
    const same = e.exec({ type: 'setAdjustments', values: { adj: 16_667 } });
    expect(same.ok).toBe(false);
    if (!same.ok) expect([same.reason, same.noop]).toEqual(['disabled.noAdjustChange', true]);
    e.setSelection({ kind: 'none' });
    expect(e.can({ type: 'setAdjustments', values: { adj: 1 } }).ok).toBe(false);
  });
});
