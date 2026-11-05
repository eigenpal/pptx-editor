import { describe, expect, test } from 'bun:test';
import { buildPptx, shape, solid } from '../../../../../scripts/fixtures/builder.ts';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { find, readPptx, serializeXmlPart, xmlPartOf } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { distributeBoxes, flipped, rotated } from '../../automation/commands/arrange-commands.ts';

const IN = 914_400;

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

const GROUP =
  '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group 9"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="914400" y="4572000"/><a:ext cx="1828800" cy="914400"/><a:chOff x="914400" y="4572000"/><a:chExt cx="1828800" cy="914400"/></a:xfrm></p:grpSpPr>' +
  shape({ id: 11, name: 'Inner A', x: IN, y: 5 * IN, cx: IN, cy: IN, fill: solid('70AD47') }) +
  shape({ id: 12, name: 'Inner B', x: 2 * IN, y: 5 * IN, cx: IN, cy: IN, fill: solid('FFC000') }) +
  '</p:grpSp>';

const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          shape({ id: 2, name: 'Left', x: IN, y: IN, cx: IN, cy: IN, fill: solid('4472C4') }) +
          shape({
            id: 3,
            name: 'Middle',
            x: 2.5 * IN,
            y: IN,
            cx: 2 * IN,
            cy: IN,
            fill: solid('ED7D31'),
          }) +
          shape({
            id: 4,
            name: 'Right',
            x: 9 * IN,
            y: 2 * IN,
            cx: IN,
            cy: 2 * IN,
            rot: 30 * 60_000,
            fill: solid('A5A5A5'),
          }) +
          GROUP,
      },
    ],
  });

const slideXml = (e: PptxEditorInstance): string =>
  serializeXmlPart(xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!);
const top = (e: PptxEditorInstance): string[] =>
  e.session!.layout(0)!.items.map((i) => i.element.nodeId);
const xfrmOf = (e: PptxEditorInstance, id: string) =>
  e.session!.layout(0)!.items.find((i) => i.element.nodeId === id)!.element.xfrm;

describe('distribute', () => {
  test('equal gaps between three or more boxes; the outer boxes stay', () => {
    const out = distributeBoxes(
      [
        { id: 'a', x: 0, y: 0, cx: 10, cy: 10 },
        { id: 'c', x: 90, y: 5, cx: 10, cy: 10 },
        { id: 'b', x: 20, y: 9, cx: 20, cy: 10 },
      ],
      'horizontal',
      1000
    );
    expect(out).toEqual([
      { id: 'a', x: 0, y: 0 },
      { id: 'c', x: 90, y: 5 },
      { id: 'b', x: 40, y: 9 },
    ]);
    expect(distributeBoxes([{ id: 'a', x: 0, y: 3, cx: 10, cy: 20 }], 'vertical', 100)).toEqual([
      { id: 'a', x: 0, y: 40 },
    ]);
    expect(() =>
      distributeBoxes(
        [
          { id: 'a', x: 0, y: 0, cx: 1, cy: 1 },
          { id: 'b', x: 5, y: 0, cx: 1, cy: 1 },
        ],
        'horizontal',
        100
      )
    ).toThrow();
  });

  test('the command moves shapes as one undo step; two shapes are refused', async () => {
    const e = await open(deck());
    const [left, middle, right] = top(e);
    e.setSelection({ kind: 'shapes', nodeIds: [left, right] });
    expect(e.can({ type: 'distributeObjects', axis: 'horizontal' })).toEqual({
      ok: false,
      reason: 'disabled.distributeNeedsThree',
    });
    e.setSelection({ kind: 'shapes', nodeIds: [right, middle, left] });
    expect(e.exec({ type: 'distributeObjects', axis: 'horizontal' }).ok).toBe(true);
    // Left 1–2 in, right 9–10 in, middle 2 in wide: gaps of 2.5 in.
    expect(xfrmOf(e, middle).x).toBe(Math.round(4.5 * IN));
    expect(xfrmOf(e, left).x).toBe(IN);
    expect(xfrmOf(e, right).x).toBe(9 * IN);
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(xfrmOf(e, middle).x).toBe(Math.round(2.5 * IN));
    e.setSelection({ kind: 'shapes', nodeIds: [left] });
    expect(e.exec({ type: 'distributeObjects', axis: 'vertical' }).ok).toBe(true);
    expect(xfrmOf(e, left).y).toBe((e.snapshot().slideHeight - IN) / 2);
  });
});

describe('flip and rotate', () => {
  test('a flip mirrors the element: the flip toggles and a turn reverses', async () => {
    expect(flipped({ rot: 30, flipH: false, flipV: false }, 'horizontal')).toEqual({
      rot: 330,
      flipH: true,
      flipV: false,
    });
    expect(flipped({ rot: 0, flipH: true, flipV: false }, 'vertical')).toEqual({
      rot: 0,
      flipH: true,
      flipV: true,
    });
    expect(rotated(300, 90)).toBe(30);
    expect(rotated(10, -90)).toBe(280);
    expect(() => rotated(0, Number.POSITIVE_INFINITY)).toThrow();

    const e = await open(deck());
    const [, , right] = top(e);
    e.setSelection({ kind: 'shapes', nodeIds: [right] });
    expect(e.exec({ type: 'flipObjects', axis: 'horizontal' }).ok).toBe(true);
    expect(slideXml(e)).toContain('<a:xfrm rot="19800000" flipH="1">');
    const g = e.snapshot().geometry!;
    expect([g.rot, g.flipH, g.flipV]).toEqual([330, true, false]);
    expect(e.exec({ type: 'flipObjects', axis: 'horizontal' }).ok).toBe(true);
    expect(e.snapshot().geometry!.flipH).toBe(false);
    expect(e.snapshot().geometry!.rot).toBe(30);
    expect(e.exec({ type: 'rotateObjects', degrees: 90 }).ok).toBe(true);
    expect(e.snapshot().geometry!.rot).toBe(120);
    expect(e.exec({ type: 'rotateObjects', degrees: -90 }).ok).toBe(true);
    expect(e.snapshot().geometry!.rot).toBe(30);
    expect(readPptx(e.save()!).ok).toBe(true);
  });

  test('a group flips as one element, about its own center', async () => {
    const e = await open(deck());
    const group = top(e)[3];
    e.setSelection({ kind: 'shapes', nodeIds: [group] });
    const before = xfrmOf(e, group);
    expect(e.exec({ type: 'flipObjects', axis: 'vertical' }).ok).toBe(true);
    const after = xfrmOf(e, group);
    expect([after.x, after.y, after.cx, after.cy, after.flipV]).toEqual([
      before.x,
      before.y,
      before.cx,
      before.cy,
      true,
    ]);
  });
});

describe('lock aspect ratio', () => {
  test('writes the locks as the first child and shows in the snapshot', async () => {
    const e = await open(deck());
    const [left] = top(e);
    e.setSelection({ kind: 'shapes', nodeIds: [left] });
    expect(e.snapshot().geometry!.lockAspectRatio).toBe(false);
    expect(e.exec({ type: 'setLockAspectRatio', lock: true }).ok).toBe(true);
    expect(slideXml(e)).toContain('<p:cNvSpPr><a:spLocks noChangeAspect="1"/></p:cNvSpPr>');
    expect(e.snapshot().geometry!.lockAspectRatio).toBe(true);
    // The same setting again changes nothing and adds no undo step.
    expect(e.exec({ type: 'setLockAspectRatio', lock: true })).toMatchObject({
      ok: false,
      reason: 'disabled.noLockChange',
    });
    expect(e.exec({ type: 'setLockAspectRatio', lock: false }).ok).toBe(true);
    expect(e.snapshot().geometry!.lockAspectRatio).toBe(false);
    expect(slideXml(e)).not.toContain('noChangeAspect');
    const group = top(e)[3];
    e.setSelection({ kind: 'shapes', nodeIds: [group] });
    expect(e.exec({ type: 'setLockAspectRatio', lock: true }).ok).toBe(true);
    expect(slideXml(e)).toContain(
      '<p:cNvGrpSpPr><a:grpSpLocks noChangeAspect="1"/></p:cNvGrpSpPr>'
    );
  });
});

describe('selection pane', () => {
  test('lists the elements top first, with group children', async () => {
    const e = await open(deck());
    const list = e.snapshot().shapes!;
    expect(list.map((s) => s.name)).toEqual(['Group 9', 'Right', 'Middle', 'Left']);
    expect(list[0].kind).toBe('group');
    expect(list[0].children!.map((s) => s.name)).toEqual(['Inner B', 'Inner A']);
    // Reference-stable while the slide does not change.
    e.setSelection({ kind: 'shapes', nodeIds: [list[1].nodeId] });
    expect(e.snapshot().shapes).toBe(list);
  });

  test('hide and show: a hidden element leaves the layout and the selection', async () => {
    const e = await open(deck());
    const [left, middle] = top(e);
    e.setSelection({ kind: 'shapes', nodeIds: [left, middle] });
    expect(e.exec({ type: 'setShapesVisible', nodeIds: [left], visible: false }).ok).toBe(true);
    expect(slideXml(e)).toContain('<p:cNvPr id="2" name="Left" hidden="1"/>');
    expect(top(e)).not.toContain(left);
    expect(e.snapshot().selection).toEqual({ kind: 'shapes', nodeIds: [middle] });
    expect(e.snapshot().shapes!.find((s) => s.nodeId === left)!.hidden).toBe(true);
    const inner = e.snapshot().shapes![0].children![0].nodeId;
    expect(e.exec({ type: 'setShapesVisible', nodeIds: [inner], visible: false }).ok).toBe(true);
    expect(e.snapshot().shapes![0].children![0].hidden).toBe(true);
    expect(e.exec({ type: 'setShapesVisible', nodeIds: [left, inner], visible: true }).ok).toBe(
      true
    );
    expect(top(e)).toContain(left);
    expect(slideXml(e)).not.toContain('hidden=');
    expect(e.exec({ type: 'setShapesVisible', nodeIds: [], visible: true }).ok).toBe(false);
  });

  test('rename writes the name; empty and control-character names are refused', async () => {
    const e = await open(deck());
    const [left] = top(e);
    expect(e.exec({ type: 'renameShape', nodeId: left, name: '  Logo & <mark>  ' }).ok).toBe(true);
    expect(slideXml(e)).toContain('name="Logo &amp; &lt;mark&gt;"');
    expect(e.snapshot().shapes!.at(-1)!.name).toBe('Logo & <mark>');
    expect(e.exec({ type: 'renameShape', nodeId: left, name: 'Logo & <mark>' })).toMatchObject({
      ok: false,
      reason: 'disabled.noNameChange',
    });
    expect(e.exec({ type: 'renameShape', nodeId: left, name: '   ' })).toMatchObject({
      ok: false,
      reason: 'errors.invalidName',
    });
    expect(e.exec({ type: 'renameShape', nodeId: left, name: 'a\u0007b' }).ok).toBe(false);
    expect(e.exec({ type: 'renameShape', nodeId: left, name: 'x'.repeat(256) }).ok).toBe(false);
    expect(readPptx(e.save()!).ok).toBe(true);
  });
});

describe('round 1 review fixes', () => {
  const groupDeck = (xfrm: string): Uint8Array =>
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes:
            '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
            `<p:grpSpPr><a:xfrm ${xfrm}><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/><a:chOff x="914400" y="914400"/><a:chExt cx="1828800" cy="914400"/></a:xfrm></p:grpSpPr>` +
            shape({
              id: 11,
              name: 'C',
              x: IN,
              y: IN,
              cx: IN,
              cy: 0.5 * IN,
              prst: 'rightArrow',
              rot: 20 * 60_000,
              fill: solid('70AD47'),
            }) +
            shape({
              id: 12,
              name: 'D',
              x: 2 * IN,
              y: IN,
              cx: 0.5 * IN,
              cy: 0.5 * IN,
              fill: solid('FFC000'),
            }) +
            '</p:grpSp>',
        },
      ],
    });
  const linear = (e: PptxEditorInstance, id: string): number[] => {
    const find = (
      items: readonly {
        element: { nodeId: string };
        matrix: readonly number[];
        children?: unknown;
      }[]
    ): readonly number[] | undefined => {
      for (const it of items) {
        if (it.element.nodeId === id) return it.matrix;
        const kids = (it as { children?: typeof items }).children;
        const inner = kids ? find(kids) : undefined;
        if (inner) return inner;
      }
      return undefined;
    };
    const m = find(e.session!.layout(0)!.items as never)!;
    return m.slice(0, 4).map((v) => Math.round(v * 1e6) / 1e6);
  };
  const child = (e: PptxEditorInstance): { group: string; c: string } => {
    const g = e.snapshot().shapes![0]!;
    return { group: g.nodeId, c: g.children!.find((s) => s.name === 'C')!.nodeId };
  };

  test('flip in a turned group mirrors on the slide', async () => {
    const e = await open(groupDeck('rot="5400000"'));
    const { group, c } = child(e);
    e.setSelection({ kind: 'shapes', nodeIds: [c], group });
    const [a, b, cc, d] = linear(e, c);
    expect(e.exec({ type: 'flipObjects', axis: 'horizontal' }).ok).toBe(true);
    // A horizontal mirror on the slide negates the x row of the linear part.
    expect(linear(e, c)).toEqual([-a, b, -cc, d].map((v) => Math.round(v * 1e6) / 1e6 + 0));
  });

  test('rotate right in a mirrored group turns clockwise on the slide', async () => {
    const e = await open(groupDeck('flipH="1"'));
    const { group, c } = child(e);
    e.setSelection({ kind: 'shapes', nodeIds: [c], group });
    const [a, b, cc, d] = linear(e, c);
    expect(e.exec({ type: 'rotateObjects', degrees: 90 }).ok).toBe(true);
    // R(90°) · M with y down: (x, y) → (−y, x).
    expect(linear(e, c)).toEqual([-b, a, -d, cc].map((v) => Math.round(v * 1e6) / 1e6 + 0));
  });

  test('hiding a group clears the selection of its child; show all with nothing hidden is a no-op', async () => {
    const e = await open(groupDeck(''));
    const { group, c } = child(e);
    e.setSelection({ kind: 'shapes', nodeIds: [c], group });
    expect(e.exec({ type: 'setShapesVisible', nodeIds: [group], visible: false }).ok).toBe(true);
    expect(e.snapshot().selection).toEqual({ kind: 'none' });
    expect(e.exec({ type: 'setShapesVisible', nodeIds: [group], visible: true }).ok).toBe(true);
    expect(e.exec({ type: 'setShapesVisible', nodeIds: [group, c], visible: true })).toMatchObject({
      ok: false,
      reason: 'disabled.noVisibilityChange',
      noop: true,
    });
  });

  test('names with noncharacters or lone surrogates are refused as invalid names', async () => {
    const e = await open(groupDeck(''));
    const { c } = child(e);
    for (const name of ['a￾b', 'a\uD800b', 'a\uDC00'])
      expect(e.exec({ type: 'renameShape', nodeId: c, name })).toMatchObject({
        ok: false,
        reason: 'errors.invalidName',
      });
    expect(e.exec({ type: 'renameShape', nodeId: c, name: 'emoji 😀' }).ok).toBe(true);
  });

  test('the pane lists the drawn mc:Choice copy; hide, rename, and the lock reach the Fallback too', async () => {
    const bytes = new Uint8Array(
      readFileSync(join(import.meta.dir, '../../../../../fixtures/preserve-unsupported.pptx'))
    );
    const e = await open(bytes);
    const drawn = e.snapshot().shapes!.find((s) => s.kind === 'shape')!;
    // The `Requires="p14"` Choice is the branch the slide draws (ECMA-376 Part 3, 10.2.1).
    const part = xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!;
    const choiceSp = find(find(part.root, 'mc:Choice'), 'p:sp')!;
    expect(drawn.nodeId).toBe(choiceSp.id);
    expect(top(e)).toContain(drawn.nodeId);
    e.setSelection({ kind: 'shapes', nodeIds: [drawn.nodeId] });
    expect(e.exec({ type: 'renameShape', nodeId: drawn.nodeId, name: 'Both' }).ok).toBe(true);
    expect(e.exec({ type: 'setLockAspectRatio', lock: true }).ok).toBe(true);
    expect(e.exec({ type: 'setShapesVisible', nodeIds: [drawn.nodeId], visible: false }).ok).toBe(
      true
    );
    const xml = slideXml(e);
    expect(xml.match(/<p:cNvPr id="3" name="Both" hidden="1"\/>/g)).toHaveLength(2);
    expect(xml.match(/<a:spLocks noChangeAspect="1"\/>/g)).toHaveLength(2);
  });
});

describe('arrange commands that change nothing', () => {
  test('a turn of 0° or 360° and a second Distribute make no undo step', async () => {
    const e = await open(deck());
    const [a, b, c] = top(e);
    e.setSelection({ kind: 'shapes', nodeIds: [a!, b!, c!] });
    expect(e.exec({ type: 'distributeObjects', axis: 'horizontal' }).ok).toBe(true);
    const xml = slideXml(e);
    e.exec({ type: 'undo' });
    e.exec({ type: 'redo' });
    for (const command of [
      { type: 'distributeObjects', axis: 'horizontal' },
      { type: 'rotateObjects', degrees: 0 },
      { type: 'rotateObjects', degrees: 360 },
    ] as const)
      expect(e.exec(command)).toMatchObject({
        ok: false,
        reason: 'disabled.noArrangeChange',
        noop: true,
      });
    expect(slideXml(e)).toBe(xml);
    // One undo step back to the deck as loaded: the refusals added none.
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(e.snapshot().canUndo).toBe(false);
  });
});
