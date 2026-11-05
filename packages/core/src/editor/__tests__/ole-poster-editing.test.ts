// An OLE frame that shows its `p:oleObj/p:pic` picture (ECMA-376 Part 1, 19.3.1.36) edits as
// one frame: move, resize, copy, delete, and undo act on the frame, and the picture follows its
// box. The picture is read-only: nothing writes to `p:oleObj`, so it saves as it was read.

import { describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { buildSlideStructureDeck } from '../../../../../scripts/fixtures/slide-structure.ts';
import type { LayoutItem } from '../../layout/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

const OLE_SLIDE = 3;

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  e.goToSlide(OLE_SLIDE);
  return e;
}

type Frame = Extract<LayoutItem, { kind: 'frame' }>;

function oleFrames(e: PptxEditorInstance): Frame[] {
  const items = e.session!.layout(OLE_SLIDE)!.items;
  return items.filter((i): i is Frame => i.kind === 'frame' && i.element.frameKind === 'ole');
}

function oleFrame(e: PptxEditorInstance): Frame {
  const [frame] = oleFrames(e);
  if (!frame?.poster) throw new Error('no OLE frame with a poster');
  return frame;
}

const boxOf = (f: Frame) => {
  const { x, y, cx, cy } = f.element.xfrm;
  return { x, y, cx, cy };
};

/** Every `p:oleObj` element of a saved slide, as written. */
function oleXml(bytes: Uint8Array): string[] {
  const xml = strFromU8(unzipSync(bytes)[`ppt/slides/slide${OLE_SLIDE + 1}.xml`]!);
  return xml.match(/<p:oleObj\b[\s\S]*?<\/p:oleObj>/g) ?? [];
}

describe('editing an OLE frame that shows its picture', () => {
  test('a move and a resize move the picture with the frame; undo puts both back', async () => {
    const e = await open(buildSlideStructureDeck());
    const before = oleFrame(e);
    const id = before.element.nodeId;
    e.setSelection({ kind: 'shapes', nodeIds: [id] });

    expect(e.exec({ type: 'nudge', dx: 914_400, dy: 0 }).ok).toBe(true);
    expect(boxOf(oleFrame(e)).x).toBe(before.element.xfrm.x + 914_400);
    expect(oleFrame(e).poster!.element.xfrm).toMatchObject(boxOf(oleFrame(e)));

    const resized = { nodeId: id, x: 0, y: 0, cx: 4 * 914_400, cy: 3 * 914_400, rot: 0 };
    expect(e.exec({ type: 'setTransforms', items: [resized] }).ok).toBe(true);
    expect(boxOf(oleFrame(e))).toEqual({ x: 0, y: 0, cx: resized.cx, cy: resized.cy });
    expect(oleFrame(e).poster!.element.xfrm).toMatchObject(boxOf(oleFrame(e)));
    expect(oleFrame(e).poster!.matrix).toEqual(oleFrame(e).matrix);

    e.exec({ type: 'undo' });
    e.exec({ type: 'undo' });
    expect(boxOf(oleFrame(e))).toEqual(boxOf(before));
    expect(oleFrame(e).poster!.element.xfrm).toMatchObject(boxOf(before));
  });

  test('duplicate gives a second frame with its own picture', async () => {
    const e = await open(buildSlideStructureDeck());
    e.setSelection({ kind: 'shapes', nodeIds: [oleFrame(e).element.nodeId] });
    expect(e.exec({ type: 'duplicate' }).ok).toBe(true);
    const frames = oleFrames(e);
    expect(frames.length).toBe(2);
    for (const f of frames) expect(f.poster?.element.nodeId).toBe(f.element.nodeId);
  });

  test('delete removes the frame and its picture; undo restores both', async () => {
    const e = await open(buildSlideStructureDeck());
    e.setSelection({ kind: 'shapes', nodeIds: [oleFrame(e).element.nodeId] });
    expect(e.exec({ type: 'delete' }).ok).toBe(true);
    expect(oleFrames(e)).toEqual([]);
    e.exec({ type: 'undo' });
    expect(oleFrame(e).poster!.element.fill.type).toBe('picture');
  });

  test('a moved frame saves its p:oleObj unchanged, and reopens with its picture', async () => {
    const e = await open(buildSlideStructureDeck());
    const untouched = oleXml(e.save()!);
    expect(untouched.length).toBe(1);
    e.setSelection({ kind: 'shapes', nodeIds: [oleFrame(e).element.nodeId] });
    e.exec({ type: 'nudge', dx: 914_400, dy: 914_400 });
    const moved = oleFrame(e);
    const saved = e.save()!;
    expect(oleXml(saved)).toEqual(untouched);

    const again = await open(saved);
    expect(boxOf(oleFrame(again))).toEqual(boxOf(moved));
    expect(oleFrame(again).poster!.element.xfrm).toMatchObject(boxOf(moved));
  });
});
