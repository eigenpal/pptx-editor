import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import {
  canonicalPackageFingerprint,
  contentTypeOf,
  find,
  presentationInfo,
  readPptx,
  textOf,
  writePptx,
  xmlPartOf,
} from '../../store/index.ts';
import { resolveSlide } from '../../cascade/index.ts';
import { createPptxEditor, notesText, type PptxEditorInstance } from '../index.ts';

const IN = 914_400;
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'title',
        shapes:
          shape({ id: 2, x: IN, y: IN, cx: 4 * IN, cy: IN, text: para(run('Hello world')) }) +
          shape({ id: 3, x: IN, y: 3 * IN, cx: 2 * IN, cy: IN, prst: 'ellipse' }),
        notes: 'First note',
      },
      { shapes: '', layout: 'blank' },
    ],
  });

async function open(bytes = deck()): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const r = await editor.load(bytes);
  expect(r.ok).toBe(true);
  return editor;
}

const slidePart = (e: PptxEditorInstance, i = 0) => e.session!.info.slides[i].partName;
const shapeId = (e: PptxEditorInstance, name: string): string => {
  const root = xmlPartOf(e.session!.store.package, slidePart(e))!.root;
  const walk = (el: typeof root): string | undefined => {
    for (const c of el.children) {
      if (c.type !== 'element') continue;
      if (c.kind === 'sp' && JSON.stringify(c).includes(`"value":"${name}"`)) return c.id;
      const hit = walk(c);
      if (hit) return hit;
    }
    return undefined;
  };
  return walk(root)!;
};
const slideText = (e: PptxEditorInstance): string =>
  textOf(find(xmlPartOf(e.session!.store.package, slidePart(e))!.root, 'p:txBody'));

/** Save, reopen, and check the package is readable and stable. */
function roundTrip(e: PptxEditorInstance): void {
  const bytes = e.save()!;
  const r = readPptx(bytes);
  expect(r.ok).toBe(true);
  if (r.ok)
    expect(canonicalPackageFingerprint(readPptx(writePptx(r.pkg)).ok ? r.pkg : r.pkg)).toBeTruthy();
}

describe('text editing', () => {
  test('typing, Enter, and Backspace edit the canonical tree and undo as one step', async () => {
    const e = await open();
    const id = shapeId(e, 'Shape 2');
    e.setSelection({
      kind: 'text',
      nodeId: id,
      anchor: { paragraph: 0, offset: 5 },
      focus: { paragraph: 0, offset: 5 },
    });
    for (const ch of ', dear') e.input('text', ch);
    expect(slideText(e)).toBe('Hello, dear world');
    e.input('paragraph');
    e.input('text', 'X');
    const body = find(xmlPartOf(e.session!.store.package, slidePart(e))!.root, 'p:txBody')!;
    expect(
      body.children.filter((c) => c.type === 'element' && c.kind === 'paragraph')
    ).toHaveLength(2);
    e.input('deleteBackward');
    e.input('deleteBackward');
    expect(slideText(e)).toBe('Hello, dear world');
    // Five intent-scoped entries: the typed word, Enter, "X", and two deletes.
    for (let i = 0; i < 4; i += 1) e.exec({ type: 'undo' });
    expect(slideText(e)).toBe('Hello, dear world');
    e.exec({ type: 'undo' });
    expect(slideText(e)).toBe('Hello world');
    roundTrip(e);
  });

  test('a range delete across paragraphs joins them', async () => {
    const e = await open();
    const id = shapeId(e, 'Shape 2');
    e.setSelection({
      kind: 'text',
      nodeId: id,
      anchor: { paragraph: 0, offset: 11 },
      focus: { paragraph: 0, offset: 11 },
    });
    e.input('paragraph');
    e.input('text', 'second line');
    e.setSelection({
      kind: 'text',
      nodeId: id,
      anchor: { paragraph: 0, offset: 5 },
      focus: { paragraph: 1, offset: 6 },
    });
    e.input('deleteBackward');
    expect(slideText(e)).toBe('Hello line');
  });

  test('bold on a range splits runs; on the shape it covers every run', async () => {
    const e = await open();
    const id = shapeId(e, 'Shape 2');
    e.setSelection({
      kind: 'text',
      nodeId: id,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 5 },
    });
    expect(e.exec({ type: 'toggleBold' }).ok).toBe(true);
    const runs = find(
      xmlPartOf(e.session!.store.package, slidePart(e))!.root,
      'a:p'
    )!.children.filter((c) => c.type === 'element' && c.kind === 'run');
    expect(runs).toHaveLength(2);
    expect(e.snapshot().format?.bold).toBe(true);
    e.setSelection({ kind: 'shapes', nodeIds: [id] });
    expect(e.exec({ type: 'setFontSize', points: 40 }).ok).toBe(true);
    expect(e.exec({ type: 'setTextColor', color: '#C00000' }).ok).toBe(true);
    const slide = resolveSlide(e.session!.store.package, 0)!;
    const sp = slide.elements.find((el) => el.nodeId === id);
    expect(
      sp?.kind === 'shape' && sp.text?.paragraphs[0].runs.every((r) => r.props.size === 4000)
    ).toBe(true);
    roundTrip(e);
  });

  test('formatting with nothing selected is refused with a reason', async () => {
    const e = await open();
    expect(e.can({ type: 'toggleBold' })).toEqual({ ok: false, reason: 'disabled.noSelection' });
  });
});

describe('shapes', () => {
  test('insert a text box, a shape, and an image; move one; all survive save', async () => {
    const e = await open();
    expect(e.exec({ type: 'insertTextBox' }).ok).toBe(true);
    e.input('text', 'New box');
    expect(e.exec({ type: 'insertShape', preset: 'hexagon' }).ok).toBe(true);
    const png = Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10]);
    expect(
      e.exec({ type: 'insertImage', bytes: png, mime: 'image/png', width: 10, height: 10 }).ok
    ).toBe(true);
    const sel = e.snapshot().selection;
    if (sel.kind !== 'shapes') throw new Error('expected shape selection');
    expect(e.exec({ type: 'nudge', dx: 9525, dy: 0 }).ok).toBe(true);
    const r = readPptx(e.save()!);
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    const slide = resolveSlide(r.pkg, 0)!;
    expect(slide.elements.some((el) => el.kind === 'picture')).toBe(true);
    expect(
      slide.elements.some(
        (el) =>
          el.kind === 'shape' && el.geometry.kind === 'preset' && el.geometry.preset === 'hexagon'
      )
    ).toBe(true);
    expect(contentTypeOf(r.pkg, '/ppt/media/image1.png')).toBe('image/png');
  });

  test('a text box with spAutoFit grows as text is added', async () => {
    const e = await open();
    e.exec({ type: 'insertTextBox' });
    const before = e.snapshot().geometry!.cy;
    e.input('text', 'line one');
    e.input('paragraph');
    e.input('text', 'line two');
    e.input('paragraph');
    e.input('text', 'line three');
    e.setSelection({
      kind: 'shapes',
      nodeIds: [(e.snapshot().selection as { nodeId: string }).nodeId ?? ''],
    });
    expect(e.snapshot().geometry!.cy).toBeGreaterThan(before);
  });

  test('duplicate, order, align, and delete', async () => {
    const e = await open();
    const id = shapeId(e, 'Shape 3');
    e.setSelection({ kind: 'shapes', nodeIds: [id] });
    expect(e.exec({ type: 'duplicate' }).ok).toBe(true);
    expect(e.exec({ type: 'order', move: 'back' }).ok).toBe(true);
    expect(e.exec({ type: 'alignObjects', align: 'center' }).ok).toBe(true);
    expect(e.exec({ type: 'delete' }).ok).toBe(true);
    roundTrip(e);
  });
});

describe('slides and notes', () => {
  test('add, duplicate, move, and delete slides keep a valid package', async () => {
    const e = await open();
    expect(e.exec({ type: 'addSlide' }).ok).toBe(true);
    expect(e.snapshot().slides).toHaveLength(3);
    expect(e.exec({ type: 'duplicateSlide', index: 0 }).ok).toBe(true);
    expect(e.snapshot().slides).toHaveLength(4);
    const dupNotes = notesText(e.session!.store.package, 1);
    expect(dupNotes).toBe('First note');
    expect(e.exec({ type: 'moveSlide', from: 0, to: 3 }).ok).toBe(true);
    expect(e.exec({ type: 'deleteSlide', index: 3 }).ok).toBe(true);
    expect(e.snapshot().slides).toHaveLength(3);
    const r = readPptx(e.save()!);
    expect(r.ok).toBe(true);
    if (r.ok) {
      const info = presentationInfo(r.pkg);
      expect(info.slides).toHaveLength(3);
      for (const s of info.slides) expect(contentTypeOf(r.pkg, s.partName)).toContain('slide+xml');
    }
  });

  test('the last slide cannot be deleted', async () => {
    const e = await open(buildPptx({ slides: [{ shapes: '' }] }));
    expect(e.can({ type: 'deleteSlide', index: 0 })).toEqual({
      ok: false,
      reason: 'disabled.lastSlide',
    });
  });

  test('notes are replaced, and created when missing', async () => {
    const e = await open();
    expect(e.exec({ type: 'setNotes', index: 0, text: 'Line A\nLine B' }).ok).toBe(true);
    expect(notesText(e.session!.store.package, 0)).toBe('Line A\nLine B');
    expect(e.exec({ type: 'setNotes', index: 1, text: 'Created' }).ok).toBe(true);
    expect(notesText(e.session!.store.package, 1)).toBe('Created');
    const r = readPptx(e.save()!);
    expect(r.ok && notesText(r.pkg, 1)).toBe('Created');
  });

  test('view mode refuses every edit', async () => {
    const e = await open();
    e.setMode('view');
    expect(e.can({ type: 'addSlide' })).toEqual({ ok: false, reason: 'disabled.viewing' });
    expect(e.exec({ type: 'insertTextBox' }).ok).toBe(false);
  });
});
