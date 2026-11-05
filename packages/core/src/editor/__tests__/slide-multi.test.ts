// Multi-slide selection in the rail (Shift+click, Ctrl/Cmd+click) and the commands that act on
// the whole selection, and Insert > Slide numbers.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { find, presentationInfo, textOf, xmlPartOf } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { deleteSlide } from '../../automation/commands/slide-commands.ts';

const IN = 914_400;

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

/** Five slides whose only shape says S0 … S4. */
const fiveSlides = (): Uint8Array =>
  buildPptx({
    slides: ['S0', 'S1', 'S2', 'S3', 'S4'].map((n) => ({
      shapes: shape({ id: 2, name: n, x: IN, y: IN, cx: 4 * IN, cy: IN, text: para(run(n)) }),
    })),
  });

/** The slide texts in deck order. */
const order = (e: PptxEditorInstance): string[] => {
  const pkg = e.session!.store.package;
  return presentationInfo(pkg).slides.map((s) =>
    textOf(find(xmlPartOf(pkg, s.partName)!.root, 'p:txBody'))
  );
};

describe('slide selection', () => {
  test('click, Ctrl/Cmd+click, and Shift+click select slides', async () => {
    const e = await open(fiveSlides());
    expect(e.snapshot().selectedSlides).toEqual([0]);
    e.selectSlide(2, 'toggle');
    expect(e.snapshot().selectedSlides).toEqual([0, 2]);
    expect(e.snapshot().currentSlide).toBe(2);
    e.selectSlide(4, 'range');
    // The range runs from the anchor (the last Ctrl/Cmd+clicked slide) to the clicked slide.
    expect(e.snapshot().selectedSlides).toEqual([2, 3, 4]);
    e.selectSlide(3, 'toggle');
    expect(e.snapshot().selectedSlides).toEqual([2, 4]);
    e.selectSlide(1);
    expect(e.snapshot().selectedSlides).toEqual([1]);
    // The last selected slide stays selected.
    e.selectSlide(1, 'toggle');
    expect(e.snapshot().selectedSlides).toEqual([1]);
  });

  test('the selection follows its slides when they move', async () => {
    const e = await open(fiveSlides());
    e.selectSlide(1);
    e.selectSlide(2, 'range');
    expect(e.exec({ type: 'moveSlides', indices: [1, 2], to: 3 }).ok).toBe(true);
    expect(order(e)).toEqual(['S0', 'S3', 'S4', 'S1', 'S2']);
    expect(e.snapshot().selectedSlides).toEqual([3, 4]);
  });
});

describe('commands on several slides', () => {
  test('delete removes every selected slide in one undo step', async () => {
    const e = await open(fiveSlides());
    expect(e.exec({ type: 'deleteSlides', indices: [1, 3] }).ok).toBe(true);
    expect(order(e)).toEqual(['S0', 'S2', 'S4']);
    expect(e.snapshot().selectedSlides).toEqual([1]);
    e.exec({ type: 'undo' });
    expect(order(e)).toEqual(['S0', 'S1', 'S2', 'S3', 'S4']);
  });

  test('delete refuses to remove every slide', async () => {
    const e = await open(fiveSlides());
    const r = e.exec({ type: 'deleteSlides', indices: [0, 1, 2, 3, 4] });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.reason).toBe('disabled.lastSlide');
    expect(order(e)).toHaveLength(5);
  });

  test('duplicate puts the copies after the last selected slide, in order, and selects them', async () => {
    const e = await open(fiveSlides());
    expect(e.exec({ type: 'duplicateSlides', indices: [0, 2] }).ok).toBe(true);
    expect(order(e)).toEqual(['S0', 'S1', 'S2', 'S0', 'S2', 'S3', 'S4']);
    expect(e.snapshot().selectedSlides).toEqual([3, 4]);
    e.exec({ type: 'undo' });
    expect(order(e)).toHaveLength(5);
  });

  test('move takes a scattered selection as one block', async () => {
    const e = await open(fiveSlides());
    expect(e.exec({ type: 'moveSlides', indices: [0, 3], to: 2 }).ok).toBe(true);
    expect(order(e)).toEqual(['S1', 'S2', 'S0', 'S3', 'S4']);
    const saved = e.save()!;
    const again = await open(saved);
    expect(order(again)).toEqual(['S1', 'S2', 'S0', 'S3', 'S4']);
  });
});

describe('Insert > Slide numbers', () => {
  const sample = (): Uint8Array =>
    new Uint8Array(
      readFileSync(join(import.meta.dir, '../../../../../examples/vite/public/sample.pptx'))
    );

  test('numbers every slide through the layout placeholder, and removes them again', async () => {
    const e = await open(sample());
    expect(e.snapshot().slideNumbers).toBe(false);
    expect(e.exec({ type: 'setSlideNumbers', show: true }).ok).toBe(true);
    expect(e.snapshot().slideNumbers).toBe(true);
    const count = e.session!.info.slides.length;
    for (let i = 0; i < count; i += 1) {
      const items = e.session!.layout(i)!.items;
      const painted = items.flatMap((item) =>
        item.kind === 'shape' && item.text
          ? item.text.lines.flatMap((l) => l.runs.map((r) => r.text))
          : []
      );
      expect(painted).toContain(String(i + 1));
    }
    expect(e.exec({ type: 'setSlideNumbers', show: false }).ok).toBe(true);
    expect(e.snapshot().slideNumbers).toBe(false);
    e.exec({ type: 'undo' });
    expect(e.snapshot().slideNumbers).toBe(true);
  });

  test('a deck without slide-number placeholders refuses with a reason', async () => {
    const e = await open(fiveSlides());
    const r = e.exec({ type: 'setSlideNumbers', show: true });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.reason).toBe('disabled.noSlideNumberPlaceholder');
  });
});

describe('review fixes', () => {
  test('moveSlides refuses a target that is not an integer', async () => {
    const e = await open(fiveSlides());
    for (const to of [Number.NaN, 1.5]) {
      const r = e.exec({ type: 'moveSlides', indices: [2], to });
      expect(r.ok).toBe(false);
    }
    expect(e.snapshot().currentSlide).toBe(0);
    expect(order(e)).toEqual(['S0', 'S1', 'S2', 'S3', 'S4']);
  });

  test('a host deleting the current slide leaves a valid one-slide selection', async () => {
    const e = await open(fiveSlides());
    e.selectSlide(1);
    e.selectSlide(2, 'range');
    const batch = e.session!.batch();
    deleteSlide(batch, 2);
    expect(e.session!.store.transact(batch.ops).ok).toBe(true);
    expect(e.snapshot().selectedSlides).toEqual([e.snapshot().currentSlide]);
  });
});
