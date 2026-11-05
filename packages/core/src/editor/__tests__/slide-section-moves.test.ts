// Moves across a section header: `moveSlide` and `moveSlides` follow one rule. A block joins
// the section of the slide before it; `section` puts it into the section named. The slide pane
// uses `section` for a drop below a section header or a Ctrl+Down step across one.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  buildSlideLayoutsDeck,
  LAYOUT_SECTION_IDS,
} from '../../../../../scripts/fixtures/slide-layouts.ts';
import { readPptx } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { deckSections } from '../../automation/commands/slide-sections.ts';

const [OPENING, DETAILS] = LAYOUT_SECTION_IDS as [string, string];
const EMPTY = '{5C1D2E3F-4A5B-4C6D-8E7F-9A0B1C2D3E4F}';

async function open(bytes = buildSlideLayoutsDeck()): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const sections = (e: PptxEditorInstance): string[] =>
  deckSections(e.session!.store.package).map((s) => `${s.name}:${s.firstSlide}+${s.slideCount}`);

const ids = (e: PptxEditorInstance): string[] => e.snapshot().slides.map((s) => s.id);

/** The fixture with an empty section between Opening and Details. */
function withEmptySection(): Uint8Array {
  const f = unzipSync(buildSlideLayoutsDeck());
  const xml = strFromU8(f['ppt/presentation.xml']!);
  f['ppt/presentation.xml'] = strToU8(
    xml.replace(
      `<p14:section name="Details"`,
      `<p14:section name="Empty" id="${EMPTY}"><p14:sldIdLst/></p14:section><p14:section name="Details"`
    )
  );
  return zipSync(f);
}

describe('moves across a section header', () => {
  test('a drop before the first slide of a later section joins that section', async () => {
    const e = await open();
    const before = ids(e);
    // Among the slides that do not move, position 1 is right before Details' first slide.
    expect(e.exec({ type: 'moveSlides', indices: [0], to: 1, section: DETAILS }).ok).toBe(true);
    expect(ids(e)).toEqual([before[1], before[0], ...before.slice(2)]);
    expect(sections(e)).toEqual(['Opening:0+1', 'Details:1+5']);
    // The same place without `section` (above the header) stays in the section before.
    e.exec({ type: 'undo' });
    expect(e.exec({ type: 'moveSlides', indices: [0], to: 1 }).ok).toBe(true);
    expect(sections(e)).toEqual(['Opening:0+2', 'Details:2+4']);
  });

  test('a step down from the last slide of a section makes it the first of the next', async () => {
    const e = await open();
    const before = ids(e);
    expect(e.exec({ type: 'moveSlides', indices: [1], to: 1, section: DETAILS }).ok).toBe(true);
    expect(ids(e)).toEqual(before);
    expect(sections(e)).toEqual(['Opening:0+1', 'Details:1+5']);
    // It is one undo step, and the saved file keeps it.
    const r = readPptx(e.save()!);
    expect(r.ok && deckSections(r.pkg).map((s) => s.slideCount)).toEqual([1, 5]);
    e.exec({ type: 'undo' });
    expect(sections(e)).toEqual(['Opening:0+2', 'Details:2+4']);
  });

  test('a step up from the first slide of a section makes it the last of the one before', async () => {
    const e = await open();
    expect(e.exec({ type: 'moveSlides', indices: [2], to: 2, section: OPENING }).ok).toBe(true);
    expect(sections(e)).toEqual(['Opening:0+3', 'Details:3+3']);
  });

  test('moveSlide follows the same rule as moveSlides', async () => {
    const a = await open();
    const b = await open();
    a.exec({ type: 'moveSlide', from: 0, to: 1, section: DETAILS });
    b.exec({ type: 'moveSlides', indices: [0], to: 1, section: DETAILS });
    expect(sections(a)).toEqual(sections(b));
    expect(ids(a)).toEqual(ids(b));
    // Without `section`, both join the section of the slide before: up or down.
    const c = await open();
    const d = await open();
    c.exec({ type: 'moveSlide', from: 4, to: 2 });
    d.exec({ type: 'moveSlides', indices: [4], to: 2 });
    expect(sections(c)).toEqual(['Opening:0+3', 'Details:3+3']);
    expect(sections(d)).toEqual(sections(c));
  });

  test('a block can fill an empty section between its neighbours', async () => {
    const e = await open(withEmptySection());
    expect(sections(e)).toEqual(['Opening:0+2', 'Empty:2+0', 'Details:2+4']);
    expect(e.exec({ type: 'moveSlides', indices: [1], to: 1, section: EMPTY }).ok).toBe(true);
    expect(sections(e)).toEqual(['Opening:0+1', 'Empty:1+1', 'Details:2+4']);
  });

  test('a section that is not next to the block, or unknown, is refused', async () => {
    const e = await open();
    const before = sections(e);
    expect(e.exec({ type: 'moveSlides', indices: [0], to: 0, section: DETAILS }).ok).toBe(false);
    expect(e.exec({ type: 'moveSlides', indices: [3], to: 3, section: 'nope' }).ok).toBe(false);
    expect(e.exec({ type: 'moveSlide', from: 4, to: 4, section: OPENING }).ok).toBe(false);
    expect(sections(e)).toEqual(before);
  });

  test('a move that changes nothing adds no undo step', async () => {
    const e = await open();
    const before = sections(e);
    // Slides 2 and 3 are already the first slides of Details, in place.
    expect(e.exec({ type: 'moveSlides', indices: [2, 3], to: 2, section: DETAILS }).ok).toBe(true);
    expect(e.exec({ type: 'moveSlide', from: 2, to: 2, section: DETAILS }).ok).toBe(true);
    expect(e.exec({ type: 'moveSlide', from: 1, to: 1 }).ok).toBe(true);
    expect(e.exec({ type: 'moveSlides', indices: [4], to: 4 }).ok).toBe(true);
    expect(sections(e)).toEqual(before);
    expect(e.snapshot().canUndo).toBe(false);
  });

  test('moveSlide follows the rule of moveSlides: up to a section start joins the section before', async () => {
    const e = await open();
    const before = ids(e);
    // Slide 4 moves up to place 2, the first place of Details: it follows Opening's last slide.
    expect(e.exec({ type: 'moveSlide', from: 4, to: 2 }).ok).toBe(true);
    expect(ids(e)).toEqual([before[0], before[1], before[4], before[2], before[3], before[5]]);
    expect(sections(e)).toEqual(['Opening:0+3', 'Details:3+3']);
    // With `section`, it joins Details, as moveSlide did before without it.
    e.exec({ type: 'undo' });
    expect(e.exec({ type: 'moveSlide', from: 4, to: 2, section: DETAILS }).ok).toBe(true);
    expect(sections(e)).toEqual(['Opening:0+2', 'Details:2+4']);
    const b = await open();
    expect(b.exec({ type: 'moveSlides', indices: [4], to: 2 }).ok).toBe(true);
    expect(sections(b)).toEqual(['Opening:0+3', 'Details:3+3']);
  });
});

describe('a new section in the saved part', () => {
  const declarations = (e: PptxEditorInstance): number =>
    (strFromU8(unzipSync(e.save()!)['ppt/presentation.xml']!).match(/xmlns:p14=/g) ?? []).length;

  test('declares p14 once, in a list that has it and in a new list', async () => {
    const e = await open();
    e.goToSlide(1);
    expect(e.exec({ type: 'addSection' }).ok).toBe(true);
    e.goToSlide(4);
    expect(e.exec({ type: 'addSection' }).ok).toBe(true);
    expect(declarations(e)).toBe(1);
    const flat = await open();
    expect(flat.exec({ type: 'removeAllSections' }).ok).toBe(true);
    flat.goToSlide(2);
    expect(flat.exec({ type: 'addSection' }).ok).toBe(true);
    expect(sections(flat).length).toBe(2);
    expect(declarations(flat)).toBe(1);
  });
});
