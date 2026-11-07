// Commands on many selected slides (Select All in the slide list of a long deck): Delete, Hide
// Slide, and Duplicate read the deck once and write the slide list once, so their cost grows with the
// slides they change, not with its square. Their result is the result of one slide at a time:
// the parts, notes, content types, relationships, and sections.

import { expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import {
  attr,
  elements,
  find,
  presentationInfo,
  relationshipsOf,
  textOf,
  xmlPartOf,
  type PptxPackage,
} from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { readSlidesClipboard, savedSlideCopy, slidesClipboard } from '../slide-clipboard.ts';

const IN = 914_400;
const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';

/** `n` slides that say S0, S1, …, each with notes; with `sections`, two sections. */
function deck(n: number, sections = false): Uint8Array {
  const names = Array.from({ length: n }, (_, i) => `S${i}`);
  const sldId = (i: number) => `<p14:sldId id="${256 + i}"/>`;
  const half = Math.ceil(n / 2);
  const ext = sections
    ? `<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="${P14}">` +
      `<p14:section name="One" id="{00000000-0000-0000-0000-000000000001}"><p14:sldIdLst>${names
        .slice(0, half)
        .map((_, i) => sldId(i))
        .join('')}</p14:sldIdLst></p14:section>` +
      `<p14:section name="Two" id="{00000000-0000-0000-0000-000000000002}"><p14:sldIdLst>${names
        .slice(half)
        .map((_, i) => sldId(half + i))
        .join('')}</p14:sldIdLst></p14:section>` +
      '</p14:sectionLst></p:ext></p:extLst>'
    : undefined;
  return buildPptx({
    slides: names.map((name) => ({
      shapes: shape({ id: 2, name, x: IN, y: IN, cx: 4 * IN, cy: IN, text: para(run(name)) }),
      notes: `Notes of ${name}`,
    })),
    ...(ext ? { presentationExt: ext } : {}),
  });
}

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const pkgOf = (e: PptxEditorInstance): PptxPackage => e.session!.store.package;

/** The slide texts in deck order, with the text of each slide's notes. */
const slides = (e: PptxEditorInstance): string[] => {
  const pkg = pkgOf(e);
  return presentationInfo(pkg).slides.map((s) => {
    // A pasted slide can also hold empty placeholders of its layout: all text of the tree.
    const own = textOf(find(xmlPartOf(pkg, s.partName)!.root, 'p:spTree'));
    const notes = s.notesPart ? textOf(find(xmlPartOf(pkg, s.notesPart)!.root, 'p:txBody')) : '';
    return `${own}|${notes}`;
  });
};

/** The slide ids of each section, as the slide texts. */
const sections = (e: PptxEditorInstance): string[][] => {
  const pkg = pkgOf(e);
  const info = presentationInfo(pkg);
  const text = new Map(
    info.slides.map((s) => [s.sldId, textOf(find(xmlPartOf(pkg, s.partName)!.root, 'p:txBody'))])
  );
  const root = xmlPartOf(pkg, info.presentationPart)!.root;
  const lists: string[][] = [];
  const walk = (el: typeof root): void => {
    if (el.local === 'sldIdLst' && el.ns === P14)
      lists.push(elements(el).map((s) => text.get(attr(s, 'id') ?? '') ?? '?'));
    for (const c of elements(el)) walk(c);
  };
  walk(root);
  return lists;
};

/** Every slide and notes part has one content type override, and every slide one relationship. */
function expectConsistent(e: PptxEditorInstance): void {
  const pkg = pkgOf(e);
  const info = presentationInfo(pkg);
  const overrides = elements(xmlPartOf(pkg, '/[Content_Types].xml')!.root)
    .filter((o) => o.local === 'Override')
    .map((o) => (attr(o, 'PartName') ?? '').toLowerCase());
  const parts = [...pkg.parts.keys()].filter((p) => /\/(slides|notesSlides)\/[^/]+\.xml$/.test(p));
  expect(parts.length).toBe(info.slides.length * 2);
  for (const p of parts) expect(overrides.filter((o) => o === p.toLowerCase())).toHaveLength(1);
  const slideRels = relationshipsOf(pkg, info.presentationPart).filter((r) =>
    r.type.endsWith('/slide')
  );
  expect(slideRels).toHaveLength(info.slides.length);
  expect(new Set(info.slides.map((s) => s.sldId)).size).toBe(info.slides.length);
}

test('Delete of several slides drops their parts, notes, relationships, and section entries', async () => {
  const e = await open(deck(6, true));
  expect(e.exec({ type: 'deleteSlides', indices: [1, 3, 4] }).ok).toBe(true);
  expect(slides(e)).toEqual(['S0|Notes of S0', 'S2|Notes of S2', 'S5|Notes of S5']);
  expect(sections(e)).toEqual([['S0', 'S2'], ['S5']]);
  expectConsistent(e);
  const again = await open(e.save()!);
  expect(slides(again)).toEqual(slides(e));
  e.exec({ type: 'undo' });
  expect(slides(e)).toHaveLength(6);
  expect(sections(e)).toEqual([
    ['S0', 'S1', 'S2'],
    ['S3', 'S4', 'S5'],
  ]);
});

test('Duplicate of several slides puts the copies, with notes, after the last one in its section', async () => {
  const e = await open(deck(6, true));
  expect(e.exec({ type: 'duplicateSlides', indices: [0, 2] }).ok).toBe(true);
  expect(slides(e).map((s) => s.split('|')[0])).toEqual([
    'S0',
    'S1',
    'S2',
    'S0',
    'S2',
    'S3',
    'S4',
    'S5',
  ]);
  expect(slides(e)[3]).toBe('S0|Notes of S0');
  expect(sections(e)).toEqual([
    ['S0', 'S1', 'S2', 'S0', 'S2'],
    ['S3', 'S4', 'S5'],
  ]);
  expectConsistent(e);
  // Each copy has its own parts: the copies' notes name their own slide.
  const pkg = pkgOf(e);
  const info = presentationInfo(pkg);
  for (const s of info.slides.slice(3, 5)) {
    const back = relationshipsOf(pkg, s.notesPart!).find((r) => r.type.endsWith('/slide'));
    expect(back?.target).toBe(s.partName);
  }
  const again = await open(e.save()!);
  expect(slides(again)).toEqual(slides(e));
  expect(sections(again)).toEqual(sections(e));
});

test('a slide copy writes only the copied slides, and pastes them in their order', async () => {
  const e = await open(deck(40, true));
  const raw = slidesClipboard(e.session!, [7, 2])!;
  const read = readSlidesClipboard(raw)!;
  expect(read.slides).toEqual([1, 0]);
  const file = await open(read.file);
  expect(slides(file)).toEqual(['S2|Notes of S2', 'S7|Notes of S7']);
  expectConsistent(file);
  const into = await open(deck(2));
  expect(into.exec({ type: 'pasteSlides', ...read }).ok).toBe(true);
  expect(slides(into)).toEqual([
    'S0|Notes of S0',
    'S7|Notes of S7',
    'S2|Notes of S2',
    'S1|Notes of S1',
  ]);
});

test('copies of different slides keep one saved file per revision, and a repeat reuses it', async () => {
  const e = await open(deck(6));
  const pkg = pkgOf(e);
  const first = slidesClipboard(e.session!, [0]);
  expect(savedSlideCopy(pkg)).toBe('0');
  slidesClipboard(e.session!, [1]);
  slidesClipboard(e.session!, [2, 3]);
  expect(savedSlideCopy(pkg)).toBe('2,3');
  const again = slidesClipboard(e.session!, [3, 2]);
  expect(readSlidesClipboard(again!)!.slides).toEqual([1, 0]);
  expect(savedSlideCopy(pkg)).toBe('2,3');
  // A copy of slide 0 again saves it again, with the same slides.
  expect(slidesClipboard(e.session!, [0])).toBe(first);
  expect(savedSlideCopy(pkg)).toBe('0');
});

/** The fastest of three runs of `command` on a deck of `n` slides, in milliseconds. */
async function fastest(
  n: number,
  command: (n: number) => Parameters<PptxEditorInstance['exec']>[0]
) {
  const bytes = deck(n);
  let best = Infinity;
  for (let k = 0; k < 3; k += 1) {
    const e = await open(bytes);
    const t0 = performance.now();
    expect(e.exec(command(n)).ok).toBe(true);
    best = Math.min(best, performance.now() - t0);
    e.destroy();
  }
  return best;
}

const all = (from: number, n: number) => Array.from({ length: n - from }, (_, i) => from + i);

// One slide at a time, these took 3 s (Delete), 11 s (Duplicate), and 0.8 s (Hide Slide) on 600
// slides: each slide read the whole deck again.
test('Delete of every slide but one of 600 takes less than 0.8 s', async () => {
  expect(await fastest(600, (n) => ({ type: 'deleteSlides', indices: all(1, n) }))).toBeLessThan(
    800
  );
});

test('Duplicate of every slide of 600 takes less than 2.5 s', async () => {
  expect(await fastest(600, (n) => ({ type: 'duplicateSlides', indices: all(0, n) }))).toBeLessThan(
    2_500
  );
});

test('Hide Slide on every slide of 1,200 takes less than 0.8 s, and the toggle shows them again', async () => {
  expect(
    await fastest(1_200, (n) => ({ type: 'setSlidesHidden', indices: all(0, n) }))
  ).toBeLessThan(800);
  const e = await open(deck(4));
  expect(e.exec({ type: 'setSlidesHidden', indices: [0, 2] }).ok).toBe(true);
  const hidden = () => presentationInfo(pkgOf(e)).slides.map((s) => s.hidden);
  expect(hidden()).toEqual([true, false, true, false]);
  // The toggle shows them again when every one of them is hidden.
  expect(e.exec({ type: 'setSlidesHidden', indices: [0, 2] }).ok).toBe(true);
  expect(hidden()).toEqual([false, false, false, false]);
});
