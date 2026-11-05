// Repeated slide entries in `p:sldIdLst`, which concurrent edits can leave: the slide list that the editor shows,
// the list that commands act on, and the list that the writer saves are the same list. An
// entry that repeats a slide part goes; an entry that repeats only a slide id is another slide
// and gets a fresh id (ECMA-376 Part 1, 19.2.1.33 `sldId`).

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildSlideLayoutsDeck } from '../../../../../scripts/fixtures/slide-layouts.ts';
import {
  attr,
  elements,
  path,
  presentationInfo,
  readPptx,
  repairedPresentationRoot,
  textOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { deckSections } from '../../automation/commands/slide-sections.ts';

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const pkgOf = (e: PptxEditorInstance): PptxPackage => e.session!.store.package;

function reopened(e: PptxEditorInstance): PptxPackage {
  const r = readPptx(e.save()!);
  if (!r.ok) throw new Error('reopen');
  return r.pkg;
}

/** The fixture with its presentation part (and, with `rels`, its relationships) rewritten. */
function patched(edit: (xml: string) => string, rels?: (xml: string) => string): Uint8Array {
  const f = unzipSync(buildSlideLayoutsDeck());
  f['ppt/presentation.xml'] = strToU8(edit(strFromU8(f['ppt/presentation.xml']!)));
  if (rels) {
    const name = 'ppt/_rels/presentation.xml.rels';
    f[name] = strToU8(rels(strFromU8(f[name]!)));
  }
  return zipSync(f);
}

/** The title text of each slide, in order. */
function titles(pkg: PptxPackage): string[] {
  return presentationInfo(pkg).slides.map((s) => {
    const tree = path(xmlPartOf(pkg, s.partName)?.root, 'p:cSld', 'p:spTree');
    const out: string[] = [];
    const walk = (el: Element | undefined): void => {
      if (!el) return;
      if (el.local === 't') out.push(textOf(el));
      for (const c of elements(el)) walk(c);
    };
    walk(elements(tree).find((el) => el.local === 'sp'));
    return out.join('');
  });
}

/** The slide ids of `p:sldIdLst` in the part, and the slide parts they name. */
function entries(pkg: PptxPackage): string[] {
  const root = xmlPartOf(pkg, presentationInfo(pkg).presentationPart)!.root;
  return elements(elements(root).find((e) => e.local === 'sldIdLst')).map(
    (e) => attr(e, 'id') ?? ''
  );
}

const sections = (pkg: PptxPackage): string[] =>
  deckSections(pkg).map((s) => `${s.name}:${s.firstSlide}+${s.slideCount}`);

describe('repeated slide entries', () => {
  test('an entry that repeats a slide part shows once, and Delete removes the right slide', async () => {
    const bytes = patched((xml) =>
      xml.replace(
        '<p:sldId id="258" r:id="rId103"/>',
        '<p:sldId id="258" r:id="rId103"/>'.repeat(2)
      )
    );
    const e = await open(bytes);
    const clean = await open(buildSlideLayoutsDeck());
    const before = titles(pkgOf(clean));
    // The rail shows each slide once, as the file will hold it.
    expect(e.snapshot().slides.length).toBe(6);
    expect(titles(pkgOf(e))).toEqual(before);
    // The slide after the repeated entry is index 3: Delete removes that slide, not slide 3.
    expect(e.exec({ type: 'deleteSlides', indices: [3] }).ok).toBe(true);
    const after = titles(pkgOf(e));
    expect(after).toEqual(before.filter((_, i) => i !== 3));
    expect(titles(reopened(e))).toEqual(after);
    expect(entries(pkgOf(e))).toEqual(['256', '257', '258', '260', '261']);
  });

  test('two relationships to one slide part are one slide', async () => {
    const bytes = patched(
      (xml) => xml.replace('</p:sldIdLst>', '<p:sldId id="300" r:id="rId999"/></p:sldIdLst>'),
      (xml) =>
        xml.replace(
          '</Relationships>',
          '<Relationship Id="rId999" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/></Relationships>'
        )
    );
    const e = await open(bytes);
    expect(e.snapshot().slides.length).toBe(6);
    expect(entries(reopened(e))).toEqual(['256', '257', '258', '259', '260', '261']);
  });

  test('an entry that repeats only a slide id keeps its slide, with a fresh id', async () => {
    // Slide 4 (rId104) takes the id of slide 3, as two replicas that each add a slide can do.
    const bytes = patched((xml) =>
      xml.replace('<p:sldId id="259" r:id="rId104"/>', '<p:sldId id="258" r:id="rId104"/>')
    );
    const clean = await open(buildSlideLayoutsDeck());
    const e = await open(bytes);
    expect(e.snapshot().slides.length).toBe(6);
    expect(titles(pkgOf(e))).toEqual(titles(pkgOf(clean)));
    // The fresh id is past every id that the part names, in the range of ST_SlideId.
    const ids = presentationInfo(pkgOf(e)).slides.map((s) => s.sldId);
    expect(ids).toEqual(['256', '257', '258', '262', '260', '261']);
    const pkg = reopened(e);
    expect(entries(pkg)).toEqual(ids);
    expect(titles(pkg)).toEqual(titles(pkgOf(clean)));
    // No section names the fresh id: the slide stays in the section of the slide before it.
    expect(sections(pkg)).toEqual(['Opening:0+2', 'Details:2+4']);
    expect(strFromU8(unzipSync(e.save()!)['ppt/presentation.xml']!)).toContain(
      '<p14:sldId id="262"/>'
    );
  });

  test('a command renews the repeated id in place, and Delete acts on the slide shown', async () => {
    const bytes = patched((xml) =>
      xml.replace('<p:sldId id="259" r:id="rId104"/>', '<p:sldId id="258" r:id="rId104"/>')
    );
    const e = await open(bytes);
    const before = titles(pkgOf(e));
    expect(e.exec({ type: 'deleteSlides', indices: [4] }).ok).toBe(true);
    const pkg = pkgOf(e);
    expect(titles(pkg)).toEqual(before.filter((_, i) => i !== 4));
    expect(entries(pkg)).toEqual(['256', '257', '258', '262', '261']);
    // The part is as the writer saves it: no second repair.
    const root = xmlPartOf(pkg, presentationInfo(pkg).presentationPart)!.root;
    expect(repairedPresentationRoot(pkg)).toBe(root);
    expect(sections(pkg)).toEqual(['Opening:0+2', 'Details:2+3']);
    // Undo brings the part back as it was.
    e.exec({ type: 'undo' });
    expect(entries(pkgOf(e))).toEqual(['256', '257', '258', '258', '260', '261']);
    expect(e.snapshot().slides.length).toBe(6);
  });

  test('a fresh id skips the ids that only a section names', async () => {
    const bytes = patched((xml) =>
      xml
        .replace('<p:sldId id="259" r:id="rId104"/>', '<p:sldId id="258" r:id="rId104"/>')
        .replace('<p14:sldId id="261"/>', '<p14:sldId id="261"/><p14:sldId id="262"/>')
    );
    const e = await open(bytes);
    expect(presentationInfo(pkgOf(e)).slides.map((s) => s.sldId)[3]).toBe('263');
    expect(sections(reopened(e))).toEqual(['Opening:0+2', 'Details:2+4']);
  });

  test('two slide lists join into the first, in document order', async () => {
    // Two replicas that each add the first slide to a deck without a list each create one.
    const bytes = patched((xml) =>
      xml.replace(
        '<p:sldId id="259" r:id="rId104"/>',
        '</p:sldIdLst><p:sldIdLst><p:sldId id="259" r:id="rId104"/>'
      )
    );
    const clean = await open(buildSlideLayoutsDeck());
    const e = await open(bytes);
    expect(e.snapshot().slides.length).toBe(6);
    expect(titles(pkgOf(e))).toEqual(titles(pkgOf(clean)));
    const pkg = reopened(e);
    const root = xmlPartOf(pkg, presentationInfo(pkg).presentationPart)!.root;
    expect(elements(root).filter((el) => el.local === 'sldIdLst').length).toBe(1);
    expect(entries(pkg)).toEqual(['256', '257', '258', '259', '260', '261']);
    expect(sections(pkg)).toEqual(['Opening:0+2', 'Details:2+4']);
  });

  test('a command joins two slide lists in place, and Delete acts on the slide shown', async () => {
    const bytes = patched((xml) =>
      xml.replace(
        '<p:sldId id="259" r:id="rId104"/>',
        '</p:sldIdLst><p:sldIdLst><p:sldId id="259" r:id="rId104"/>'
      )
    );
    const e = await open(bytes);
    const before = titles(pkgOf(e));
    expect(e.exec({ type: 'deleteSlides', indices: [4] }).ok).toBe(true);
    const pkg = pkgOf(e);
    const root = xmlPartOf(pkg, presentationInfo(pkg).presentationPart)!.root;
    expect(elements(root).filter((el) => el.local === 'sldIdLst').length).toBe(1);
    expect(titles(pkg)).toEqual(before.filter((_, i) => i !== 4));
    expect(entries(pkg)).toEqual(['256', '257', '258', '259', '261']);
    expect(repairedPresentationRoot(pkg)).toBe(root);
    e.exec({ type: 'undo' });
    expect(e.snapshot().slides.length).toBe(6);
  });

  test('an entry with a missing, non-integer, or out-of-range id gets a fresh id', async () => {
    const bytes = patched((xml) =>
      xml
        .replace('<p:sldId id="256" r:id="rId101"/>', '<p:sldId r:id="rId101"/>')
        .replace('<p:sldId id="257" r:id="rId102"/>', '<p:sldId id="x7" r:id="rId102"/>')
        .replace('<p:sldId id="258" r:id="rId103"/>', '<p:sldId id="100" r:id="rId103"/>')
        .replace('<p:sldId id="259" r:id="rId104"/>', '<p:sldId id="2147483648" r:id="rId104"/>')
    );
    const clean = await open(buildSlideLayoutsDeck());
    const e = await open(bytes);
    expect(titles(pkgOf(e))).toEqual(titles(pkgOf(clean)));
    const ids = presentationInfo(pkgOf(e)).slides.map((s) => s.sldId);
    expect(ids).toEqual(['262', '263', '264', '265', '260', '261']);
    const pkg = reopened(e);
    expect(entries(pkg)).toEqual(ids);
    expect(titles(pkg)).toEqual(titles(pkgOf(clean)));
    // A command renews the ids in place, as the writer saves them.
    expect(e.exec({ type: 'moveSlides', indices: [5], to: 4 }).ok).toBe(true);
    const live = pkgOf(e);
    expect(entries(live)).toEqual(['262', '263', '264', '265', '261', '260']);
    const root = xmlPartOf(live, presentationInfo(live).presentationPart)!.root;
    expect(repairedPresentationRoot(live)).toBe(root);
  });

  test('an invalid id that a section names keeps the slide in that section', async () => {
    // Slide 3 opens Details. Its id "100" is not an ST_SlideId; Details names it by "100".
    const bytes = patched((xml) =>
      xml
        .replace('<p:sldId id="258" r:id="rId103"/>', '<p:sldId id="100" r:id="rId103"/>')
        .replace('<p14:sldId id="258"/>', '<p14:sldId id="100"/>')
    );
    const e = await open(bytes);
    expect(presentationInfo(pkgOf(e)).slides.map((s) => s.sldId)[2]).toBe('262');
    expect(sections(pkgOf(e))).toEqual(['Opening:0+2', 'Details:2+4']);
    const pkg = reopened(e);
    expect(entries(pkg)).toEqual(['256', '257', '262', '259', '260', '261']);
    expect(sections(pkg)).toEqual(['Opening:0+2', 'Details:2+4']);
    // A command renews the id in place, and the section names the fresh id.
    expect(e.exec({ type: 'moveSlides', indices: [5], to: 4 }).ok).toBe(true);
    const live = pkgOf(e);
    expect(sections(live)).toEqual(['Opening:0+2', 'Details:2+4']);
    const root = xmlPartOf(live, presentationInfo(live).presentationPart)!.root;
    expect(repairedPresentationRoot(live)).toBe(root);
    expect(sections(reopened(e))).toEqual(['Opening:0+2', 'Details:2+4']);
  });

  test('a repeated invalid id does not take over the section entry', async () => {
    // Slides 3 and 4 both have "100"; Details names "100" once. Neither renewed id inherits it.
    const bytes = patched((xml) =>
      xml
        .replace('<p:sldId id="258" r:id="rId103"/>', '<p:sldId id="100" r:id="rId103"/>')
        .replace('<p:sldId id="259" r:id="rId104"/>', '<p:sldId id="100" r:id="rId104"/>')
        .replace('<p14:sldId id="258"/>', '<p14:sldId id="100"/>')
    );
    const e = await open(bytes);
    expect(e.snapshot().slides.length).toBe(6);
    const pkg = reopened(e);
    expect(entries(pkg)).toEqual(['256', '257', '262', '263', '260', '261']);
    // No section names the fresh ids: both join the section of the slide before them.
    expect(sections(pkg)).toEqual(['Opening:0+4', 'Details:4+2']);
  });

  test('an id such as "1e3" sets no maximum: the fresh id follows the largest ST_SlideId', async () => {
    const bytes = patched((xml) =>
      xml.replace('<p:sldId id="259" r:id="rId104"/>', '<p:sldId id="1e3" r:id="rId104"/>')
    );
    const e = await open(bytes);
    const ids = presentationInfo(pkgOf(e)).slides.map((s) => s.sldId);
    expect(ids).toEqual(['256', '257', '258', '262', '260', '261']);
    expect(entries(reopened(e))).toEqual(ids);
  });

  test('ids compare as numbers: "0258" next to "258" gets a fresh id', async () => {
    const bytes = patched((xml) =>
      xml.replace('<p:sldId id="259" r:id="rId104"/>', '<p:sldId id="0258" r:id="rId104"/>')
    );
    const clean = await open(buildSlideLayoutsDeck());
    const e = await open(bytes);
    expect(e.snapshot().slides.length).toBe(6);
    const ids = presentationInfo(pkgOf(e)).slides.map((s) => s.sldId);
    expect(ids).toEqual(['256', '257', '258', '262', '260', '261']);
    const pkg = reopened(e);
    expect(entries(pkg)).toEqual(ids);
    expect(titles(pkg)).toEqual(titles(pkgOf(clean)));
    expect(sections(pkg)).toEqual(['Opening:0+2', 'Details:2+4']);
  });

  test('section moves read "0258" in p:sldIdLst as the section entry "258"', async () => {
    const bytes = patched((xml) =>
      xml.replace('<p:sldId id="258" r:id="rId103"/>', '<p:sldId id="0258" r:id="rId103"/>')
    );
    const e = await open(bytes);
    expect(presentationInfo(pkgOf(e)).slides[2]!.sldId).toBe('0258');
    expect(sections(pkgOf(e))).toEqual(['Opening:0+2', 'Details:2+4']);
    const [opening, details] = deckSections(pkgOf(e)).map((s) => s.id);
    // Slide 3 is already the first slide of Details: the move changes nothing, so no undo step.
    expect(e.exec({ type: 'moveSlide', from: 2, to: 2, section: details! }).ok).toBe(true);
    expect(e.snapshot().canUndo).toBe(false);
    // Slide 4 follows slide 3 of Details: it cannot join Opening without a reorder.
    expect(e.exec({ type: 'moveSlide', from: 3, to: 3, section: opening! }).ok).toBe(false);
    expect(sections(pkgOf(e))).toEqual(['Opening:0+2', 'Details:2+4']);
    // Slide 3 steps up into Opening: its section entry moves, and no second entry appears.
    expect(e.exec({ type: 'moveSlide', from: 2, to: 2, section: opening! }).ok).toBe(true);
    expect(sections(pkgOf(e))).toEqual(['Opening:0+3', 'Details:3+3']);
    expect(sections(reopened(e))).toEqual(['Opening:0+3', 'Details:3+3']);
  });

  test('a slide relationship whose target differs only in case names the slide', async () => {
    // OPC part names fold ASCII case; the rail and the writer both count the slide.
    const bytes = patched(
      (xml) => xml,
      (xml) => xml.replace('Target="slides/slide3.xml"', 'Target="Slides/SLIDE3.xml"')
    );
    const clean = await open(buildSlideLayoutsDeck());
    const e = await open(bytes);
    expect(e.snapshot().slides.length).toBe(6);
    expect(presentationInfo(pkgOf(e)).slides[2]!.partName).toBe('/ppt/slides/slide3.xml');
    expect(titles(pkgOf(e))).toEqual(titles(pkgOf(clean)));
    expect(entries(reopened(e))).toEqual(['256', '257', '258', '259', '260', '261']);
  });
});
