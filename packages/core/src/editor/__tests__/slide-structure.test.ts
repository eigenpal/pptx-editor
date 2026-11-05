// Slide structure on add, duplicate, move, and delete:
// the section list (`p14:sectionLst`), parts a slide owns (comments, charts, embedded
// objects), and parts only a deleted slide used. Each test failed before its fix.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import {
  buildSlideStructureDeck,
  slideStructureDeck,
} from '../../../../../scripts/fixtures/slide-structure.ts';
import {
  attr,
  CONTENT_TYPES_PART,
  contentTypeOf,
  elements,
  NS,
  path,
  presentationInfo,
  readPptx,
  relationshipsOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { treeOf } from '../../automation/commands/slide-structure.ts';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

async function open(bytes = buildSlideStructureDeck()): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const pkgOf = (e: PptxEditorInstance): PptxPackage => e.session!.store.package;
const presentation = (e: PptxEditorInstance): Element =>
  xmlPartOf(pkgOf(e), e.session!.info.presentationPart)!.root;
const slideIds = (e: PptxEditorInstance): string[] =>
  elements(path(presentation(e), 'p:sldIdLst')).map((s) => attr(s, 'id')!);

/** Each section's slide ids, in order. */
function sections(e: PptxEditorInstance): string[][] {
  const out: string[][] = [];
  const walk = (el: Element): void => {
    if (el.ns === NS.p14 && el.local === 'sldIdLst')
      out.push(elements(el).map((s) => attr(s, 'id')!));
    for (const c of elements(el)) walk(c);
  };
  walk(presentation(e));
  return out;
}

/** The section list names every slide once, in slide order. */
function expectConsistent(e: PptxEditorInstance): void {
  expect(sections(e).flat()).toEqual(slideIds(e));
}

const slidePart = (e: PptxEditorInstance, i: number): string => e.session!.info.slides[i].partName;
const target = (pkg: PptxPackage, source: string, type: string): string | undefined =>
  relationshipsOf(pkg, source).find((r) => r.type === `${R}${type}`)?.target;
const overrides = (pkg: PptxPackage): string[] =>
  elements(xmlPartOf(pkg, CONTENT_TYPES_PART)!.root)
    .filter((el) => el.local === 'Override')
    .map((el) => attr(el, 'PartName')!);

/** Save, reopen, and check every internal relationship lands on a part. */
function reopened(e: PptxEditorInstance): PptxPackage {
  const r = readPptx(e.save()!);
  expect(r.ok).toBe(true);
  if (!r.ok) throw new Error('reopen');
  for (const source of ['/', ...r.pkg.parts.keys()])
    for (const rel of relationshipsOf(r.pkg, source))
      if (!rel.external) expect(r.pkg.parts.has(rel.target!)).toBe(true);
  return r.pkg;
}

describe('sections', () => {
  test('a new slide joins the section of the slide before it', async () => {
    const e = await open();
    e.goToSlide(1);
    expect(e.exec({ type: 'addSlide' }).ok).toBe(true);
    const added = slideIds(e)[2];
    expect(sections(e)).toEqual([
      ['256', '257', added],
      ['258', '259'],
    ]);
    expectConsistent(e);
  });

  test('a duplicate joins the section of its original, right after it', async () => {
    const e = await open();
    expect(e.exec({ type: 'duplicateSlide', index: 2 }).ok).toBe(true);
    const copy = slideIds(e)[3];
    expect(sections(e)).toEqual([
      ['256', '257'],
      ['258', copy, '259'],
    ]);
    expectConsistent(e);
  });

  test('a moved slide changes section with its place', async () => {
    const e = await open();
    expect(e.exec({ type: 'moveSlide', from: 0, to: 3 }).ok).toBe(true);
    expect(sections(e)).toEqual([['257'], ['258', '259', '256']]);
    expect(e.exec({ type: 'moveSlide', from: 3, to: 0 }).ok).toBe(true);
    expect(sections(e)).toEqual([
      ['256', '257'],
      ['258', '259'],
    ]);
    expectConsistent(e);
  });

  test('a deleted slide leaves its section', async () => {
    const e = await open();
    expect(e.exec({ type: 'deleteSlide', index: 1 }).ok).toBe(true);
    expect(sections(e)).toEqual([['256'], ['258', '259']]);
    expectConsistent(e);
  });

  test('a section whose slides are all deleted stays, empty', async () => {
    const e = await open();
    expect(e.exec({ type: 'deleteSlide', index: 0 }).ok).toBe(true);
    expect(e.exec({ type: 'deleteSlide', index: 0 }).ok).toBe(true);
    expect(sections(e)).toEqual([[], ['258', '259']]);
    reopened(e);
  });
});

/** Each custom show's slide part names, in order. */
function customShows(e: PptxEditorInstance): string[][] {
  const pres = e.session!.info.presentationPart;
  const rels = relationshipsOf(pkgOf(e), pres);
  return elements(path(presentation(e), 'p:custShowLst')).map((show) =>
    elements(path(show, 'p:sldLst')).map((sld) => {
      const rId = sld.attrs.find((a) => a.ns === NS.r && a.local === 'id')?.value;
      return rels.find((r) => r.id === rId)?.target ?? `missing ${rId}`;
    })
  );
}

describe('relationship ids', () => {
  test('a huge rId number in the file does not make two new slides share an id', async () => {
    const files = unzipSync(buildSlideStructureDeck());
    const xml = strFromU8(files['ppt/presentation.xml']!);
    files['ppt/presentation.xml'] = strToU8(
      xml.replace('<p:sldIdLst>', `<p:sldIdLst><p:sldId id="9001" r:id="rId${'9'.repeat(400)}"/>`)
    );
    const e = await open(zipSync(files));
    expect(e.exec({ type: 'addSlide' }).ok).toBe(true);
    expect(e.exec({ type: 'addSlide' }).ok).toBe(true);
    const ids = relationshipsOf(pkgOf(e), e.session!.info.presentationPart).map((r) => r.id);
    expect(new Set(ids).size).toBe(ids.length);
    expect(ids.every((id) => /^rId\d{1,9}$|^rId[A-Za-z]+$/.test(id))).toBe(true);
  });
});

describe('a slide entry that names no slide', () => {
  // A slide that one collaborator deleted while another moved it comes back in `p:sldIdLst`
  // with its old `r:id` and no relationship.
  async function withDeadEntry(): Promise<{ e: PptxEditorInstance; dead: string }> {
    const files = unzipSync(buildSlideStructureDeck());
    const rels = strFromU8(files['ppt/_rels/presentation.xml.rels']!);
    const max = Math.max(...[...rels.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1])));
    const dead = `rId${max + 1}`;
    const xml = strFromU8(files['ppt/presentation.xml']!);
    files['ppt/presentation.xml'] = strToU8(
      xml.replace('<p:sldIdLst>', `<p:sldIdLst><p:sldId id="9000" r:id="${dead}"/>`)
    );
    return { e: await open(zipSync(files)), dead };
  }

  test('New slide takes its place after the slide before it, and a new relationship id', async () => {
    const { e } = await withDeadEntry();
    const count = e.snapshot().slides.length;
    const before = slideIds(e);
    e.goToSlide(1);
    expect(e.exec({ type: 'addSlide' }).ok).toBe(true);
    expect(e.snapshot().slides.length).toBe(count + 1);
    const after = slideIds(e);
    const added = after.find((id) => !before.includes(id))!;
    // Right after the entry of slide 2, whatever the entry that names no slide does.
    expect(after[after.indexOf(before[2]!) + 1]).toBe(added);
    expect(e.snapshot().currentSlide).toBe(2);
    // No two entries name one relationship: the entry that names no slide never gets one.
    const entries = elements(path(presentation(e), 'p:sldIdLst'));
    const rIds = entries.map((x) => attr(x, 'r:id'));
    expect(new Set(rIds).size).toBe(rIds.length);
    // The entry that names no slide stays without one, or it is gone.
    const dead9000 = entries.find((x) => attr(x, 'id') === '9000');
    if (dead9000) {
      const live = relationshipsOf(pkgOf(e), e.session!.info.presentationPart).map((r) => r.id);
      expect(live).not.toContain(attr(dead9000, 'r:id'));
    }
  });
});

describe('custom shows', () => {
  test('a deleted slide leaves every custom show; the shows stay', async () => {
    const e = await open();
    expect(e.exec({ type: 'deleteSlide', index: 0 }).ok).toBe(true);
    expect(customShows(e)).toEqual([['/ppt/slides/slide2.xml', '/ppt/slides/slide4.xml'], []]);
    const pkg = reopened(e);
    const rels = relationshipsOf(pkg, presentationInfo(pkg).presentationPart);
    expect(rels.some((r) => r.target === '/ppt/slides/slide1.xml')).toBe(false);
  });

  test('a duplicate is not added to a custom show; the original stays in it', async () => {
    const e = await open();
    expect(e.exec({ type: 'duplicateSlide', index: 0 }).ok).toBe(true);
    expect(customShows(e)).toEqual([
      ['/ppt/slides/slide1.xml', '/ppt/slides/slide2.xml', '/ppt/slides/slide4.xml'],
      ['/ppt/slides/slide1.xml'],
    ]);
    reopened(e);
  });

  test('a moved slide keeps its custom show entries, and undo restores a deleted one', async () => {
    const e = await open();
    expect(e.exec({ type: 'moveSlide', from: 0, to: 3 }).ok).toBe(true);
    expect(customShows(e)[1]).toEqual(['/ppt/slides/slide1.xml']);
    expect(e.exec({ type: 'deleteSlide', index: 3 }).ok).toBe(true);
    expect(customShows(e)[1]).toEqual([]);
    e.exec({ type: 'undo' });
    expect(customShows(e)[1]).toEqual(['/ppt/slides/slide1.xml']);
  });
});

describe('duplicateSlide copies the parts a slide owns', () => {
  test('a chart, its workbook, and comments are copied, not shared', async () => {
    const e = await open();
    expect(e.exec({ type: 'duplicateSlide', index: 0 }).ok).toBe(true);
    const pkg = pkgOf(e);
    const [a, b] = [slidePart(e, 0), slidePart(e, 1)];
    const chartA = target(pkg, a, 'chart')!;
    const chartB = target(pkg, b, 'chart')!;
    expect(chartB).not.toBe(chartA);
    expect(chartB).toMatch(/^\/ppt\/charts\/chart\d+\.xml$/);
    expect(contentTypeOf(pkg, chartB)).toBe(contentTypeOf(pkg, chartA));
    const bookA = target(pkg, chartA, 'package')!;
    const bookB = target(pkg, chartB, 'package')!;
    expect(bookB).not.toBe(bookA);
    expect(pkg.parts.get(bookB)).toMatchObject({ kind: 'binary' });
    const commentsA = target(pkg, a, 'comments')!;
    const commentsB = target(pkg, b, 'comments')!;
    expect(commentsB).not.toBe(commentsA);
    expect(overrides(pkg)).toContain(commentsB);
    expect(overrides(pkg)).toContain(chartB);
    // Comment indexes stay unique per author; the reply follows its parent.
    const idx = (part: string) => elements(treeOf(pkg, part)!.root).map((cm) => attr(cm, 'idx'));
    expect(idx(commentsA)).toEqual(['1', '2']);
    expect(idx(commentsB)).toEqual(['3', '4']);
    const reply = elements(treeOf(pkg, commentsB)!.root)[1];
    const parent = JSON.stringify(reply);
    expect(parent).toContain('"local":"parentCm"');
    expect(parent).toMatch(/"local":"idx","value":"3"/);
    const author = elements(treeOf(pkg, '/ppt/commentAuthors.xml')!.root)[0];
    expect(attr(author, 'lastIdx')).toBe('4');
    reopened(e);
  });

  test('the copy gets its own creation id', async () => {
    const e = await open();
    e.exec({ type: 'duplicateSlide', index: 0 });
    const creation = (i: number): string | undefined => {
      const root = xmlPartOf(pkgOf(e), slidePart(e, i))!.root;
      const ext = elements(path(root, 'p:extLst'))[0];
      return attr(elements(ext)[0], 'val');
    };
    const all = e.session!.info.slides.map((_, i) => creation(i));
    expect(new Set(all).size).toBe(all.length);
  });

  test('an embedded object is copied; pictures stay shared', async () => {
    const e = await open();
    expect(e.exec({ type: 'duplicateSlide', index: 3 }).ok).toBe(true);
    const pkg = pkgOf(e);
    const [a, b] = [slidePart(e, 3), slidePart(e, 4)];
    expect(target(pkg, b, 'package')).not.toBe(target(pkg, a, 'package'));
    expect(target(pkg, b, 'image')).toBe(target(pkg, a, 'image'));
    reopened(e);
  });

  test('threaded comments are refused, never shared', async () => {
    const options = slideStructureDeck(new Uint8Array([1]), new Uint8Array([2]));
    const modern = 'http://schemas.microsoft.com/office/2018/10/relationships/comments';
    const bytes = buildPptx({
      ...options,
      slides: [
        {
          ...options.slides[2],
          rels: [...options.slides[2].rels!, ['rIdM', modern, '../comments/modernComment_1.xml']],
        },
      ],
      extraParts: [
        ...options.extraParts!,
        [
          'ppt/comments/modernComment_1.xml',
          'application/vnd.ms-powerpoint.comments+xml',
          '<p188:cmLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"/>',
        ],
      ],
      presentationExt: undefined,
    });
    const e = await open(bytes);
    const r = e.exec({ type: 'duplicateSlide', index: 0 });
    expect(r).toMatchObject({ ok: false, reason: 'disabled.slideNotCopyable' });
    expect(e.snapshot().slides.length).toBe(1);
  });
});

describe('deleteSlide removes parts only it used', () => {
  test('a picture only the deleted slide used goes; a shared one stays', async () => {
    const e = await open();
    expect(e.exec({ type: 'deleteSlide', index: 1 }).ok).toBe(true);
    const pkg = pkgOf(e);
    expect(pkg.parts.has('/ppt/media/image2.png')).toBe(false);
    expect(pkg.parts.has('/ppt/media/image1.png')).toBe(true);
    reopened(e);
  });

  test('a chart, its workbook, and comments go with their slide; the authors stay', async () => {
    const e = await open();
    expect(e.exec({ type: 'deleteSlide', index: 0 }).ok).toBe(true);
    const pkg = pkgOf(e);
    for (const part of [
      '/ppt/charts/chart1.xml',
      '/ppt/charts/_rels/chart1.xml.rels',
      '/ppt/embeddings/Embedded_Worksheet1.xlsx',
      '/ppt/comments/comment1.xml',
    ])
      expect(pkg.parts.has(part)).toBe(false);
    expect(overrides(pkg)).not.toContain('/ppt/charts/chart1.xml');
    expect(overrides(pkg)).not.toContain('/ppt/comments/comment1.xml');
    expect(pkg.parts.has('/ppt/commentAuthors.xml')).toBe(true);
    // The other workbook belongs to the embedded object on another slide.
    expect(pkg.parts.has('/ppt/embeddings/Embedded_Worksheet2.xlsx')).toBe(true);
    reopened(e);
    // One undo brings every part back.
    e.exec({ type: 'undo' });
    expect(pkgOf(e).parts.has('/ppt/charts/chart1.xml')).toBe(true);
    expect(pkgOf(e).parts.has('/ppt/embeddings/Embedded_Worksheet1.xlsx')).toBe(true);
  });

  test('deleting a duplicate keeps the original parts, and deleting the original keeps the copy', async () => {
    const e = await open();
    e.exec({ type: 'duplicateSlide', index: 0 });
    const chartA = target(pkgOf(e), slidePart(e, 0), 'chart')!;
    const chartB = target(pkgOf(e), slidePart(e, 1), 'chart')!;
    e.exec({ type: 'deleteSlide', index: 1 });
    expect(pkgOf(e).parts.has(chartA)).toBe(true);
    expect(pkgOf(e).parts.has(chartB)).toBe(false);
    e.exec({ type: 'undo' });
    e.exec({ type: 'deleteSlide', index: 0 });
    expect(pkgOf(e).parts.has(chartA)).toBe(false);
    expect(pkgOf(e).parts.has(chartB)).toBe(true);
    reopened(e);
  });

  test('a part a layout still references stays', async () => {
    const e = await open();
    // Give the layout a relationship to the picture only slide 2 uses, then delete slide 2.
    const layout = e.session!.info.slides[1].layoutPart!;
    const rels = xmlPartOf(pkgOf(e), layout.replace(/([^/]+)$/, '_rels/$1.rels'))!;
    e.session!.store.transact([
      {
        op: 'insertNode',
        part: layout.replace(/([^/]+)$/, '_rels/$1.rels'),
        parent: rels.root.id,
        index: rels.root.children.length,
        spec: {
          ns: NS.rel,
          local: 'Relationship',
          attrs: [
            { ns: '', local: 'Id', value: 'rId99' },
            { ns: '', local: 'Type', value: `${R}image` },
            { ns: '', local: 'Target', value: '../media/image2.png' },
          ],
        },
      },
    ]);
    expect(e.exec({ type: 'deleteSlide', index: 1 }).ok).toBe(true);
    expect(pkgOf(e).parts.has('/ppt/media/image2.png')).toBe(true);
    expect(presentationInfo(pkgOf(e)).slides.length).toBe(3);
  });
});
