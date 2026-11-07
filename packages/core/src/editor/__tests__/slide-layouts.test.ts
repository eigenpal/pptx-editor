// Change layout, reset slide, hide slide, and New Slide with a layout (`parity-baseline` 4.4,
// `editor-ui-parity` 2.1–2.3), against the `slide-layouts` fixture. Every edit is checked
// after save and reopen, and undoes in one step.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { buildSlideLayoutsDeck, LAYOUTS } from '../../../../../scripts/fixtures/slide-layouts.ts';
import {
  attr,
  canonicalPackageFingerprint,
  child,
  elements,
  path,
  presentationInfo,
  readPptx,
  relationshipsOf,
  shapeTreeOf,
  textOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../../store/index.ts';
import { resolveSlide } from '../../cascade/index.ts';
import { createPptxEditor, toolbarCommandState, type PptxEditorInstance } from '../index.ts';

async function open(): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(buildSlideLayoutsDeck())).ok).toBe(true);
  return e;
}

const pkgOf = (e: PptxEditorInstance): PptxPackage => e.session!.store.package;
const layoutId = (e: PptxEditorInstance, name: string): string =>
  e.snapshot().layouts!.find((l) => l.name === name)!.id;

/** Save and reopen; every internal relationship lands on a part. */
function reopen(e: PptxEditorInstance): PptxPackage {
  const r = readPptx(e.save()!);
  if (!r.ok) throw new Error('reopen');
  for (const source of ['/', ...r.pkg.parts.keys()])
    for (const rel of relationshipsOf(r.pkg, source))
      if (!rel.external) expect(r.pkg.parts.has(rel.target!)).toBe(true);
  return r.pkg;
}

/** Each top-level shape of a slide: its name, `p:ph` type/idx/sz, own box, and text. */
function shapes(pkg: PptxPackage, index: number) {
  const part = presentationInfo(pkg).slides[index].partName;
  return elements(shapeTreeOf(pkg, part))
    .filter((el) => el.kind === 'sp' || el.kind === 'pic')
    .map((el: Element) => {
      const ph = path(el, 'p:nvSpPr', 'p:nvPr', 'p:ph');
      const texts: string[] = [];
      const walk = (n: Element): void => {
        if (n.local === 't') texts.push(textOf(n));
        for (const c of elements(n)) walk(c);
      };
      walk(el);
      return {
        name: attr(path(el, 'p:nvSpPr', 'p:cNvPr'), 'name'),
        ph: ph
          ? `${attr(ph, 'type') ?? 'obj'}:${attr(ph, 'idx') ?? ''}:${attr(ph, 'sz') ?? ''}`
          : null,
        xfrm: !!child(child(el, 'p:spPr'), 'a:xfrm'),
        text: texts.join(''),
      };
    });
}

const layoutNameOf = (pkg: PptxPackage, index: number): string | undefined =>
  attr(
    child(xmlPartOf(pkg, presentationInfo(pkg).slides[index].layoutPart)?.root, 'p:cSld'),
    'name'
  );

/**
 * The fixture with a second master ("Second" theme) that has one Title Slide layout, listed
 * before the first master's layouts in its own `p:sldLayoutIdLst`.
 */
function twoMasterDeck(): Uint8Array {
  const f = unzipSync(buildSlideLayoutsDeck());
  const text = (name: string): string => strFromU8(f[name]!);
  const put = (name: string, xml: string): void => void (f[name] = strToU8(xml));
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  put(
    'ppt/theme/theme3.xml',
    text('ppt/theme/theme1.xml').replace('name="Fixture"', 'name="Second"')
  );
  put(
    'ppt/slideLayouts/slideLayout7.xml',
    text('ppt/slideLayouts/slideLayout3.xml').replace('name="Title Slide"', 'name="Second Title"')
  );
  put(
    'ppt/slideLayouts/_rels/slideLayout7.xml.rels',
    text('ppt/slideLayouts/_rels/slideLayout3.xml.rels').replace(
      'slideMaster1.xml',
      'slideMaster2.xml'
    )
  );
  put(
    'ppt/slideMasters/slideMaster2.xml',
    text('ppt/slideMasters/slideMaster1.xml').replace(
      /<p:sldLayoutIdLst>.*<\/p:sldLayoutIdLst>/,
      '<p:sldLayoutIdLst><p:sldLayoutId id="2147483700" r:id="rId1"/></p:sldLayoutIdLst>'
    )
  );
  put(
    'ppt/slideMasters/_rels/slideMaster2.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${R}/slideLayout" Target="../slideLayouts/slideLayout7.xml"/>` +
      `<Relationship Id="rId2" Type="${R}/theme" Target="../theme/theme3.xml"/></Relationships>`
  );
  put(
    'ppt/presentation.xml',
    text('ppt/presentation.xml').replace(
      '</p:sldMasterIdLst>',
      '<p:sldMasterId id="2147483699" r:id="rIdM2"/></p:sldMasterIdLst>'
    )
  );
  put(
    'ppt/_rels/presentation.xml.rels',
    text('ppt/_rels/presentation.xml.rels').replace(
      '</Relationships>',
      `<Relationship Id="rIdM2" Type="${R}/slideMaster" Target="slideMasters/slideMaster2.xml"/></Relationships>`
    )
  );
  const CT = 'application/vnd.openxmlformats-officedocument';
  put(
    '[Content_Types].xml',
    text('[Content_Types].xml').replace(
      '</Types>',
      `<Override PartName="/ppt/slideMasters/slideMaster2.xml" ContentType="${CT}.presentationml.slideMaster+xml"/>` +
        `<Override PartName="/ppt/slideLayouts/slideLayout7.xml" ContentType="${CT}.presentationml.slideLayout+xml"/>` +
        `<Override PartName="/ppt/theme/theme3.xml" ContentType="${CT}.theme+xml"/></Types>`
    )
  );
  return zipSync(f);
}

describe('layouts in the snapshot', () => {
  test('lists the master’s layouts in order, and each slide names its layout', async () => {
    const e = await open();
    const layouts = e.snapshot().layouts!;
    expect(layouts.map((l) => [l.name, l.type])).toEqual([
      ['Title and Content', 'obj'],
      ['Blank', 'blank'],
      ['Title Slide', 'title'],
      ['Two Content', 'twoObj'],
      ['Title Only', 'titleOnly'],
      ['Quote', 'cust'],
    ]);
    expect(new Set(layouts.map((l) => l.masterId)).size).toBe(1);
    expect(e.snapshot().slides[2].layoutId).toBe(layoutId(e, 'Two Content'));
    // Reference-stable while the layouts do not change.
    e.exec({ type: 'setSlidesHidden', indices: [0], hidden: true });
    expect(e.snapshot().layouts).toBe(layouts);
  });
});

describe('change layout', () => {
  test('Title and Content to Two Content: content keeps its place, a new place is added', async () => {
    const e = await open();
    const before = canonicalPackageFingerprint(pkgOf(e));
    e.goToSlide(1);
    expect(
      e.exec({ type: 'applyLayout', indices: [1], layoutId: layoutId(e, 'Two Content') })
    ).toEqual({ ok: true });
    for (const pkg of [pkgOf(e), reopen(e)]) {
      expect(layoutNameOf(pkg, 1)).toBe('Two Content');
      expect(shapes(pkg, 1)).toEqual([
        { name: 'Title 1', ph: 'title::', xfrm: false, text: 'Agenda' },
        // The moved body takes the layout's place (its own box goes) and keeps its runs.
        {
          name: 'Content Placeholder 2',
          ph: 'obj:1:half',
          xfrm: false,
          text: 'Results by regionNext steps',
        },
        { name: 'Content Placeholder 3', ph: 'obj:2:half', xfrm: false, text: '' },
        // A shape that is not a placeholder never changes.
        { name: 'Note box', ph: null, xfrm: true, text: 'Draft' },
      ]);
    }
    // The bold run keeps its formatting.
    const bold = resolveSlide(pkgOf(e), 1)!
      .elements.flatMap((el) => (el.kind === 'shape' && el.text ? el.text.paragraphs : []))
      .flatMap((p) => p.runs)
      .find((r) => r.kind === 'text' && r.text === 'by region');
    expect(bold && bold.kind === 'text' && bold.props.bold).toBe(true);
    // One undo step back to the file as it was.
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalPackageFingerprint(pkgOf(e))).toBe(before);
    expect(e.snapshot().canUndo).toBe(false);
  });

  test('Title and Content to Title Slide: title to centered title, content to subtitle', async () => {
    const e = await open();
    e.exec({ type: 'applyLayout', indices: [1], layoutId: layoutId(e, 'Title Slide') });
    expect(shapes(reopen(e), 1).map((s) => [s.ph, s.text])).toEqual([
      ['ctrTitle::', 'Agenda'],
      ['subTitle:1:', 'Results by regionNext steps'],
      [null, 'Draft'],
    ]);
  });

  test('Two Content to Title Only: content without a place stays where it was', async () => {
    const e = await open();
    const was = resolveSlide(pkgOf(e), 2)!.elements.filter((el) => el.placeholder);
    e.exec({ type: 'applyLayout', indices: [2], layoutId: layoutId(e, 'Title Only') });
    const pkg = reopen(e);
    expect(shapes(pkg, 2).map((s) => [s.ph, s.xfrm, s.text])).toEqual([
      ['title::', false, 'Before and after'],
      ['obj:1:half', true, 'Manual steps'],
      ['obj:2:half', true, 'One click'],
    ]);
    const now = resolveSlide(pkg, 2)!.elements.filter((el) => el.placeholder);
    expect(now.slice(1).map((el) => el.xfrm)).toEqual(was.slice(1).map((el) => el.xfrm));
  });

  test('empty placeholders without a place go; to Blank leaves only content', async () => {
    const e = await open();
    e.exec({ type: 'applyLayout', indices: [4], layoutId: layoutId(e, 'Blank') });
    expect(shapes(reopen(e), 4)).toEqual([]);
  });

  test('several slides in one undo step', async () => {
    const e = await open();
    const before = canonicalPackageFingerprint(pkgOf(e));
    e.exec({ type: 'applyLayout', indices: [0, 4], layoutId: layoutId(e, 'Two Content') });
    expect(layoutNameOf(pkgOf(e), 0)).toBe('Two Content');
    expect(layoutNameOf(pkgOf(e), 4)).toBe('Two Content');
    // The subtitle text has a content place; the empty slide gets both content places.
    expect(shapes(pkgOf(e), 0).map((s) => s.ph)).toEqual(['title::', 'obj:1:half', 'obj:2:half']);
    expect(shapes(pkgOf(e), 4).map((s) => s.ph)).toEqual(['title::', 'obj:1:half', 'obj:2:half']);
    e.exec({ type: 'undo' });
    expect(canonicalPackageFingerprint(pkgOf(e))).toBe(before);
  });

  test('refuses an unknown layout, and edits nothing in viewing mode', async () => {
    const e = await open();
    expect(
      e.can({ type: 'applyLayout', indices: [0], layoutId: '/ppt/slides/slide1.xml' })
    ).toEqual({ ok: false, reason: 'errors.layoutNotFound' });
    expect(e.can({ type: 'applyLayout', indices: [9], layoutId: layoutId(e, 'Blank') }).ok).toBe(
      false
    );
    e.setMode('view');
    expect(e.can({ type: 'resetSlides', indices: [0] })).toEqual({
      ok: false,
      reason: 'disabled.viewing',
    });
  });
});

describe('several masters', () => {
  test('lists masters in order, each with its layouts and name; a slide moves to the other master', async () => {
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await e.load(twoMasterDeck())).ok).toBe(true);
    const layouts = e.snapshot().layouts!;
    expect(layouts.map((l) => [l.name, l.masterName])).toEqual([
      ['Title and Content', 'Fixture'],
      ['Blank', 'Fixture'],
      ['Title Slide', 'Fixture'],
      ['Two Content', 'Fixture'],
      ['Title Only', 'Fixture'],
      ['Quote', 'Fixture'],
      ['Second Title', 'Second'],
    ]);
    expect(new Set(layouts.map((l) => l.masterId)).size).toBe(2);
    expect(
      e.exec({ type: 'applyLayout', indices: [0], layoutId: layoutId(e, 'Second Title') }).ok
    ).toBe(true);
    const pkg = reopen(e);
    expect(layoutNameOf(pkg, 0)).toBe('Second Title');
    expect(shapes(pkg, 0).map((s) => [s.ph, s.text])).toEqual([
      ['ctrTitle::', 'Quarterly review'],
      ['subTitle:1:', 'Second quarter'],
    ]);
  });
});

describe('content placeholders', () => {
  const tableIn = (ph: string): string =>
    '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Table 2"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>' +
    `<p:nvPr>${ph}</p:nvPr></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="1828800"/><a:ext cx="4572000" cy="914400"/></p:xfrm>` +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="4572000"/></a:tblGrid>' +
    '<a:tr h="914400"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Cell</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>';

  test('a table in a content placeholder takes a content place and keeps its own box', async () => {
    const title =
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/>' +
      '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Numbers</a:t></a:r></a:p></p:txBody></p:sp>';
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect(
      (
        await e.load(
          buildPptx({
            layouts: LAYOUTS,
            slides: [{ layout: 'title', shapes: title + tableIn('<p:ph idx="1"/>') }],
          })
        )
      ).ok
    ).toBe(true);
    const frames = (pkg: PptxPackage): string[] =>
      elements(shapeTreeOf(pkg, presentationInfo(pkg).slides[0].partName))
        .filter((el) => el.kind === 'graphicFrame')
        .map((el) => {
          const ph = path(el, 'p:nvGraphicFramePr', 'p:nvPr', 'p:ph');
          const off = path(el, 'p:xfrm', 'a:off');
          return `${attr(ph, 'idx')}:${attr(ph, 'sz') ?? ''}@${attr(off, 'x')},${attr(off, 'y')}`;
        });
    e.exec({ type: 'applyLayout', indices: [0], layoutId: layoutId(e, 'Two Content') });
    let pkg = reopen(e);
    expect(frames(pkg)).toEqual(['1:half@914400,1828800']);
    // The second content place gets an empty placeholder.
    expect(shapes(pkg, 0).map((s) => s.ph)).toEqual(['title::', 'obj:2:half']);
    // No content place: the table stays where it is.
    e.exec({ type: 'applyLayout', indices: [0], layoutId: layoutId(e, 'Title Only') });
    pkg = reopen(e);
    expect(frames(pkg)).toEqual(['1:half@914400,1828800']);
    e.exec({ type: 'resetSlides', indices: [0] });
    expect(frames(reopen(e))).toEqual(['1:half@914400,1828800']);
  });
});

describe('reset slide', () => {
  test('placeholders lose their own box and formatting; text and other shapes stay', async () => {
    const e = await open();
    e.goToSlide(1);
    expect(e.exec({ type: 'resetSlides', indices: [1] }).ok).toBe(true);
    const pkg = reopen(e);
    expect(shapes(pkg, 1)).toEqual([
      { name: 'Title 1', ph: 'title::', xfrm: false, text: 'Agenda' },
      {
        name: 'Content Placeholder 2',
        ph: 'obj:1:',
        xfrm: false,
        text: 'Results by regionNext steps',
      },
      { name: 'Note box', ph: null, xfrm: true, text: 'Draft' },
    ]);
    const runs = resolveSlide(pkg, 1)!
      .elements.filter((el) => el.placeholder)
      .flatMap((el) => (el.kind === 'shape' && el.text ? el.text.paragraphs : []))
      .flatMap((p) => p.runs);
    for (const r of runs) if (r.kind === 'text') expect(r.props.bold).toBe(false);
    // The language stays on the runs.
    const part = presentationInfo(pkg).slides[1].partName;
    const xml = JSON.stringify(xmlPartOf(pkg, part)!.root);
    expect(xml.includes('"C00000"')).toBe(false);
    expect(xml.includes('"en-US"')).toBe(true);
  });

  test('clears a placeholder’s own list style, levels and run defaults too', async () => {
    const title =
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/>' +
      '<p:txBody><a:bodyPr anchor="b"><a:normAutofit/></a:bodyPr><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="4000" b="1"/></a:lvl1pPr></a:lstStyle>' +
      '<a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>Hi</a:t></a:r></a:p></p:txBody></p:sp>';
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await e.load(buildPptx({ slides: [{ shapes: title }] }))).ok).toBe(true);
    expect(e.exec({ type: 'resetSlides' })).toEqual({ ok: true });
    const pkg = reopen(e);
    const body = child(
      elements(shapeTreeOf(pkg, presentationInfo(pkg).slides[0].partName))[2],
      'p:txBody'
    );
    expect(elements(child(body, 'a:lstStyle'))).toEqual([]);
    expect(child(body, 'a:bodyPr')!.attrs).toEqual([]);
    expect(shapes(pkg, 0)[0]!.text).toBe('Hi');
  });

  test('a deleted placeholder comes back', async () => {
    const e = await open();
    e.goToSlide(0);
    const title = e.session!.layout(0)!.items.find((i) => i.element.name === 'Title 1')!;
    e.setSelection({ kind: 'shapes', nodeIds: [title.element.nodeId] });
    e.exec({ type: 'delete' });
    expect(shapes(pkgOf(e), 0).map((s) => s.ph)).toEqual(['subTitle:1:']);
    e.exec({ type: 'resetSlides', indices: [0] });
    expect(shapes(reopen(e), 0).map((s) => [s.ph, s.text])).toEqual([
      ['ctrTitle::', ''],
      ['subTitle:1:', 'Second quarter'],
    ]);
  });
});

describe('hide slide', () => {
  test('hides and shows slides through p:sld/@show, in one undo step each', async () => {
    const e = await open();
    expect(e.snapshot().slides.map((s) => s.hidden)).toEqual([
      false,
      false,
      false,
      true,
      false,
      false,
    ]);
    e.exec({ type: 'setSlidesHidden', indices: [0, 1], hidden: true });
    let pkg = reopen(e);
    const showOf = (p: PptxPackage, i: number): string | undefined =>
      attr(xmlPartOf(p, presentationInfo(p).slides[i].partName)!.root, 'show');
    expect([showOf(pkg, 0), showOf(pkg, 1)]).toEqual(['0', '0']);
    expect(
      e
        .snapshot()
        .slides.slice(0, 2)
        .map((s) => s.hidden)
    ).toEqual([true, true]);
    e.exec({ type: 'setSlidesHidden', indices: [1, 3], hidden: false });
    pkg = reopen(e);
    // A shown slide drops the attribute: its default is true.
    expect([showOf(pkg, 1), showOf(pkg, 3)]).toEqual([undefined, undefined]);
    e.exec({ type: 'undo' });
    expect(e.snapshot().slides.map((s) => s.hidden)).toEqual([
      true,
      true,
      false,
      true,
      false,
      false,
    ]);
    // Hiding a hidden slide changes nothing and adds no undo step.
    const revision = e.snapshot().revision;
    e.exec({ type: 'setSlidesHidden', indices: [3], hidden: true });
    expect(e.snapshot().revision).toBe(revision);
  });

  test('the toggle hides the slides unless all of them are hidden', async () => {
    const e = await open();
    const hidden = (): boolean[] => e.snapshot().slides.map((s) => s.hidden);
    // Slide 4 is hidden, slide 3 is not: the toggle hides both, whatever the current slide is.
    e.goToSlide(3);
    e.exec({ type: 'setSlidesHidden', indices: [2, 3] });
    expect(hidden().slice(2, 4)).toEqual([true, true]);
    // All of them hidden: the toggle shows them.
    e.goToSlide(0);
    e.exec({ type: 'setSlidesHidden', indices: [2, 3] });
    expect(hidden().slice(2, 4)).toEqual([false, false]);
    // The toolbar's Hide slide reads the same rule over the rail selection.
    e.selectSlide(3);
    expect(toolbarCommandState(e, 'slide.hide').active).toBe(false);
    e.exec({ type: 'setSlidesHidden' });
    expect(toolbarCommandState(e, 'slide.hide').active).toBe(true);
  });
});

describe('new slide', () => {
  test('takes the layout asked for, with its content placeholders', async () => {
    const e = await open();
    e.goToSlide(2);
    expect(e.exec({ type: 'addSlide', layoutId: layoutId(e, 'Quote') }).ok).toBe(true);
    expect(e.snapshot().currentSlide).toBe(3);
    const pkg = reopen(e);
    expect(layoutNameOf(pkg, 3)).toBe('Quote');
    expect(shapes(pkg, 3).map((s) => s.ph)).toEqual(['title::', 'body:13:quarter']);
  });

  test('after a title slide, takes Title and Content', async () => {
    const e = await open();
    e.goToSlide(0);
    e.exec({ type: 'addSlide' });
    expect(layoutNameOf(pkgOf(e), 1)).toBe('Title and Content');
    e.goToSlide(2);
    e.exec({ type: 'addSlide' });
    expect(layoutNameOf(pkgOf(e), 3)).toBe('Title and Content');
  });

  test('in a deck with no slides, takes the first Title Slide layout', async () => {
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await e.load(buildPptx({ slides: [], layouts: LAYOUTS }))).ok).toBe(true);
    expect(e.can({ type: 'addSlide' })).toEqual({ ok: true });
    expect(e.exec({ type: 'addSlide' })).toEqual({ ok: true });
    const pkg = reopen(e);
    expect(presentationInfo(pkg).slides.length).toBe(1);
    expect(layoutNameOf(pkg, 0)).toBe('Title Slide');
    expect(e.snapshot().currentSlide).toBe(0);
  });

  test('in a deck whose presentation has no p:sldIdLst, adds the list in its schema place', async () => {
    const f = unzipSync(buildPptx({ slides: [], layouts: LAYOUTS }));
    f['ppt/presentation.xml'] = strToU8(
      strFromU8(f['ppt/presentation.xml']!).replace(
        /<p:sldIdLst\s*\/>|<p:sldIdLst><\/p:sldIdLst>/,
        ''
      )
    );
    expect(strFromU8(f['ppt/presentation.xml']!)).not.toContain('sldIdLst');
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await e.load(zipSync(f))).ok).toBe(true);
    expect(e.exec({ type: 'addSlide' })).toEqual({ ok: true });
    const pkg = reopen(e);
    expect(presentationInfo(pkg).slides.length).toBe(1);
    const root = xmlPartOf(pkg, presentationInfo(pkg).presentationPart)!.root;
    const order = elements(root).map((el) => el.local);
    expect(order.indexOf('sldIdLst')).toBeLessThan(order.indexOf('sldSz'));
    expect(order.indexOf('sldIdLst')).toBeGreaterThan(order.indexOf('sldMasterIdLst'));
  });

  test('refuses a layout the deck does not have', async () => {
    const e = await open();
    expect(e.exec({ type: 'addSlide', layoutId: '/ppt/slideLayouts/none.xml' }).ok).toBe(false);
  });
});

describe('placeholders that keep their place', () => {
  const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
  const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
  const ph = (id: number, name: string, phXml: string, text = ''): string =>
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr>${phXml}</p:nvPr></p:nvSpPr><p:spPr/>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p>${text ? `<a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r>` : '<a:endParaRPr lang="en-US"/>'}</a:p></p:txBody></p:sp>`;
  const wrap = (choice: string, fallback: string): string =>
    `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:p14="${P14}" Requires="p14">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;
  const load = async (deck: Uint8Array): Promise<PptxEditorInstance> => {
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await e.load(deck)).ok).toBe(true);
    return e;
  };
  /** `p:ph` of every placeholder in the slide part, wrapped ones too, as `type:idx`. */
  const allPh = (pkg: PptxPackage): string[] => {
    const out: string[] = [];
    const walk = (el: Element): void => {
      if (el.local === 'ph') out.push(`${attr(el, 'type') ?? 'obj'}:${attr(el, 'idx') ?? ''}`);
      for (const c of elements(el)) walk(c);
    };
    walk(shapeTreeOf(pkg, presentationInfo(pkg).slides[0].partName)!);
    return out;
  };

  test('an empty placeholder that an animation names stays, so p:timing names no removed shape', async () => {
    const timing =
      '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"/></p:par></p:tnLst>' +
      '<p:bldLst><p:bldP spid="3" grpId="0"/></p:bldLst></p:timing>';
    const e = await load(
      buildPptx({
        layouts: LAYOUTS,
        slides: [
          {
            layout: 'title',
            shapes: ph(2, 'Title 1', '<p:ph type="title"/>') + ph(3, 'Body', '<p:ph idx="1"/>'),
            ext: timing,
          },
        ],
      })
    );
    const was = resolveSlide(pkgOf(e), 0)!.elements.find((el) => el.name === 'Body')!.xfrm;
    expect(e.exec({ type: 'applyLayout', indices: [0], layoutId: layoutId(e, 'Blank') }).ok).toBe(
      true
    );
    const pkg = reopen(e);
    expect(shapes(pkg, 0)).toEqual([{ name: 'Body', ph: 'obj:1:', xfrm: true, text: '' }]);
    expect(resolveSlide(pkg, 0)!.elements.find((el) => el.name === 'Body')!.xfrm).toEqual(was);
    const root = xmlPartOf(pkg, presentationInfo(pkg).slides[0].partName)!.root;
    expect(attr(path(root, 'p:timing', 'p:bldLst', 'p:bldP'), 'spid')).toBe('3');
  });

  test('a slide placeholder in mc:AlternateContent keeps its p:ph and box, and its place stays taken', async () => {
    const body = (text: string): string => ph(3, 'Content Placeholder 2', '<p:ph idx="1"/>', text);
    const e = await load(
      buildPptx({
        layouts: LAYOUTS,
        slides: [
          {
            layout: 'title',
            shapes:
              ph(2, 'Title 1', '<p:ph type="title"/>', 'Agenda') +
              wrap(body('Choice'), body('Fallback')),
          },
        ],
      })
    );
    const was = resolveSlide(pkgOf(e), 0)!.elements.find((el) => el.placeholder?.idx === '1')!.xfrm;
    e.exec({ type: 'applyLayout', indices: [0], layoutId: layoutId(e, 'Two Content') });
    const pkg = reopen(e);
    // The title takes its place; only the second content place gets a new placeholder, after
    // the wrapped one in layout order.
    expect(allPh(pkg)).toEqual(['title:', 'obj:1', 'obj:1', 'obj:2']);
    const tree = shapeTreeOf(pkg, presentationInfo(pkg).slides[0].partName)!;
    const ac = elements(tree).find((el) => el.local === 'AlternateContent')!;
    for (const branch of elements(ac))
      expect(!!path(elements(branch)[0], 'p:spPr', 'a:xfrm')).toBe(true);
    const now = resolveSlide(pkg, 0)!.elements.find((el) => el.placeholder?.idx === '1')!;
    expect(now.xfrm).toEqual(was);
  });

  test('a slide placeholder in nested mc:AlternateContent holds its place, with no second copy', async () => {
    const body = (text: string): string => ph(3, 'Content Placeholder 2', '<p:ph idx="1"/>', text);
    const e = await load(
      buildPptx({
        layouts: LAYOUTS,
        slides: [
          {
            layout: 'title',
            shapes:
              ph(2, 'Title 1', '<p:ph type="title"/>', 'Agenda') +
              wrap(wrap(body('Inner'), body('Inner fallback')), body('Fallback')),
          },
        ],
      })
    );
    const was = resolveSlide(pkgOf(e), 0)!.elements.find((el) => el.placeholder?.idx === '1')!.xfrm;
    e.exec({ type: 'applyLayout', indices: [0], layoutId: layoutId(e, 'Two Content') });
    const pkg = reopen(e);
    // Every copy keeps its p:ph; only the second content place gets a new placeholder.
    expect(allPh(pkg)).toEqual(['title:', 'obj:1', 'obj:1', 'obj:1', 'obj:2']);
    const drawn = resolveSlide(pkg, 0)!.elements.filter((el) => el.placeholder?.idx === '1');
    expect(drawn).toHaveLength(1);
    expect(drawn[0]!.xfrm).toEqual(was);
  });

  test('a layout place in mc:AlternateContent is a place, read from its fallback', async () => {
    const place = (id: number): string =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="13"/></p:nvPr></p:nvSpPr>` +
      '<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="4572000" cy="914400"/></a:xfrm></p:spPr>' +
      '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Prompt</a:t></a:r></a:p></p:txBody></p:sp>';
    const layouts = [
      ...LAYOUTS,
      {
        name: 'Wrapped',
        shapes: ph(2, 'Title 1', '<p:ph type="title"/>') + wrap(place(3), place(3)),
      },
    ];
    const e = await load(
      buildPptx({
        layouts,
        slides: [
          {
            layout: 'title',
            shapes:
              ph(2, 'Title 1', '<p:ph type="title"/>', 'Agenda') +
              ph(3, 'Content Placeholder 2', '<p:ph idx="1"/>', 'Kept'),
          },
          { layout: 'title', shapes: ph(2, 'Title 1', '<p:ph type="title"/>') },
        ],
      })
    );
    e.exec({ type: 'applyLayout', indices: [0, 1], layoutId: layoutId(e, 'Wrapped') });
    const pkg = reopen(e);
    expect(shapes(pkg, 0).map((s) => [s.ph, s.text])).toEqual([
      ['title::', 'Agenda'],
      ['body:13:', 'Kept'],
    ]);
    expect(shapes(pkg, 1).map((s) => s.ph)).toEqual(['title::', 'body:13:']);
  });

  test('a picture without a place takes a new idx when the new layout uses its own', async () => {
    const sectionHeader = {
      name: 'Section Header',
      type: 'secHead',
      shapes:
        ph(2, 'Title 1', '<p:ph type="title"/>') + ph(3, 'Text 2', '<p:ph type="body" idx="1"/>'),
    };
    const picture =
      '<p:pic><p:nvPicPr><p:cNvPr id="3" name="Picture 2"/><p:cNvPicPr><a:picLocks noGrp="1"/></p:cNvPicPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvPicPr>' +
      '<p:blipFill><a:blip/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr/></p:pic>';
    const e = await load(
      buildPptx({
        layouts: [...LAYOUTS, sectionHeader],
        slides: [
          { layout: 'title', shapes: ph(2, 'Title 1', '<p:ph type="title"/>', 'Photos') + picture },
        ],
      })
    );
    e.exec({ type: 'applyLayout', indices: [0], layoutId: layoutId(e, 'Section Header') });
    const pkg = reopen(e);
    // The picture keeps its place with a new idx; the body place gets its own placeholder.
    expect(allPh(pkg)).toEqual(['title:', 'body:1', 'obj:2']);
    const pic = elements(shapeTreeOf(pkg, presentationInfo(pkg).slides[0].partName)).find(
      (el) => el.kind === 'pic'
    )!;
    expect(!!path(pic, 'p:spPr', 'a:xfrm')).toBe(true);
  });

  test('an empty placeholder that an animation in mc:AlternateContent names stays', async () => {
    const timing = (spid: number): string =>
      '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"/></p:par></p:tnLst>' +
      `<p:bldLst><p:bldP spid="${spid}" grpId="0"/></p:bldLst></p:timing>`;
    for (const [choice, fallback] of [
      [3, 9],
      [9, 3],
    ] as const) {
      const e = await load(
        buildPptx({
          layouts: LAYOUTS,
          slides: [
            {
              layout: 'title',
              shapes: ph(2, 'Title 1', '<p:ph type="title"/>') + ph(3, 'Body', '<p:ph idx="1"/>'),
              ext: wrap(timing(choice), timing(fallback)),
            },
          ],
        })
      );
      e.exec({ type: 'applyLayout', indices: [0], layoutId: layoutId(e, 'Blank') });
      expect(shapes(reopen(e), 0).map((s) => s.name)).toEqual(['Body']);
    }
  });

  test('a wrapped body with idx 10 does not hold the date place, and takes a new idx', async () => {
    const dated = {
      name: 'Dated',
      shapes:
        ph(2, 'Title 1', '<p:ph type="title"/>') +
        ph(3, 'Date 2', '<p:ph type="dt" sz="half" idx="10"/>'),
    };
    const body = (text: string): string => ph(3, 'Text 2', '<p:ph type="body" idx="10"/>', text);
    const e = await load(
      buildPptx({
        layouts: [...LAYOUTS, dated],
        slides: [
          {
            layout: 'title',
            shapes:
              ph(2, 'Title 1', '<p:ph type="title"/>', 'Plan') +
              wrap(body('Choice'), body('Fallback')) +
              ph(4, 'Date 3', '<p:ph type="dt"/>', 'Today'),
          },
        ],
      })
    );
    e.exec({ type: 'applyLayout', indices: [0], layoutId: layoutId(e, 'Dated') });
    const pkg = reopen(e);
    // The date placeholder takes the date place; the wrapped body keeps its place in every
    // branch with an idx that no other placeholder uses.
    expect(allPh(pkg)).toEqual(['title:', 'body:1', 'body:1', 'dt:10']);
    expect(shapes(pkg, 0).find((s) => s.name === 'Date 3')).toMatchObject({ xfrm: false });
  });

  test('a placeholder without a place takes the smallest free idx, even next to idx 4294967295', async () => {
    const sectionHeader = {
      name: 'Section Header',
      type: 'secHead',
      shapes:
        ph(2, 'Title 1', '<p:ph type="title"/>') + ph(3, 'Text 2', '<p:ph type="body" idx="1"/>'),
    };
    const picture =
      '<p:pic><p:nvPicPr><p:cNvPr id="3" name="Picture 2"/><p:cNvPicPr><a:picLocks noGrp="1"/></p:cNvPicPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvPicPr>' +
      '<p:blipFill><a:blip/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr/></p:pic>';
    const e = await load(
      buildPptx({
        layouts: [...LAYOUTS, sectionHeader],
        slides: [
          {
            layout: 'title',
            shapes:
              ph(2, 'Title 1', '<p:ph type="title"/>', 'Photos') +
              picture +
              ph(4, 'Note', '<p:ph type="body" idx="4294967295"/>', 'Note'),
          },
        ],
      })
    );
    e.exec({ type: 'applyLayout', indices: [0], layoutId: layoutId(e, 'Section Header') });
    expect(allPh(reopen(e))).toEqual(['title:', 'obj:2', 'body:1']);
  });

  test('wrapped copies pair by shape id, not by their order in each branch', async () => {
    const a = ph(3, 'Content Placeholder 2', '<p:ph sz="half" idx="1"/>', 'A');
    const b = ph(4, 'Content Placeholder 3', '<p:ph sz="half" idx="2"/>', 'B');
    const e = await load(
      buildPptx({
        layouts: LAYOUTS,
        slides: [
          {
            layoutName: 'Two Content',
            shapes: ph(2, 'Title 1', '<p:ph type="title"/>', 'Two') + wrap(b + a, a + b),
          },
        ],
      })
    );
    e.exec({ type: 'applyLayout', indices: [0], layoutId: layoutId(e, 'Title Only') });
    const pkg = reopen(e);
    const tree = shapeTreeOf(pkg, presentationInfo(pkg).slides[0].partName)!;
    const ac = elements(tree).find((el) => el.local === 'AlternateContent')!;
    const boxes = (branch: Element): Record<string, string | undefined> =>
      Object.fromEntries(
        elements(branch).map((sp) => [
          attr(path(sp, 'p:nvSpPr', 'p:cNvPr'), 'id'),
          attr(path(sp, 'p:spPr', 'a:xfrm', 'a:off'), 'x'),
        ])
      );
    const [choice, fallback] = elements(ac);
    expect(boxes(choice!)).toEqual(boxes(fallback!));
    expect(boxes(fallback!)['3']).not.toBe(boxes(fallback!)['4']);
  });

  test('wrapped copies with other ids in each branch pair by their order', async () => {
    const copy = (id: number, idx: number, text: string): string =>
      ph(id, `Content Placeholder ${id}`, `<p:ph sz="half" idx="${idx}"/>`, text);
    const e = await load(
      buildPptx({
        layouts: LAYOUTS,
        slides: [
          {
            layoutName: 'Two Content',
            shapes:
              ph(2, 'Title 1', '<p:ph type="title"/>', 'Two') +
              wrap(copy(13, 1, 'A') + copy(14, 2, 'B'), copy(3, 1, 'A') + copy(4, 2, 'B')),
          },
        ],
      })
    );
    e.exec({ type: 'applyLayout', indices: [0], layoutId: layoutId(e, 'Title Only') });
    const pkg = reopen(e);
    const tree = shapeTreeOf(pkg, presentationInfo(pkg).slides[0].partName)!;
    const ac = elements(tree).find((el) => el.local === 'AlternateContent')!;
    const boxes = (branch: Element): (string | undefined)[] =>
      elements(branch).map((sp) => attr(path(sp, 'p:spPr', 'a:xfrm', 'a:off'), 'x'));
    const [choice, fallback] = elements(ac);
    // Content without a place keeps its own box, in both branches.
    expect(boxes(fallback!).every((x) => x !== undefined)).toBe(true);
    expect(boxes(choice!)).toEqual(boxes(fallback!));
    expect(boxes(fallback!)[0]).not.toBe(boxes(fallback!)[1]);
  });

  test('wrapped copies with partly shared ids pair by id first, then the rest by order', async () => {
    const copy = (id: number, idx: number, text: string): string =>
      ph(id, `Content Placeholder ${id}`, `<p:ph sz="half" idx="${idx}"/>`, text);
    const e = await load(
      buildPptx({
        layouts: LAYOUTS,
        slides: [
          {
            layoutName: 'Two Content',
            // The read copies (mc:Fallback) are [13, 3]; the mc:Choice copies are [3, 4].
            shapes:
              ph(2, 'Title 1', '<p:ph type="title"/>', 'Two') +
              wrap(copy(3, 2, 'B') + copy(4, 1, 'A'), copy(13, 1, 'A') + copy(3, 2, 'B')),
          },
        ],
      })
    );
    e.exec({ type: 'applyLayout', indices: [0], layoutId: layoutId(e, 'Title Only') });
    const pkg = reopen(e);
    const tree = shapeTreeOf(pkg, presentationInfo(pkg).slides[0].partName)!;
    const ac = elements(tree).find((el) => el.local === 'AlternateContent')!;
    const boxes = (branch: Element): Record<string, string | undefined> =>
      Object.fromEntries(
        elements(branch).map((sp) => [
          attr(path(sp, 'p:nvSpPr', 'p:cNvPr'), 'id'),
          attr(path(sp, 'p:spPr', 'a:xfrm', 'a:off'), 'x'),
        ])
      );
    const [choice, fallback] = elements(ac).map(boxes);
    // Read 3 takes the Choice copy 3 by id; read 13 takes the copy that is left, 4.
    expect(choice!['3']).toBe(fallback!['3']);
    expect(choice!['4']).toBe(fallback!['13']);
    expect(choice!['4']).toBeDefined();
    expect(choice!['3']).not.toBe(choice!['4']);
  });

  test('content takes a place before an empty placeholder of the same type', async () => {
    const e = await load(
      buildPptx({
        layouts: LAYOUTS,
        slides: [
          {
            layoutName: 'Two Content',
            shapes:
              ph(2, 'Title 1', '<p:ph type="title"/>', 'Quote') +
              ph(3, 'Empty body', '<p:ph type="body" idx="5"/>') +
              ph(4, 'Content', '<p:ph idx="1"/>', 'Keep'),
          },
        ],
      })
    );
    e.exec({ type: 'applyLayout', indices: [0], layoutId: layoutId(e, 'Quote') });
    expect(shapes(reopen(e), 0).map((s) => [s.name, s.ph, s.text])).toEqual([
      ['Title 1', 'title::', 'Quote'],
      ['Content', 'body:13:quarter', 'Keep'],
    ]);
  });
});

describe('the Layout control', () => {
  test('is enabled on a slide whose layout the master does not list', async () => {
    const f = unzipSync(buildSlideLayoutsDeck());
    const master = 'ppt/slideMasters/slideMaster1.xml';
    const xml = strFromU8(f[master]!);
    const unlisted = xml.replace(/<p:sldLayoutId id="\d+" r:id="rIdLayout6"\/>/, '');
    expect(unlisted).not.toBe(xml);
    f[master] = strToU8(unlisted);
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await e.load(zipSync(f))).ok).toBe(true);
    e.goToSlide(5);
    const own = e.snapshot().slides[5].layoutId;
    expect(e.snapshot().layouts!.some((l) => l.id === own)).toBe(false);
    const state = toolbarCommandState(e, 'slide.layout');
    expect(state.enabled).toBe(true);
    // The value is the slide's own layout, not the probe: no listed layout is marked current.
    expect(state.value).toBe(own);
  });
});
