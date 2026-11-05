// Sections (`p14:sectionLst`, `editor-ui-parity` 2.5): add, rename, remove, and move, each
// one undo step, checked after save and reopen. The section list always names every slide once,
// in slide order.

import { describe, expect, spyOn, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import {
  buildSlideLayoutsDeck,
  LAYOUT_SECTION_IDS,
} from '../../../../../scripts/fixtures/slide-layouts.ts';
import {
  attr,
  canonicalPackageFingerprint,
  child,
  elements,
  NS,
  path,
  presentationInfo,
  readPptx,
  repairedPresentationRoot,
  textOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../../store/index.ts';
import {
  commandForSlot,
  createPptxEditor,
  resetDevWarnings,
  runToolbarCommand,
  toolbarCommandState,
  type ChromeSlotId,
  type PptxEditorInstance,
} from '../index.ts';
import { deckSections } from '../../automation/commands/slide-sections.ts';

async function open(bytes = buildSlideLayoutsDeck()): Promise<PptxEditorInstance> {
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

/** Each section as `name: first title words…`, from the file. */
function sections(pkg: PptxPackage): string[] {
  return deckSections(pkg).map((s) => `${s.name}:${s.firstSlide}+${s.slideCount}`);
}

/** The section list names every slide once, in slide order. */
function expectConsistent(pkg: PptxPackage): void {
  const root = xmlPartOf(pkg, presentationInfo(pkg).presentationPart)!.root;
  const named: string[] = [];
  const walk = (el: Element): void => {
    if (el.ns === NS.p14 && el.local === 'sldId') named.push(attr(el, 'id')!);
    for (const c of elements(el)) walk(c);
  };
  walk(root);
  if (named.length) expect(named).toEqual(presentationInfo(pkg).slides.map((s) => s.sldId));
}

/** The title text of each slide, in order. */
function titles(pkg: PptxPackage): string[] {
  return presentationInfo(pkg).slides.map((s) => {
    const tree = path(xmlPartOf(pkg, s.partName)?.root, 'p:cSld', 'p:spTree');
    const title = elements(tree).find((el) =>
      ['title', 'ctrTitle'].includes(attr(path(el, 'p:nvSpPr', 'p:nvPr', 'p:ph'), 'type') ?? '')
    );
    const out: string[] = [];
    const walk = (el: Element | undefined): void => {
      if (!el) return;
      if (el.local === 't') out.push(textOf(el));
      for (const c of elements(el)) walk(c);
    };
    walk(child(title, 'p:txBody'));
    return out.join('');
  });
}

const idOf = (e: PptxEditorInstance, name: string): string =>
  e.snapshot().sections!.find((s) => s.name === name)!.id;

/** A package from bytes. */
function read(bytes: Uint8Array): PptxPackage {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('read');
  return r.pkg;
}

/** The fixture with its presentation part rewritten by `edit`. */
function patched(edit: (xml: string) => string): Uint8Array {
  const f = unzipSync(buildSlideLayoutsDeck());
  f['ppt/presentation.xml'] = strToU8(edit(strFromU8(f['ppt/presentation.xml']!)));
  return zipSync(f);
}

/** The fixture without a section list; `other` is an extension that stays in `p:extLst`. */
function withoutSections(other = ''): Uint8Array {
  return patched((xml) =>
    xml.replace(
      /<p:extLst>.*<\/p:extLst>/,
      other
        ? `<p:extLst><p:ext uri="{EFAFB233-063F-42B5-8137-9DF3F51BA10A}">${other}</p:ext></p:extLst>`
        : ''
    )
  );
}

describe('sections in the snapshot', () => {
  test('lists each section with its first slide and slide count', async () => {
    const e = await open();
    expect(e.snapshot().sections!.map((s) => [s.name, s.firstSlide, s.slideCount])).toEqual([
      ['Opening', 0, 2],
      ['Details', 2, 4],
    ]);
    const before = e.snapshot().sections;
    e.exec({ type: 'setSlidesHidden', indices: [0], hidden: true });
    expect(e.snapshot().sections).toBe(before);
  });

  test('a deck without sections has none', async () => {
    const e = await open(buildPptx({ slides: [{ shapes: '' }] }));
    expect(e.snapshot().sections).toEqual([]);
  });
});

describe('add section', () => {
  test('splits the section at the slide', async () => {
    const e = await open();
    const before = canonicalPackageFingerprint(pkgOf(e));
    expect(e.exec({ type: 'addSection', name: 'Backup', index: 3 })).toEqual({ ok: true });
    for (const pkg of [pkgOf(e), reopened(e)]) {
      expect(sections(pkg)).toEqual(['Opening:0+2', 'Details:2+1', 'Backup:3+3']);
      expectConsistent(pkg);
    }
    expect(e.snapshot().currentSlide).toBe(3);
    e.exec({ type: 'undo' });
    expect(canonicalPackageFingerprint(pkgOf(e))).toBe(before);
  });

  test('at the first slide of a section, leaves that section empty', async () => {
    const e = await open();
    e.exec({ type: 'addSection', name: 'More', index: 2 });
    expect(sections(reopened(e))).toEqual(['Opening:0+2', 'Details:2+0', 'More:2+4']);
  });

  test('in a deck without sections, the slides before go into a first section', async () => {
    const deck = buildPptx({ slides: [{ shapes: '' }, { shapes: '' }, { shapes: '' }] });
    let e = await open(deck);
    e.goToSlide(1);
    e.exec({ type: 'addSection', name: 'Untitled Section', defaultSectionName: 'Default Section' });
    let pkg = reopened(e);
    expect(sections(pkg)).toEqual(['Default Section:0+1', 'Untitled Section:1+2']);
    expectConsistent(pkg);
    // Section ids are GUIDs in braces.
    for (const s of deckSections(pkg)) expect(s.id).toMatch(/^\{[0-9A-F-]{36}\}$/);
    e = await open(deck);
    e.exec({ type: 'addSection', name: 'All', index: 0 });
    pkg = reopened(e);
    expect(sections(pkg)).toEqual(['All:0+3']);
    // New slides keep joining the section list.
    e.exec({ type: 'addSlide' });
    expect(sections(pkgOf(e))).toEqual(['All:0+4']);
    expectConsistent(pkgOf(e));
  });
});

describe('section ids', () => {
  test('are GUIDs in braces, also without crypto.randomUUID (an insecure context)', async () => {
    const e = await open();
    const crypto = globalThis.crypto as { randomUUID?: unknown };
    const saved = crypto.randomUUID;
    try {
      Object.defineProperty(crypto, 'randomUUID', { value: undefined, configurable: true });
      expect(e.exec({ type: 'addSection', name: 'Late', index: 4 }).ok).toBe(true);
    } finally {
      Object.defineProperty(crypto, 'randomUUID', { value: saved, configurable: true });
    }
    const id = idOf(e, 'Late');
    expect(id).toMatch(/^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/);
  });
});

describe('rename section', () => {
  test('renames, and refuses an empty name', async () => {
    const e = await open();
    e.exec({
      type: 'renameSection',
      sectionId: idOf(e, 'Details'),
      name: '  The <details> & more ',
    });
    expect(sections(reopened(e))[1]).toBe('The <details> & more:2+4');
    expect(e.can({ type: 'renameSection', sectionId: idOf(e, 'Opening'), name: ' \n ' })).toEqual({
      ok: false,
      reason: 'errors.sectionName',
    });
    expect(e.can({ type: 'renameSection', sectionId: '{nope}', name: 'x' }).ok).toBe(false);
  });
});

describe('remove section', () => {
  test('its slides join the section before it, or after it for the first', async () => {
    let e = await open();
    e.exec({ type: 'removeSection', sectionId: idOf(e, 'Details') });
    let pkg = reopened(e);
    expect(sections(pkg)).toEqual(['Opening:0+6']);
    expectConsistent(pkg);
    e = await open();
    e.exec({ type: 'removeSection', sectionId: idOf(e, 'Opening') });
    pkg = reopened(e);
    expect(sections(pkg)).toEqual(['Details:0+6']);
    expectConsistent(pkg);
    // The last section removes the list; the slides stay.
    e.exec({ type: 'removeSection', sectionId: idOf(e, 'Details') });
    pkg = reopened(e);
    expect(sections(pkg)).toEqual([]);
    expect(presentationInfo(pkg).slides.length).toBe(6);
  });

  test('with its slides deletes them, in one undo step', async () => {
    const e = await open();
    const before = canonicalPackageFingerprint(pkgOf(e));
    e.goToSlide(4);
    e.exec({ type: 'removeSection', sectionId: idOf(e, 'Details'), withSlides: true });
    const pkg = reopened(e);
    expect(titles(pkg)).toEqual(['Quarterly review', 'Agenda']);
    expect(sections(pkg)).toEqual(['Opening:0+2']);
    expect(e.snapshot().currentSlide).toBe(1);
    e.exec({ type: 'undo' });
    expect(canonicalPackageFingerprint(pkgOf(e))).toBe(before);
  });

  test('refuses to delete every slide', async () => {
    const e = await open();
    e.exec({ type: 'removeSection', sectionId: idOf(e, 'Opening') });
    expect(
      e.can({ type: 'removeSection', sectionId: idOf(e, 'Details'), withSlides: true })
    ).toEqual({ ok: false, reason: 'disabled.lastSlide' });
  });
});

describe('move section', () => {
  test('moves a section and its slides past its neighbour', async () => {
    const e = await open();
    const before = canonicalPackageFingerprint(pkgOf(e));
    expect(e.exec({ type: 'moveSection', sectionId: idOf(e, 'Details'), direction: -1 }).ok).toBe(
      true
    );
    const pkg = reopened(e);
    expect(sections(pkg)).toEqual(['Details:0+4', 'Opening:4+2']);
    expect(titles(pkg)).toEqual([
      'Before and after',
      'Backup slide',
      '',
      'In their words',
      'Quarterly review',
      'Agenda',
    ]);
    expectConsistent(pkg);
    expect(e.snapshot().currentSlide).toBe(0);
    e.exec({ type: 'undo' });
    expect(canonicalPackageFingerprint(pkgOf(e))).toBe(before);
    e.exec({ type: 'moveSection', sectionId: idOf(e, 'Opening'), direction: 1 });
    expect(sections(pkgOf(e))).toEqual(['Details:0+4', 'Opening:4+2']);
  });

  test('refuses a direction other than one step up or down', async () => {
    const e = await open();
    const id = idOf(e, 'Opening');
    for (const direction of [0, 2, -2]) {
      const r = e.exec({ type: 'moveSection', sectionId: id, direction: direction as 1 });
      expect(r).toMatchObject({ ok: false, reason: 'errors.unsupported' });
    }
    expect(sections(pkgOf(e))).toEqual(['Opening:0+2', 'Details:2+4']);
  });

  test('refuses to move past the first or the last section', async () => {
    const e = await open();
    expect(e.can({ type: 'moveSection', sectionId: idOf(e, 'Opening'), direction: -1 })).toEqual({
      ok: false,
      reason: 'disabled.sectionEdge',
    });
    expect(e.can({ type: 'moveSection', sectionId: idOf(e, 'Details'), direction: 1 }).ok).toBe(
      false
    );
  });
});

describe('presentation view', () => {
  test('Slide Sorter is a view setting: no edit, no undo step, never saved', async () => {
    const e = await open();
    const before = canonicalPackageFingerprint(pkgOf(e));
    expect(e.snapshot().presentationView).toBe('normal');
    expect(e.exec({ type: 'setPresentationView', view: 'slideSorter' })).toEqual({ ok: true });
    expect(e.snapshot().presentationView).toBe('slideSorter');
    expect(e.snapshot().canUndo).toBe(false);
    expect(canonicalPackageFingerprint(pkgOf(e))).toBe(before);
    // Viewing mode can switch views too: nothing is edited.
    e.setMode('view');
    expect(e.can({ type: 'setPresentationView', view: 'normal' })).toEqual({ ok: true });
    expect(e.can({ type: 'setPresentationView', view: 'reading' as unknown as 'normal' })).toEqual({
      ok: false,
      reason: 'errors.view',
    });
  });
});

describe('sections in an unusual file', () => {
  test(
    'a slide id named twice counts once; a huge repeated list does not crash',
    { timeout: 30_000 },
    async () => {
      const twice = patched((xml) =>
        xml.replace('<p14:sldId id="258"/>', '<p14:sldId id="256"/><p14:sldId id="258"/>')
      );
      expect(sections(read(twice))).toEqual(['Opening:0+2', 'Details:2+4']);
      const many = patched((xml) =>
        xml.replace('<p14:sldId id="256"/>', '<p14:sldId id="256"/>'.repeat(60_000))
      );
      const e = await open(many);
      expect(e.snapshot().sections!.map((x) => x.slideCount)).toEqual([2, 4]);
    }
  );

  test('two section lists read as one, and the writer joins them into one extLst', async () => {
    const second =
      '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">' +
      '<p14:section name="Late" id="{11111111-2222-4333-8444-555555555555}"><p14:sldIdLst><p14:sldId id="261"/><p14:sldId id="260"/></p14:sldIdLst></p14:section>' +
      '</p14:sectionLst></p:ext></p:extLst>';
    const e = await open(
      patched((xml) => xml.replace('</p:presentation>', `${second}</p:presentation>`))
    );
    expect(e.snapshot().sections!.map((x) => x.name)).toEqual(['Opening', 'Details', 'Late']);
    const pkg = reopened(e);
    const root = xmlPartOf(pkg, presentationInfo(pkg).presentationPart)!.root;
    expect(elements(root).filter((el) => el.local === 'extLst').length).toBe(1);
    // Each slide is in one section: Late keeps no slide that Details names.
    expect(sections(pkg)).toEqual(['Opening:0+2', 'Details:2+4', 'Late:6+0']);
    expectConsistent(pkg);
  });

  test('add section keeps other extensions; the last section leaves them', async () => {
    const other = '<p:custData/>';
    const e = await open(withoutSections(other));
    e.exec({ type: 'addSection', index: 2 });
    let pkg = reopened(e);
    expect(sections(pkg)).toEqual(['Default Section:0+2', 'Untitled Section:2+4']);
    const exts = (p: PptxPackage) =>
      elements(child(xmlPartOf(p, presentationInfo(p).presentationPart)!.root, 'p:extLst')).map(
        (x) => attr(x, 'uri')
      );
    expect(exts(pkg)).toEqual([
      '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}',
      '{EFAFB233-063F-42B5-8137-9DF3F51BA10A}',
    ]);
    e.exec({ type: 'removeAllSections' });
    pkg = reopened(e);
    expect(sections(pkg)).toEqual([]);
    expect(exts(pkg)).toEqual(['{EFAFB233-063F-42B5-8137-9DF3F51BA10A}']);
  });

  test('add section on a slide that no section names takes that slide', async () => {
    const e = await open(
      patched((xml) => xml.replace('<p14:sldId id="260"/><p14:sldId id="261"/>', ''))
    );
    // The rail shows the sections as the file will hold them: the slides join Details.
    expect(e.snapshot().sections!.map((x) => x.slideCount)).toEqual([2, 4]);
    e.exec({ type: 'addSection', name: 'Found', index: 4 });
    const pkg = reopened(e);
    expect(sections(pkg)).toEqual(['Opening:0+2', 'Details:2+2', 'Found:4+2']);
    expectConsistent(pkg);
  });
});

describe('raw slide ids', () => {
  // "0260" in `p:sldIdLst` and "260" in a section name one slide; a new section list writes
  // the slide's own value, so an edit never rewrites "0260" as "260".
  const zeroIds = (xml: string): string =>
    xml
      .replace('<p:sldId id="260"', '<p:sldId id="0260"')
      .replace('<p:sldId id="261"', '<p:sldId id="0261"');

  test("add section inside a section writes the slides' raw ids, in slide order", async () => {
    const e = await open(patched(zeroIds));
    e.exec({ type: 'addSection', name: 'Backup', index: 4 });
    for (const pkg of [pkgOf(e), reopened(e)]) {
      expect(sections(pkg)).toEqual(['Opening:0+2', 'Details:2+2', 'Backup:4+2']);
      expectConsistent(pkg);
    }
  });

  test('add section on unlisted slides writes their raw ids', async () => {
    const e = await open(
      patched((xml) => zeroIds(xml).replace('<p14:sldId id="260"/><p14:sldId id="261"/>', ''))
    );
    e.exec({ type: 'addSection', name: 'Found', index: 4 });
    const pkg = reopened(e);
    expect(sections(pkg)).toEqual(['Opening:0+2', 'Details:2+2', 'Found:4+2']);
    expectConsistent(pkg);
  });
});

describe('section defaults and refusals', () => {
  test('names come from the host translator, else English', async () => {
    let e = await open(withoutSections());
    e.exec({ type: 'addSection', index: 3 });
    expect(sections(pkgOf(e))).toEqual(['Default Section:0+3', 'Untitled Section:3+3']);
    e = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      translate: (key) =>
        ({ 'sections.untitled': 'Neu', 'sections.default': 'Standard' })[key] ?? key,
    });
    expect((await e.load(withoutSections())).ok).toBe(true);
    e.exec({ type: 'addSection', index: 3 });
    expect(sections(pkgOf(e))).toEqual(['Standard:0+3', 'Neu:3+3']);
  });

  test('unknown sections and decks without sections refuse with their own reason', async () => {
    const e = await open(withoutSections());
    expect(e.can({ type: 'renameSection', sectionId: '{nope}', name: 'x' })).toEqual({
      ok: false,
      reason: 'errors.sectionNotFound',
    });
    expect(e.can({ type: 'removeAllSections' })).toEqual({
      ok: false,
      reason: 'disabled.noSections',
    });
    expect(e.can({ type: 'addSection', index: 9 })).toEqual({
      ok: false,
      reason: 'errors.slideNotFound',
    });
  });

  test('remove section and slides keeps the current slide when it survives', async () => {
    const e = await open();
    e.goToSlide(4);
    e.exec({ type: 'removeSection', sectionId: idOf(e, 'Opening'), withSlides: true });
    // Old slide 5 is slide 3 now, and it is still the current slide.
    expect(e.snapshot().currentSlide).toBe(2);
    expect(titles(pkgOf(e))[2]).toBe(titles(read(buildSlideLayoutsDeck()))[4]);
  });

  test('moving next to an empty section keeps every slide in place', async () => {
    const e = await open();
    e.exec({ type: 'addSection', name: 'Empty', index: 2 });
    expect(sections(pkgOf(e))).toEqual(['Opening:0+2', 'Details:2+0', 'Empty:2+4']);
    e.exec({ type: 'moveSection', sectionId: idOf(e, 'Details'), direction: 1 });
    const pkg = reopened(e);
    expect(sections(pkg)).toEqual(['Opening:0+2', 'Empty:2+4', 'Details:6+0']);
    expect(titles(pkg)[2]).toBe('Before and after');
    expectConsistent(pkg);
  });
});

describe('collapsed sections and the Section menu', () => {
  test('collapse is a view setting: shared, no undo step, never saved, reset on load', async () => {
    const e = await open();
    const before = canonicalPackageFingerprint(pkgOf(e));
    const [opening, details] = e.snapshot().sections!.map((x) => x.id);
    e.exec({ type: 'setSectionsCollapsed', collapsed: true, sectionIds: [details!] });
    expect(e.snapshot().collapsedSections).toEqual([details]);
    e.exec({ type: 'setSectionsCollapsed', collapsed: true });
    expect([...e.snapshot().collapsedSections!].sort()).toEqual([opening, details].sort());
    expect(e.snapshot().canUndo).toBe(false);
    expect(canonicalPackageFingerprint(pkgOf(e))).toBe(before);
    e.exec({ type: 'setPresentationView', view: 'slideSorter' });
    await e.load(buildSlideLayoutsDeck());
    expect(e.snapshot().collapsedSections).toEqual([]);
    expect(e.snapshot().presentationView).toBe('normal');
  });

  test('a raw run of section.rename names requestSectionRename and warns', async () => {
    const e = await open();
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => undefined);
    try {
      const r = runToolbarCommand(e, 'section.rename');
      expect(r).toMatchObject({
        ok: false,
        reason: expect.stringContaining('requestSectionRename'),
      });
      expect(String(warn.mock.calls[0]?.[0])).toContain('runToolbarCommand("section.rename")');
      // With a value (a name), the slot still refuses with the same reason.
      expect(runToolbarCommand(e, 'section.rename', 'Name')).toEqual(r);
    } finally {
      warn.mockRestore();
    }
  });

  test('the section.* slots run on the current slide’s section, or on all of them', async () => {
    const e = await open();
    e.goToSlide(3);
    const run = (slot: ChromeSlotId) => e.exec(commandForSlot(slot)!);
    expect(toolbarCommandState(e, 'section.add').enabled).toBe(true);
    expect(toolbarCommandState(e, 'section.rename').enabled).toBe(true);
    run('section.add');
    expect(sections(pkgOf(e))).toEqual(['Opening:0+2', 'Details:2+1', 'Untitled Section:3+3']);
    run('section.remove');
    expect(sections(pkgOf(e))).toEqual(['Opening:0+2', 'Details:2+4']);
    run('section.collapseAll');
    expect(e.snapshot().collapsedSections!.length).toBe(2);
    run('section.expandAll');
    expect(e.snapshot().collapsedSections).toEqual([]);
    run('section.removeAll');
    expect(sections(pkgOf(e))).toEqual([]);
    // Without sections, every row but Add says why it is off.
    for (const slot of ['section.remove', 'section.collapseAll', 'section.rename'] as const)
      expect(toolbarCommandState(e, slot).disabledReason).toBe('disabled.noSections');
  });

  test('in the slide sorter, undo, Find, and Select All bring back no shape selection', async () => {
    const e = await open();
    const title = e.session!.layout(0)!.items.find((i) => i.kind === 'shape')!.element.nodeId;
    e.setSelection({ kind: 'shapes', nodeIds: [title] });
    expect(e.exec({ type: 'toggleBold' }).ok).toBe(true);
    e.exec({ type: 'setPresentationView', view: 'slideSorter' });
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(e.snapshot().selection).toEqual({ kind: 'none' });
    expect(toolbarCommandState(e, 'text.bold').enabled).toBe(false);
    expect(e.exec({ type: 'delete' }).ok).toBe(false);
    e.setSelection({ kind: 'shapes', nodeIds: [title] });
    expect(e.snapshot().selection).toEqual({ kind: 'none' });
    e.find.next('e');
    expect(e.snapshot().selection).toEqual({ kind: 'none' });
    // Select All selects every slide there, as the sorter's own Cmd+A.
    e.selectAll();
    expect(e.snapshot().selectedSlides.length).toBe(e.snapshot().slides.length);
    expect(e.snapshot().selection).toEqual({ kind: 'none' });
  });

  test('in the slide sorter, commands that add a shape to the slide wait for Normal', async () => {
    const e = await open();
    e.exec({ type: 'setPresentationView', view: 'slideSorter' });
    expect(e.can({ type: 'insertTextBox' })).toEqual({
      ok: false,
      reason: 'disabled.slideSorter',
    });
    expect(e.exec({ type: 'insertTextBox' }).ok).toBe(false);
    expect(toolbarCommandState(e, 'insert.textBox').enabled).toBe(false);
    // Insert > Table is refused too, until the Normal view.
    expect(e.exec({ type: 'insertTable', rows: 2, cols: 2 }).ok).toBe(false);
    expect(toolbarCommandState(e, 'table.insert')).toMatchObject({
      enabled: false,
      disabledReason: 'disabled.slideSorter',
    });
    expect(e.snapshot().selection).toEqual({ kind: 'none' });
    // Slide commands still run.
    expect(e.exec({ type: 'addSlide' }).ok).toBe(true);
    e.exec({ type: 'setPresentationView', view: 'normal' });
    expect(e.can({ type: 'insertTextBox' }).ok).toBe(true);
  });

  test('switching to the slide sorter leaves no shape selected', async () => {
    const e = await open();
    const title = e.session!.layout(0)!.items[0]!.element.nodeId;
    e.setSelection({ kind: 'shapes', nodeIds: [title] });
    e.exec({ type: 'setPresentationView', view: 'slideSorter' });
    expect(e.snapshot().selection).toEqual({ kind: 'none' });
  });
});

describe('the writer keeps the section list whole', () => {
  test('a section entry for a gone slide goes; a slide in no section joins the one before it', async () => {
    const e = await open(
      patched((xml) => xml.replace('<p14:sldId id="258"/>', '<p14:sldId id="999"/>'))
    );
    // Slide 258 (index 2) is in no section, and Details names a slide the deck lacks.
    const pkg = reopened(e);
    expect(sections(pkg)).toEqual(['Opening:0+3', 'Details:3+3']);
    expectConsistent(pkg);
    const xml = JSON.stringify(xmlPartOf(pkg, presentationInfo(pkg).presentationPart)!.root);
    expect(xml).not.toContain('"999"');
  });

  test('an empty section extension first keeps the sections of the full one', async () => {
    const e = await open(
      patched((xml) =>
        xml.replace('<p:extLst>', '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"/>')
      )
    );
    expect(sections(reopened(e))).toEqual(['Opening:0+2', 'Details:2+4']);
  });

  test('a section id named twice is written once; its slides stay in a section', async () => {
    const e = await open(
      patched((xml) => xml.replace(/id="\{2D4E6F80[^"]*\}"/, `id="${LAYOUT_SECTION_IDS[0]}"`))
    );
    const pkg = reopened(e);
    expect(deckSections(pkg).length).toBe(1);
    expect(sections(pkg)).toEqual(['Opening:0+6']);
    expectConsistent(pkg);
  });

  test('a slide id whose slide part is gone leaves the slide list and its section', async () => {
    // A slide that one collaborator deleted while another moved it comes back this way.
    const e = await open(
      patched((xml) =>
        xml
          .replace('</p:sldIdLst>', '<p:sldId id="300" r:id="rId999"/></p:sldIdLst>')
          .replace('<p14:sldId id="261"/>', '<p14:sldId id="261"/><p14:sldId id="300"/>')
      )
    );
    expect(e.snapshot().slides.length).toBe(6);
    const pkg = reopened(e);
    const xml = JSON.stringify(xmlPartOf(pkg, presentationInfo(pkg).presentationPart)!.root);
    expect(xml).not.toContain('rId999');
    expect(xml).not.toContain('"300"');
    expect(sections(pkg)).toEqual(['Opening:0+2', 'Details:2+4']);
    expectConsistent(pkg);
  });

  test('sections listed out of slide order follow the slides, once and for all', async () => {
    const opening =
      '<p14:section name="Opening" id="{9A3C1F52-7B2E-4D61-8E0A-3C5B7D9E1F23}"><p14:sldIdLst><p14:sldId id="256"/><p14:sldId id="257"/></p14:sldIdLst></p14:section>';
    const e = await open(
      patched((xml) =>
        xml
          .replace(opening, '')
          .replace('</p14:sectionLst>', `${opening}</p14:sectionLst>`)
          .replace('<p:sldId id="261" r:id="rId106"/>', '')
          .replace('<p:sldIdLst>', '<p:sldIdLst><p:sldId id="261" r:id="rId106"/>')
      )
    );
    // Slide 261 (Details) is first, then Opening's slides, then Details' others: Details keeps
    // its first slide, and the slides after Opening's stay in Opening.
    const pkg = reopened(e);
    expect(sections(pkg)).toEqual(['Details:0+1', 'Opening:1+5']);
    expectConsistent(pkg);
    // The repaired file needs no second repair.
    const again = reopened(await open(e.save()!));
    expect(canonicalPackageFingerprint(again)).toBe(canonicalPackageFingerprint(pkg));
    // Listed in reverse, with the slides in place: the sections take the slides' order.
    const reverse = await open(
      patched((xml) =>
        xml.replace(opening, '').replace('</p14:sectionLst>', `${opening}</p14:sectionLst>`)
      )
    );
    expect(sections(reopened(reverse))).toEqual(['Opening:0+2', 'Details:2+4']);
  });

  test('a slide whose target differs only in case from its part name stays in the file', async () => {
    const files = unzipSync(buildSlideLayoutsDeck());
    const rels = strFromU8(files['ppt/_rels/presentation.xml.rels']!);
    files['ppt/_rels/presentation.xml.rels'] = strToU8(
      rels.replace('Target="slides/slide3.xml"', 'Target="slides/Slide3.xml"')
    );
    const e = await open(zipSync(files));
    const xml = strFromU8(unzipSync(e.save()!)['ppt/presentation.xml']!);
    expect(xml).toContain('<p:sldId id="258"');
    expect(xml).toContain('<p14:sldId id="258"/>');
  });

  test('the rail shows the sections as they will be saved, and commands act on that', async () => {
    const opening =
      '<p14:section name="Opening" id="{9A3C1F52-7B2E-4D61-8E0A-3C5B7D9E1F23}"><p14:sldIdLst><p14:sldId id="256"/><p14:sldId id="257"/></p14:sldIdLst></p14:section>';
    const e = await open(
      patched((xml) =>
        xml
          .replace(opening, '')
          .replace('</p14:sectionLst>', `${opening}</p14:sectionLst>`)
          .replace('<p:sldId id="261" r:id="rId106"/>', '')
          .replace('<p:sldIdLst>', '<p:sldIdLst><p:sldId id="261" r:id="rId106"/>')
      )
    );
    const shown = e.snapshot().sections!.map((x) => `${x.name}:${x.firstSlide}+${x.slideCount}`);
    expect(shown).toEqual(['Details:0+1', 'Opening:1+5']);
    // Remove Section and slides on Details removes the one slide that the rail shows in it.
    expect(
      e.exec({ type: 'removeSection', sectionId: idOf(e, 'Details'), withSlides: true }).ok
    ).toBe(true);
    expect(e.snapshot().slides.length).toBe(5);
    expectConsistent(pkgOf(e));
    expect(sections(reopened(e))).toEqual(['Opening:0+5']);
  });

  test('a slide command drops a slide entry that names no slide from the part', async () => {
    const e = await open(
      patched((xml) => xml.replace('<p:sldIdLst>', '<p:sldIdLst><p:sldId id="900" r:id="rId900"/>'))
    );
    expect(e.exec({ type: 'addSlide' }).ok).toBe(true);
    const root = xmlPartOf(pkgOf(e), presentationInfo(pkgOf(e)).presentationPart)!.root;
    expect(JSON.stringify(root)).not.toContain('rId900');
    expectConsistent(pkgOf(e));
    // Undo brings the part back as it was.
    e.exec({ type: 'undo' });
    expect(
      JSON.stringify(xmlPartOf(pkgOf(e), presentationInfo(pkgOf(e)).presentationPart)!.root)
    ).toContain('rId900');
  });

  test('a dead entry that reuses a live slide id goes, and the live one stays', async () => {
    const e = await open(
      patched((xml) => xml.replace('<p:sldIdLst>', '<p:sldIdLst><p:sldId id="258" r:id="rId999"/>'))
    );
    expect(e.exec({ type: 'addSection', name: 'Late', index: 4 }).ok).toBe(true);
    const xml = strFromU8(unzipSync(e.save()!)['ppt/presentation.xml']!);
    expect(xml).not.toContain('rId999');
    expect(xml).toContain('<p:sldId id="258" r:id="rId103"/>');
    expect(e.snapshot().slides.length).toBe(6);
  });

  test('a deck without extension lists repairs its slide entries and keeps working', async () => {
    const e = await open(
      patched((xml) =>
        xml
          .replace(/<p:extLst>.*<\/p:extLst>/, '')
          .replace('<p:sldIdLst>', '<p:sldIdLst><p:sldId id="900" r:id="rId900"/>')
      )
    );
    expect(e.exec({ type: 'addSlide' }).ok).toBe(true);
    expect(e.exec({ type: 'addSection', name: 'A', index: 1 }).ok).toBe(true);
    const pkg = pkgOf(e);
    const root = xmlPartOf(pkg, presentationInfo(pkg).presentationPart)!.root;
    expect(JSON.stringify(root)).not.toContain('rId900');
    expect(repairedPresentationRoot(pkg)).toBe(root);
  });

  test('a section without a slide list gets one on save', async () => {
    const e = await open(
      patched((xml) =>
        xml.replace(
          '</p14:sectionLst>',
          '<p14:section name="Empty" id="{11111111-2222-4333-8444-555555555555}"/></p14:sectionLst>'
        )
      )
    );
    const xml = strFromU8(unzipSync(e.save()!)['ppt/presentation.xml']!);
    expect(xml).toMatch(
      /name="Empty" id="\{11111111-2222-4333-8444-555555555555\}"><p14:sldIdLst\/?>/
    );
  });

  test('a slide entry whose relationship has no part goes on save', async () => {
    const files = unzipSync(buildSlideLayoutsDeck());
    delete files['ppt/slides/slide3.xml'];
    const e = await open(zipSync(files));
    expect(e.snapshot().slides.length).toBe(5);
    const xml = strFromU8(unzipSync(e.save()!)['ppt/presentation.xml']!);
    expect(xml).not.toContain('<p:sldId id="258"');
    expect(xml).not.toContain('<p14:sldId id="258"/>');
  });

  test('after a slide or section command the part needs no repair', async () => {
    const opening =
      '<p14:section name="Opening" id="{9A3C1F52-7B2E-4D61-8E0A-3C5B7D9E1F23}"><p14:sldIdLst><p14:sldId id="256"/><p14:sldId id="257"/></p14:sldIdLst></p14:section>';
    const e = await open(
      patched((xml) =>
        xml
          .replace(opening, '')
          .replace('</p14:sectionLst>', `${opening}</p14:sectionLst>`)
          .replace('<p:sldIdLst>', '<p:sldIdLst><p:sldId id="900" r:id="rId900"/>')
          .replace(
            '</p:extLst>',
            '</p:extLst><p:extLst><p:ext uri="{EFAFB233-063F-42B5-8137-9DF3F51BA10A}"/></p:extLst>'
          )
      )
    );
    const before = e.snapshot().sections;
    expect(e.exec({ type: 'renameSection', sectionId: idOf(e, 'Details'), name: 'More' }).ok).toBe(
      true
    );
    const pkg = pkgOf(e);
    const root = xmlPartOf(pkg, presentationInfo(pkg).presentationPart)!.root;
    expect(repairedPresentationRoot(pkg)).toBe(root);
    expect(JSON.stringify(root)).toContain('EFAFB233');
    expect(e.snapshot().sections!.map((s) => [s.firstSlide, s.slideCount])).toEqual(
      before!.map((s) => [s.firstSlide, s.slideCount])
    );
  });

  test('a section-list extension without a list goes, once', async () => {
    const e = await open(
      patched((xml) =>
        xml.replace(
          '</p:extLst>',
          '<p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"/></p:extLst>'
        )
      )
    );
    const once = e.save()!;
    expect(strFromU8(unzipSync(once)['ppt/presentation.xml']!).match(/521415D9/g)!.length).toBe(1);
    const again = reopened(await open(once));
    expect(sections(again)).toEqual(['Opening:0+2', 'Details:2+4']);
  });

  test('two extensions of another kind with one URI both stay', async () => {
    const guides = (n: string) =>
      `<p:ext uri="{EFAFB233-063F-42B5-8137-9DF3F51BA10A}"><p15:sldGuideLst xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"><p15:guide id="${n}" orient="horz" pos="2160"/></p15:sldGuideLst></p:ext>`;
    const e = await open(
      patched((xml) =>
        xml.replace(
          '</p:extLst>',
          `${guides('1')}${guides('2')}</p:extLst><p:extLst>${guides('3')}</p:extLst>`
        )
      )
    );
    const pkg = reopened(e);
    const root = xmlPartOf(pkg, presentationInfo(pkg).presentationPart)!.root;
    expect(root.children.filter((c) => c.type === 'element' && c.local === 'extLst').length).toBe(
      1
    );
    const xml = JSON.stringify(root);
    for (const n of ['1', '2', '3']) expect(xml).toContain(`"value":"${n}"`);
    expect(sections(pkg)).toEqual(['Opening:0+2', 'Details:2+4']);
  });

  test('add section splits by slide order, not by the order the list names the slides', async () => {
    const e = await open(
      patched((xml) =>
        xml.replace(
          '<p14:sldId id="258"/><p14:sldId id="259"/><p14:sldId id="260"/><p14:sldId id="261"/>',
          '<p14:sldId id="261"/><p14:sldId id="260"/><p14:sldId id="259"/><p14:sldId id="258"/>'
        )
      )
    );
    e.exec({ type: 'addSection', name: 'Late', index: 4 });
    const pkg = reopened(e);
    expect(sections(pkg)).toEqual(['Opening:0+2', 'Details:2+2', 'Late:4+2']);
    expectConsistent(pkg);
  });

  test('the snapshot reports only collapsed sections that the deck still has', async () => {
    const e = await open();
    const details = idOf(e, 'Details');
    expect(e.exec({ type: 'setSectionsCollapsed', collapsed: true }).ok).toBe(true);
    expect(e.snapshot().collapsedSections!.length).toBe(2);
    expect(e.exec({ type: 'removeSection', sectionId: details }).ok).toBe(true);
    const shown = e.snapshot().collapsedSections!;
    expect(shown).toEqual([idOf(e, 'Opening')]);
    // Reference-stable while it does not change.
    e.goToSlide(3);
    expect(e.snapshot().collapsedSections).toBe(shown);
  });
});
