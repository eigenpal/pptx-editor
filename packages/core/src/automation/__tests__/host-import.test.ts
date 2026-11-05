import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import {
  packageIntegrity,
  relationshipReferences,
} from '../../../../../scripts/fixtures/package-integrity.ts';
import {
  attr,
  canonicalPackageFingerprint,
  child,
  contentTypeOf,
  elements,
  path,
  presentationInfo,
  readPptx,
  relationshipsOf,
  shapeTreeOf,
  textOf,
  xmlPartOf,
  type PptxPackage,
} from '../../store/index.ts';
import { toBase64 } from '../commands/base64.ts';
import { REFUSED } from '../commands/slide-import.ts';
import { P15, treeOf } from '../commands/slide-structure.ts';
import {
  createServerAutomationHost,
  type AutomationBatchResponse,
  type AutomationHandle,
  type AutomationHost,
  type AutomationInsertSlidesOptions,
  type AutomationOperation,
  type AutomationValue,
} from '../index.ts';

const IN = 914_400;
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument.presentationml';
/** A 1 by 1 PNG. */
const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
  ),
  (c) => c.charCodeAt(0)
);

/** Rewrite named files of a zip. */
function patch(bytes: Uint8Array, edits: Record<string, (text: string) => string>): Uint8Array {
  const files = unzipSync(bytes);
  for (const [name, change] of Object.entries(edits))
    files[name] = strToU8(change(files[name] ? strFromU8(files[name]) : ''));
  return zipSync(files);
}

const title = (text: string) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${para(run(text))}</p:txBody></p:sp>`;
const body = (id: number, idx: number, text: string) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Body ${id}"/><p:cNvSpPr/><p:nvPr><p:ph idx="${idx}"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${para(run(text))}</p:txBody></p:sp>`;
const picture = (id: number) =>
  `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${IN}" cy="${IN}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
const jump = (id: number) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Jump"><a:hlinkClick r:id="rIdJump" action="ppaction://hlinksldjump"/></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${IN}" cy="${IN}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>`;

/**
 * A source deck with two masters. Slide 1 ("Alpha") is on Title and Content with a picture,
 * notes, and a link to slide 2. Slide 2 ("Beta") is on a section header layout. Slide 3
 * ("Gamma") is on the second master's only layout (`twoObj`).
 */
function sourceDeck(): Uint8Array {
  const bytes = buildPptx({
    colors: {
      dk1: '000000',
      lt1: 'FFFFFF',
      dk2: '112233',
      lt2: 'EEEEEE',
      accent1: 'AA0000',
      accent2: '00AA00',
      accent3: '0000AA',
      accent4: 'AAAA00',
      accent5: '00AAAA',
      accent6: 'AA00AA',
      hlink: '0000FF',
      folHlink: '800080',
    },
    layouts: [{ name: 'Section Header', type: 'secHead', shapes: title('Section') }],
    media: [{ name: 'image1.png', bytes: PNG, contentType: 'image/png' }],
    slides: [
      {
        layout: 'title',
        shapes: title('Alpha') + body(3, 1, 'Alpha body') + picture(4) + jump(5),
        notes: 'Alpha notes',
        rels: [
          ['rIdImg', `${R}/image`, '../media/image1.png'],
          ['rIdJump', `${R}/slide`, 'slide2.xml'],
          ['rIdWeb', `${R}/hyperlink`, 'https://example.com/', 'External'],
        ],
      },
      { layoutName: 'Section Header', shapes: title('Beta') },
      { layout: 'blank', shapes: title('Gamma') + body(3, 2, 'Gamma body') },
    ],
  });
  return withSecondMaster(bytes, 3);
}

/** Add a second master with one `twoObj` layout and its own theme, and put slide `n` on it. */
function withSecondMaster(bytes: Uint8Array, n: number): Uint8Array {
  const files = unzipSync(bytes);
  const text = (name: string) => strFromU8(files[name]!);
  const master = text('ppt/slideMasters/slideMaster1.xml').replace(
    /<p:sldLayoutIdLst>.*?<\/p:sldLayoutIdLst>/,
    '<p:sldLayoutIdLst><p:sldLayoutId id="2147483700" r:id="rId1"/></p:sldLayoutIdLst>'
  );
  const layout = text('ppt/slideLayouts/slideLayout1.xml')
    .replace('type="obj"', 'type="twoObj"')
    .replace('name="Title and Content"', 'name="Two Content"');
  const theme = text('ppt/theme/theme1.xml').replace(
    /<a:theme ([^>]*)name="[^"]*"/,
    '<a:theme $1name="Second"'
  );
  const rels = (items: [string, string, string][]) =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items
      .map(
        ([id, type, target]) => `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`
      )
      .join('')}</Relationships>`;
  return patch(zipSync(files), {
    'ppt/slideMasters/slideMaster2.xml': () => master,
    'ppt/slideMasters/_rels/slideMaster2.xml.rels': () =>
      rels([
        ['rId1', 'slideLayout', '../slideLayouts/slideLayout9.xml'],
        ['rId2', 'theme', '../theme/theme3.xml'],
      ]),
    'ppt/slideLayouts/slideLayout9.xml': () => layout,
    'ppt/slideLayouts/_rels/slideLayout9.xml.rels': () =>
      rels([['rId1', 'slideMaster', '../slideMasters/slideMaster2.xml']]),
    'ppt/theme/theme3.xml': () => theme,
    'ppt/presentation.xml': (t) =>
      t.replace(
        '</p:sldMasterIdLst>',
        '<p:sldMasterId id="2147483690" r:id="rIdM2"/></p:sldMasterIdLst>'
      ),
    'ppt/_rels/presentation.xml.rels': (t) =>
      t.replace(
        '</Relationships>',
        `<Relationship Id="rIdM2" Type="${R}/slideMaster" Target="slideMasters/slideMaster2.xml"/></Relationships>`
      ),
    '[Content_Types].xml': (t) =>
      t.replace(
        '</Types>',
        `<Override PartName="/ppt/slideMasters/slideMaster2.xml" ContentType="${CT}.slideMaster+xml"/>` +
          `<Override PartName="/ppt/slideLayouts/slideLayout9.xml" ContentType="${CT}.slideLayout+xml"/>` +
          '<Override PartName="/ppt/theme/theme3.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>'
      ),
    [`ppt/slides/_rels/slide${n}.xml.rels`]: (t) =>
      t.replace('../slideLayouts/slideLayout2.xml', '../slideLayouts/slideLayout9.xml'),
  });
}

/** The destination: two slides and a section header layout. */
function destDeck(): Uint8Array {
  return buildPptx({
    layouts: [{ name: 'Section Header', type: 'secHead', shapes: title('Section') }],
    slides: [
      { layout: 'title', shapes: title('One') },
      { layout: 'blank', shapes: shape({ id: 2, name: 'Box', x: IN, y: IN, cx: IN, cy: IN }) },
    ],
  });
}

function open(bytes: Uint8Array): AutomationHost {
  const opened = createServerAutomationHost(bytes);
  if (!opened.ok) throw new Error(JSON.stringify(opened.error));
  return opened.host;
}

function exec(host: AutomationHost, operations: AutomationOperation[]): AutomationBatchResponse {
  return host.execute({ operations });
}

function ok(host: AutomationHost, operations: AutomationOperation[]): AutomationValue[] {
  const response = exec(host, operations);
  if (!response.ok) throw new Error(JSON.stringify(response.results));
  return response.results.map((r) => (r.status === 'ok' ? r.value : { kind: 'null' as const }));
}

function insert(host: AutomationHost, source: Uint8Array, options?: AutomationInsertSlidesOptions) {
  return ok(host, [
    { op: 'insertSlides', base64: toBase64(source), ...(options ? { options } : {}) },
  ]);
}

function refusal(host: AutomationHost, operations: AutomationOperation[]) {
  const revision = host.revision();
  const response = exec(host, operations);
  expect([response.ok, response.changed, host.revision()]).toEqual([false, false, revision]);
  const failed = response.results.find((r) => r.status === 'error');
  if (failed?.status !== 'error') throw new Error('expected a refusal');
  return { code: failed.error.code, detail: failed.error.detail };
}

function save(host: AutomationHost): Uint8Array {
  const saved = host.save();
  if (!saved.ok) throw new Error(saved.error.code);
  return saved.bytes;
}

function readBack(bytes: Uint8Array): PptxPackage {
  const read = readPptx(bytes);
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  return read.pkg;
}

/**
 * The saved package is whole (`scripts/fixtures/package-integrity.ts`), and a second save of the
 * reopened package has the same fingerprint.
 */
function expectWhole(bytes: Uint8Array): PptxPackage {
  const { issues } = packageIntegrity(bytes);
  expect(issues).toEqual([]);
  const pkg = readBack(bytes);
  const again = save(open(bytes));
  expect(canonicalPackageFingerprint(readBack(again))).toBe(canonicalPackageFingerprint(pkg));
  return pkg;
}

/** The text of each slide's shapes, in order. */
function slideTexts(pkg: PptxPackage): string[] {
  return presentationInfo(pkg).slides.map((s) =>
    elements(shapeTreeOf(pkg, s.partName))
      .map((el) => textOf(child(el, 'p:txBody')))
      .filter((t) => t)
      .join('|')
  );
}

const masterCount = (pkg: PptxPackage) =>
  elements(child(xmlPartOf(pkg, presentationInfo(pkg).presentationPart)!.root, 'p:sldMasterIdLst'))
    .length;

describe('insertSlides: keep source formatting', () => {
  test('every slide goes first, with its layout, master, theme, picture, and notes', () => {
    const host = open(destDeck());
    insert(host, sourceDeck());
    const pkg = expectWhole(save(host));
    const info = presentationInfo(pkg);
    expect(slideTexts(pkg)).toEqual(['Alpha|Alpha body', 'Beta', 'Gamma|Gamma body', 'One', '']);
    // The source's two masters come along; each lists only the layouts its slides use.
    expect(masterCount(pkg)).toBe(3);
    const [alpha, beta, gamma] = info.slides;
    expect(alpha!.masterPart).toBe(beta!.masterPart);
    expect(gamma!.masterPart).not.toBe(alpha!.masterPart);
    expect(
      elements(child(xmlPartOf(pkg, alpha!.masterPart)!.root, 'p:sldLayoutIdLst'))
    ).toHaveLength(2);
    expect(attr(xmlPartOf(pkg, gamma!.layoutPart)!.root, 'type')).toBe('twoObj');
    expect(attr(xmlPartOf(pkg, gamma!.themePart)!.root, 'name')).toBe('Second');
    expect(gamma!.themePart).not.toBe(info.slides[3]!.themePart);
    // The theme colors of the source travel with its master.
    const scheme = path(
      xmlPartOf(pkg, alpha!.themePart)!.root,
      'a:themeElements',
      'a:clrScheme',
      'a:accent1'
    );
    expect(attr(child(scheme, 'a:srgbClr'), 'val')).toBe('AA0000');
    // The picture is a byte-identical copy.
    const image = relationshipsOf(pkg, alpha!.partName).find((r) => r.id === 'rIdImg')!.target!;
    const part = pkg.parts.get(image)!;
    expect(part.kind === 'binary' ? Array.from(part.bytes) : []).toEqual(Array.from(PNG));
    // Notes come along, on this presentation's notes master, and name their new slide.
    expect(textOf(xmlPartOf(pkg, alpha!.notesPart)!.root)).toContain('Alpha notes');
    const notesRels = relationshipsOf(pkg, alpha!.notesPart!);
    expect(notesRels.find((r) => r.type === `${R}/slide`)?.target).toBe(alpha!.partName);
    expect(notesRels.find((r) => r.type === `${R}/notesMaster`)?.target).toBe(
      '/ppt/notesMasters/notesMaster1.xml'
    );
    // The link to slide 2 now names the copy of slide 2; the external link stays as it was.
    const rels = relationshipsOf(pkg, alpha!.partName);
    expect(rels.find((r) => r.id === 'rIdJump')?.target).toBe(beta!.partName);
    expect(rels.find((r) => r.id === 'rIdWeb')?.externalTarget).toBe('https://example.com/');
  });

  test('sourceSlideIds pick and order the slides; targetSlideId places them', () => {
    const host = open(destDeck());
    const [first] = ok(host, [
      { op: 'getSlide', key: { index: 0 } },
      { op: 'readSlide', slide: { result: 0 } },
    ]).slice(1);
    const target = first?.kind === 'slide' ? first.slide.id : '';
    insert(host, sourceDeck(), { sourceSlideIds: ['258', '256#'], targetSlideId: target });
    const pkg = expectWhole(save(host));
    expect(slideTexts(pkg)).toEqual(['One', 'Gamma|Gamma body', 'Alpha|Alpha body', '']);
    // Slide 2 was not imported, so the link to it loses its action and its relationship.
    const alpha = presentationInfo(pkg).slides[2]!;
    expect(relationshipsOf(pkg, alpha.partName).some((r) => r.id === 'rIdJump')).toBe(false);
    expect(relationshipReferences(xmlPartOf(pkg, alpha.partName)!.root)).not.toContain('rIdJump');
    // New slide ids follow the largest one.
    expect(presentationInfo(pkg).slides.map((s) => s.sldId)).toEqual(['256', '258', '259', '257']);
  });

  test('a repeated import copies everything again under new names', () => {
    const host = open(destDeck());
    insert(host, sourceDeck(), { sourceSlideIds: ['256'] });
    insert(host, sourceDeck(), { sourceSlideIds: ['256'] });
    const pkg = expectWhole(save(host));
    const [a, b] = presentationInfo(pkg).slides;
    expect(slideTexts(pkg).slice(0, 2)).toEqual(['Alpha|Alpha body', 'Alpha|Alpha body']);
    expect(a!.layoutPart).not.toBe(b!.layoutPart);
    expect(a!.masterPart).not.toBe(b!.masterPart);
    expect(a!.notesPart).not.toBe(b!.notesPart);
    expect(masterCount(pkg)).toBe(3);
  });

  test('a destination without a notes master takes the source one', () => {
    const bare = patch(destDeck(), {
      'ppt/presentation.xml': (t) => t.replace(/<p:notesMasterIdLst>.*?<\/p:notesMasterIdLst>/, ''),
      'ppt/_rels/presentation.xml.rels': (t) => t.replace(/<Relationship Id="rId2"[^>]*\/>/, ''),
    });
    const host = open(bare);
    insert(host, sourceDeck(), { sourceSlideIds: ['256'] });
    const pkg = expectWhole(save(host));
    const presentation = xmlPartOf(pkg, presentationInfo(pkg).presentationPart)!.root;
    const list = child(presentation, 'p:notesMasterIdLst');
    expect(elements(list)).toHaveLength(1);
    expect(presentation.children.indexOf(list!)).toBe(
      presentation.children.indexOf(child(presentation, 'p:sldMasterIdLst')!) + 1
    );
  });

  test('custom XML parts of a slide come along with new item ids', () => {
    const source = open(sourceDeck());
    const [slide] = ok(source, [{ op: 'getSlide', key: { index: 0 } }]);
    ok(source, [
      {
        op: 'addCustomXmlPart',
        owner: (slide as { handle: AutomationHandle }).handle,
        xml: '<d xmlns="urn:d"/>',
        id: '{12345678-1234-4234-8234-123456789012}',
      },
    ]);
    const bytes = save(source);
    const host = open(destDeck());
    insert(host, bytes, { sourceSlideIds: ['256'] });
    insert(host, bytes, { sourceSlideIds: ['256'] });
    const [one, two] = ok(host, [
      { op: 'getSlide', key: { index: 0 } },
      { op: 'getSlide', key: { index: 1 } },
    ]);
    const parts = ok(host, [
      { op: 'listCustomXmlParts', owner: (one as { handle: AutomationHandle }).handle },
      { op: 'listCustomXmlParts', owner: (two as { handle: AutomationHandle }).handle },
    ]).map((v) => (v.kind === 'customXmlParts' ? v.parts : []));
    expect(parts.map((p) => p.map((x) => x.namespaceUri))).toEqual([['urn:d'], ['urn:d']]);
    const ids = [parts[0]![0]!.id, parts[1]![0]!.id, '{12345678-1234-4234-8234-123456789012}'];
    expect(new Set(ids).size).toBe(3);
    expectWhole(save(host));
  });
});

describe('insertSlides: use destination theme', () => {
  test('each slide goes on this presentation layout of its type, else the first layout', () => {
    const host = open(destDeck());
    insert(host, sourceDeck(), { formatting: 'useDestinationTheme' });
    const pkg = expectWhole(save(host));
    expect(masterCount(pkg)).toBe(1);
    const slides = presentationInfo(pkg).slides;
    expect(slides.map((s) => s.layoutPart)).toEqual([
      '/ppt/slideLayouts/slideLayout1.xml',
      '/ppt/slideLayouts/slideLayout3.xml',
      '/ppt/slideLayouts/slideLayout1.xml',
      '/ppt/slideLayouts/slideLayout1.xml',
      '/ppt/slideLayouts/slideLayout2.xml',
    ]);
    expect(new Set(slides.map((s) => s.themePart))).toEqual(new Set(['/ppt/theme/theme1.xml']));
    // No layout, master, or theme of the source came along.
    expect(
      [...pkg.parts.keys()]
        .filter((p) => /slideLayout|slideMaster|\/theme\//.test(p) && !p.endsWith('.rels'))
        .sort()
    ).toEqual([
      '/ppt/slideLayouts/slideLayout1.xml',
      '/ppt/slideLayouts/slideLayout2.xml',
      '/ppt/slideLayouts/slideLayout3.xml',
      '/ppt/slideMasters/slideMaster1.xml',
      '/ppt/theme/theme1.xml',
      '/ppt/theme/theme2.xml',
    ]);
    expect(slideTexts(pkg).slice(0, 3)).toEqual(['Alpha|Alpha body', 'Beta', 'Gamma|Gamma body']);
    // A placeholder with a place inherits its box; the picture keeps its own.
    const alpha = xmlPartOf(pkg, slides[0]!.partName)!.root;
    const shapes = elements(path(alpha, 'p:cSld', 'p:spTree')).filter(
      (e) => e.local === 'sp' || e.local === 'pic'
    );
    expect(child(child(shapes[0]!, 'p:spPr'), 'a:xfrm')).toBeUndefined();
    expect(child(child(shapes[2]!, 'p:spPr'), 'a:xfrm')).toBeDefined();
  });
});

describe('insertSlides: refusals', () => {
  test('bad input is refused and changes nothing', () => {
    const host = open(destDeck());
    const source = toBase64(sourceDeck());
    const op = (base64: string, options?: AutomationInsertSlidesOptions): AutomationOperation => ({
      op: 'insertSlides',
      base64,
      ...(options ? { options } : {}),
    });
    expect(refusal(host, [op('')]).detail).toBe('import-base64');
    expect(refusal(host, [op('!!not base64!!')]).detail).toBe('import-base64');
    expect(refusal(host, [op(toBase64(strToU8('plain text')))]).detail).toBe('import-package:zip');
    expect(refusal(host, [op(toBase64(zipSync({ 'a.txt': strToU8('x') })))]).detail).toBe(
      'import-package:no-content-types'
    );
    expect(refusal(host, [op(source, { sourceSlideIds: ['999'] })])).toEqual({
      code: 'item-not-found',
      detail: 'import-slide',
    });
    expect(refusal(host, [op(source, { targetSlideId: '999' })])).toEqual({
      code: 'item-not-found',
      detail: 'import-target',
    });
    expect(refusal(host, [op(source, { sourceSlideIds: ['abc'] })]).detail).toBe('sourceSlideIds');
    expect(refusal(host, [op(source, { sourceSlideIds: ['256', '256'] })]).detail).toBe(
      'sourceSlideIds'
    );
    expect(refusal(host, [op(source, { formatting: 'other' as never })]).detail).toBe('formatting');
    expect(
      refusal(host, [{ op: 'insertSlides', base64: source, options: null as never }]).detail
    ).toBe('options');
  });

  test('a slide with threaded comments is refused, naming the slide', () => {
    const threaded = [...REFUSED].find((type) => type.endsWith('/2018/10/relationships/comments'));
    const withThreads = patch(sourceDeck(), {
      'ppt/slides/_rels/slide2.xml.rels': (t) =>
        t.replace(
          '</Relationships>',
          `<Relationship Id="rIdC" Type="${threaded}" Target="../comments/modernComment1.xml"/></Relationships>`
        ),
      'ppt/comments/modernComment1.xml': () => '<cmLst xmlns="urn:threaded"/>',
      '[Content_Types].xml': (t) =>
        t.replace('</Types>', '<Default Extension="xml" ContentType="application/xml"/></Types>'),
    });
    const host = open(destDeck());
    expect(refusal(host, [{ op: 'insertSlides', base64: toBase64(withThreads) }])).toEqual({
      code: 'unsupported-content',
      detail: 'disabled.slideNotCopyable:import-comments:257',
    });
    // The other slides import.
    insert(host, withThreads, { sourceSlideIds: ['256', '258'] });
    expectWhole(save(host));
  });

  test('comment lists come along, with their authors merged by name and initials', () => {
    const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
    const authorsCt = `${CT}.commentAuthors+xml`;
    const author = (id: number, name: string, initials: string, lastIdx: number) =>
      `<p:cmAuthor id="${id}" name="${name}" initials="${initials}" lastIdx="${lastIdx}" clrIdx="${id}"/>`;
    const withAuthors = (
      bytes: Uint8Array,
      authors: string,
      extra: Record<string, (t: string) => string> = {}
    ) =>
      patch(bytes, {
        'ppt/commentAuthors.xml': () => `<p:cmAuthorLst xmlns:p="${P}">${authors}</p:cmAuthorLst>`,
        'ppt/_rels/presentation.xml.rels': (t) =>
          t.replace(
            '</Relationships>',
            `<Relationship Id="rIdAuthors" Type="${R}/commentAuthors" Target="commentAuthors.xml"/></Relationships>`
          ),
        '[Content_Types].xml': (t) =>
          t.replace(
            '</Types>',
            `<Override PartName="/ppt/commentAuthors.xml" ContentType="${authorsCt}"/>` +
              `<Override PartName="/ppt/comments/comment1.xml" ContentType="${CT}.comments+xml"/></Types>`
          ),
        ...extra,
      });
    const cm = (authorId: number, idx: number, text: string, reply = '') =>
      `<p:cm authorId="${authorId}" dt="2026-01-01T00:00:00.000" idx="${idx}"><p:pos x="0" y="0"/><p:text>${text}</p:text>${reply}</p:cm>`;
    const parent = (authorId: number, idx: number) =>
      `<p:extLst><p:ext uri="{C676402C-5697-4E1C-873F-D02D1690AC5C}"><p15:threadingInfo xmlns:p15="${P15}" timeZoneBias="0"><p15:parentCm authorId="${authorId}" idx="${idx}"/></p15:threadingInfo></p:ext></p:extLst>`;
    // The source lists Ann twice (ids 0 and 2): both become one destination author, and the
    // reply to her first comment must still find that comment, not the second.
    const sourceAuthors =
      author(0, 'Ann', 'A', 1) + author(1, 'Bob', 'B', 2) + author(2, 'Ann', 'A', 1);
    const source = withAuthors(sourceDeck(), sourceAuthors, {
      'ppt/slides/_rels/slide2.xml.rels': (t) =>
        t.replace(
          '</Relationships>',
          `<Relationship Id="rIdC" Type="${R}/comments" Target="../comments/comment1.xml"/></Relationships>`
        ),
      'ppt/comments/comment1.xml': () =>
        `<p:cmLst xmlns:p="${P}">${cm(0, 1, 'Question')}${cm(2, 1, 'Second')}` +
        `${cm(1, 1, 'Answer', parent(0, 1))}${cm(1, 2, 'Orphan', parent(7, 1))}${cm(7, 1, 'Nobody')}</p:cmLst>`,
    });
    // The destination knows Bob and Ann, with comments already.
    // An author whose id is not an unsignedInt is no author a comment can name.
    const invalid =
      '<p:cmAuthor id="x" name="Ann" initials="A" lastIdx="0" clrIdx="0"/>' +
      '<p:cmAuthor id="4294967296" name="Zed" initials="Z" lastIdx="0" clrIdx="0"/>';
    const host = open(
      withAuthors(destDeck(), invalid + author(0, 'Bob', 'B', 3) + author(1, 'Ann', 'A', 5))
    );
    insert(host, source, { sourceSlideIds: ['257'] });
    const pkg = expectWhole(save(host));
    const authorsPart = relationshipsOf(pkg, '/ppt/presentation.xml').find((r) =>
      r.type.endsWith('/commentAuthors')
    )?.target;
    const authors = elements(treeOf(pkg, authorsPart)?.root).map((a) =>
      ['id', 'name', 'initials', 'lastIdx'].map((n) => attr(a, n))
    );
    expect(authors).toEqual([
      ['x', 'Ann', 'A', '0'],
      ['4294967296', 'Zed', 'Z', '0'],
      ['0', 'Bob', 'B', '5'],
      ['1', 'Ann', 'A', '7'],
    ]);
    const slide = presentationInfo(pkg).slides[0]!.partName;
    const list = relationshipsOf(pkg, slide).find((r) => r.type === `${R}/comments`)?.target;
    expect(contentTypeOf(pkg, list ?? '')).toBe(`${CT}.comments+xml`);
    const cms = elements(treeOf(pkg, list)?.root);
    // Each comment takes its author's destination id and next index; a comment by an author
    // the source does not list is dropped.
    expect(cms.map((c) => [attr(c, 'authorId'), attr(c, 'idx')])).toEqual([
      ['1', '6'],
      ['1', '7'],
      ['0', '4'],
      ['0', '5'],
    ]);
    // The reply follows the comment it answered, by its source author and index.
    const parentCm = elements(elements(elements(path(cms[2], 'p:extLst'))[0])[0])[0];
    expect([parentCm?.ns, attr(parentCm, 'authorId'), attr(parentCm, 'idx')]).toEqual([
      P15,
      '1',
      '6',
    ]);
    // A reply whose parent was dropped loses its threading extension.
    expect(path(cms[3], 'p:extLst')).toBeUndefined();
  });

  test('a destination without comment authors gets a list', () => {
    const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
    const source = patch(sourceDeck(), {
      'ppt/commentAuthors.xml': () =>
        `<p:cmAuthorLst xmlns:p="${P}"><p:cmAuthor id="3" name="Ann" initials="A" lastIdx="2" clrIdx="0"/></p:cmAuthorLst>`,
      'ppt/_rels/presentation.xml.rels': (t) =>
        t.replace(
          '</Relationships>',
          `<Relationship Id="rIdAuthors" Type="${R}/commentAuthors" Target="commentAuthors.xml"/></Relationships>`
        ),
      'ppt/slides/_rels/slide2.xml.rels': (t) =>
        t.replace(
          '</Relationships>',
          `<Relationship Id="rIdC" Type="${R}/comments" Target="../comments/comment1.xml"/></Relationships>`
        ),
      'ppt/comments/comment1.xml': () =>
        `<p:cmLst xmlns:p="${P}"><p:cm authorId="3" dt="2026-01-01T00:00:00.000" idx="2"><p:pos x="0" y="0"/><p:text>Hi</p:text></p:cm></p:cmLst>`,
      '[Content_Types].xml': (t) =>
        t.replace(
          '</Types>',
          `<Override PartName="/ppt/commentAuthors.xml" ContentType="${CT}.commentAuthors+xml"/>` +
            `<Override PartName="/ppt/comments/comment1.xml" ContentType="${CT}.comments+xml"/></Types>`
        ),
    });
    const host = open(destDeck());
    insert(host, source);
    const pkg = expectWhole(save(host));
    const authorsPart = relationshipsOf(pkg, '/ppt/presentation.xml').find((r) =>
      r.type.endsWith('/commentAuthors')
    )?.target;
    expect(contentTypeOf(pkg, authorsPart ?? '')).toBe(`${CT}.commentAuthors+xml`);
    expect(
      elements(treeOf(pkg, authorsPart)?.root).map((a) =>
        ['id', 'name', 'lastIdx'].map((n) => attr(a, n))
      )
    ).toEqual([['0', 'Ann', '1']]);
  });

  test('a crafted relationship never copies a relationships part or a structural part', () => {
    const extra = (id: string, type: string, target: string) =>
      `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`;
    const crafted = patch(sourceDeck(), {
      'ppt/slides/_rels/slide1.xml.rels': (t) =>
        t.replace(
          '</Relationships>',
          extra('rX1', 'image', '_rels/slide2.xml.rels') +
            extra('rX2', 'image', 'slide2.xml') +
            extra('rX3', 'image', '../presentation.xml') +
            extra('rX4', 'image', '../slideMasters/slideMaster1.xml') +
            '</Relationships>'
        ),
      'ppt/slides/_rels/slide2.xml.rels': (t) =>
        t.replace(
          '</Relationships>',
          `${extra('rX5', 'notesSlide', 'slide1.xml')}</Relationships>`
        ),
    });
    for (const formatting of ['keepSource', 'useDestinationTheme'] as const) {
      const host = open(destDeck());
      insert(host, crafted, { formatting, sourceSlideIds: ['256', '257'] });
      const pkg = expectWhole(save(host));
      const parts = [...pkg.parts.keys()];
      expect(parts.filter((n) => /^\/ppt\/slides\/slide\d+\.xml$/.test(n))).toHaveLength(4);
      expect(parts.filter((n) => n.endsWith('.xml') && n.includes('presentation'))).toEqual([
        '/ppt/presentation.xml',
      ]);
      expect(parts.filter((n) => n.includes('/_rels/') && !n.endsWith('.rels'))).toEqual([]);
      for (const part of parts.filter((n) => /\/slides\/slide\d+\.xml$/.test(n)))
        for (const r of relationshipsOf(pkg, part))
          if (r.type.endsWith('/image'))
            expect(contentTypeOf(pkg, r.target ?? '')).toBe('image/png');
          else if (r.type.endsWith('/notesSlide'))
            expect(contentTypeOf(pkg, r.target ?? '')).toBe(`${CT}.notesSlide+xml`);
    }
  });

  test('an import is one atomic batch with the rest of the request', () => {
    const host = open(destDeck());
    const response = exec(host, [
      { op: 'insertSlides', base64: toBase64(sourceDeck()) },
      { op: 'getSlide', key: { index: 99 } },
    ]);
    expect(response.ok).toBe(false);
    expect(presentationInfo(readBack(save(host))).slides).toHaveLength(2);
  });
});
