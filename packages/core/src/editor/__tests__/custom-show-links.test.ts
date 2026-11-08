// Links to custom shows and to other files: the `#custom-show-<id>` and `./path` forms of
// `setLink`, the action and relationship they write, the read back, the refusals, the link form's
// helpers, and the popup and slide show, which never open a file or a program.

import { describe, expect, test } from 'bun:test';
import { createTranslator } from '@pptx-editor.dev/i18n';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import {
  attr,
  elements,
  find,
  readPptx,
  relationshipsOf,
  xmlPartOf,
  type Element,
} from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { createHyperlinkPopupController } from '../link-popup.ts';
import {
  customShowDraft,
  customShowLinkLabel,
  linkKindDraft,
  linkKindOf,
  linkKindsFor,
} from '../link-fields.ts';
import { slideshowLinkAt } from '../slideshow-view.ts';
import type { ClipboardPayload } from '../../contracts/index.ts';
import { createServerAutomationHost } from '../../automation/index.ts';
import {
  customShowLinkOf,
  fileDraftUrl,
  fileLinkTarget,
  fileLinkUrl,
} from '../../automation/commands/custom-show-links.ts';

const IN = 914_400;
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const SHOWS =
  '<p:custShowLst><p:custShow name="Short" id="0"><p:sldLst><p:sld r:id="rId101"/></p:sldLst>' +
  '</p:custShow><p:custShow name="Two &amp; more" id="7"><p:sldLst><p:sld r:id="rId102"/>' +
  '</p:sldLst></p:custShow></p:custShowLst>';
const slide = (text: string) => ({
  shapes: shape({ id: 2, name: 'T', x: IN, y: IN, cx: 6 * IN, cy: 2 * IN, text: para(run(text)) }),
});

async function open(shows = SHOWS, extra = {}): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const deck = buildPptx({
    slides: [{ ...slide('Go there'), ...extra }, slide('Two')],
    presentationLists: shows,
  });
  expect((await e.load(deck)).ok).toBe(true);
  return e;
}

const root = (e: PptxEditorInstance): Element =>
  xmlPartOf(e.session!.store.package, e.session!.info.slides[0]!.partName)!.root;
const shapeId = (e: PptxEditorInstance): string =>
  elements(find(root(e), 'p:spTree')!).find((c) => attr(find(c, 'p:cNvPr'), 'name') === 'T')!.id;
const selectWord = (e: PptxEditorInstance): void =>
  e.setSelection({
    kind: 'text',
    nodeId: shapeId(e),
    anchor: { paragraph: 0, offset: 0 },
    focus: { paragraph: 0, offset: 2 },
  });
const link = (e: PptxEditorInstance): Element | undefined => find(root(e), 'a:hlinkClick');
const rels = (e: PptxEditorInstance) =>
  relationshipsOf(e.session!.store.package, e.session!.info.slides[0]!.partName).filter((r) =>
    r.type.endsWith('/hyperlink')
  );

describe('custom shows', () => {
  test('the snapshot lists the custom shows of the deck', async () => {
    const e = await open();
    expect(e.snapshot().customShows).toEqual([
      { id: 0, name: 'Short' },
      { id: 7, name: 'Two & more' },
    ]);
    expect(e.snapshot().customShows).toBe(e.snapshot().customShows);
    const none = await open('');
    expect(none.snapshot().customShows).toEqual([]);
  });

  test('setLink writes the custom show action with no relationship, and reads it back', async () => {
    const e = await open();
    selectWord(e);
    expect(e.exec({ type: 'setLink', url: '#custom-show-7?return=true' }).ok).toBe(true);
    expect([attr(link(e), 'r:id'), attr(link(e), 'action')]).toEqual([
      '',
      'ppaction://customshow?id=7&return=true',
    ]);
    expect(rels(e)).toEqual([]);
    expect(e.snapshot().format?.link).toBe('#custom-show-7?return=true');
    expect(e.exec({ type: 'setLink', url: '#custom-show-0' }).ok).toBe(true);
    expect(attr(link(e), 'action')).toBe('ppaction://customshow?id=0');
    // Save escapes the ampersand, and the link reads back the same.
    const saved = e.save()!;
    expect(readPptx(saved).ok).toBe(true);
    const again = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await again.load(saved)).ok).toBe(true);
    selectWord(again);
    expect(again.snapshot().format?.link).toBe('#custom-show-0');
  });

  test('a show the deck does not have is refused', async () => {
    const e = await open();
    selectWord(e);
    for (const url of ['#custom-show-3', '#custom-show-x', '#custom-show-0?return=no'])
      expect(e.exec({ type: 'setLink', url })).toMatchObject({ ok: false });
    expect(link(e)).toBeUndefined();
  });

  test('reads the action in any field order, and no other action', () => {
    expect(customShowLinkOf('ppaction://customshow?return=true&id=4')).toBe(
      '#custom-show-4?return=true'
    );
    expect(customShowLinkOf('ppaction://customshow?id=4&return=false')).toBe('#custom-show-4');
    expect(customShowLinkOf('ppaction://customshow?id=x')).toBeUndefined();
    expect(customShowLinkOf('ppaction://hlinkshowjump?jump=nextslide')).toBeUndefined();
  });

  test('the form offers Custom show only with shows, and labels a show by its name', () => {
    const t = createTranslator();
    const shows = [{ id: 7, name: 'Short' }];
    expect(linkKindsFor([], '')).not.toContain('customShow');
    expect(linkKindsFor([], '#custom-show-2')).toContain('customShow');
    expect(linkKindsFor(shows, '')).toContain('customShow');
    expect(linkKindDraft('customShow', shows)).toBe('#custom-show-7?return=true');
    expect(linkKindOf('#custom-show-7')).toBe('customShow');
    expect(customShowDraft('#custom-show-7')).toEqual({ id: 7, returns: false });
    expect(customShowLinkLabel(t, '#custom-show-7', shows)).toBe('Custom show: Short');
    expect(customShowLinkLabel(t, '#custom-show-9', shows)).toBe('Custom show: 9');
    expect(customShowLinkLabel(t, '#slide-2', shows)).toBe('');
  });
});

describe('links to other files', () => {
  test('setLink writes a relative path as an external hyperlink with no action', async () => {
    const e = await open();
    selectWord(e);
    expect(e.exec({ type: 'setLink', url: './docs/Q3 report.pdf' }).ok).toBe(true);
    expect(attr(link(e), 'action')).toBeUndefined();
    const rel = rels(e)[0]!;
    expect([rel.external, rel.externalTarget, rel.relativeTarget]).toEqual([
      true,
      undefined,
      'docs/Q3%20report.pdf',
    ]);
    expect(e.snapshot().format?.link).toBe('./docs/Q3 report.pdf');
    expect(linkKindOf('./docs/Q3 report.pdf')).toBe('file');
  });

  test('refuses paths that escape, name a scheme, or name a program', () => {
    for (const bad of [
      'report.pdf',
      './',
      './/x.pdf',
      '../../../x.pdf',
      './a/../b.pdf',
      './c:/x.pdf',
      './a\\b.pdf',
      './file:x',
      './x\u0000.pdf',
      './run.exe',
      './Script.JS',
      './tool.sh',
      './x.pdf ',
    ])
      expect(() => fileLinkTarget(bad)).toThrow();
    expect(fileLinkTarget('../../shared/a b.pdf')).toBe('../../shared/a%20b.pdf');
    expect(fileLinkTarget('./notes.txt')).toBe('notes.txt');
  });

  test('reads a relative target in its ./ form, and no URL or absolute path', () => {
    expect(fileLinkUrl('notes.txt')).toBe('./notes.txt');
    expect(fileLinkUrl('../a%20b.pdf')).toBe('../a b.pdf');
    for (const other of ['https://a.test/', 'file:///c:/x', '/abs/x', '\\\\server\\x', '#x'])
      expect(fileLinkUrl(other)).toBeUndefined();
  });

  test('the popup and the slide show never open a file or a program', async () => {
    const text =
      '<a:p><a:r><a:rPr lang="en-US" sz="4000"><a:hlinkClick r:id="rIdFile"/></a:rPr>' +
      '<a:t>File</a:t></a:r><a:r><a:rPr lang="en-US" sz="4000"><a:hlinkClick r:id="rIdTool" ' +
      'action="ppaction://program"/></a:rPr><a:t>Tool</a:t></a:r></a:p>';
    const e = await open(SHOWS, {
      shapes: shape({ id: 2, name: 'T', x: 0, y: 0, cx: 6 * IN, cy: 2 * IN, text }),
      rels: [
        ['rIdFile', `${R}/hyperlink`, 'notes.txt', 'External'],
        ['rIdTool', `${R}/hyperlink`, 'tool.exe', 'External'],
      ],
    });
    const popup = createHyperlinkPopupController(e);
    e.setSelection({
      kind: 'text',
      nodeId: shapeId(e),
      anchor: { paragraph: 0, offset: 1 },
      focus: { paragraph: 0, offset: 1 },
    });
    popup.openReading(null);
    expect(popup.getState().url).toBe('./notes.txt');
    expect(popup.getState().href).toBeNull();
    expect(popup.openTarget()).toBe(false);
    e.setSelection({
      kind: 'text',
      nodeId: shapeId(e),
      anchor: { paragraph: 0, offset: 6 },
      focus: { paragraph: 0, offset: 6 },
    });
    // A program link shows its target as stored, and is never opened.
    expect(e.snapshot().format?.link).toBe('tool.exe');
    // The slide show finds no link to open in either run.
    const lines = e.session!.layout(0)!.items.find((x) => x.kind === 'shape')!;
    expect(lines.kind).toBe('shape');
    for (const x of [IN / 4, 2 * IN])
      expect(slideshowLinkAt(e.session!, 0, x, 0.4 * IN)).toBeNull();
    popup.dispose();
  });
});

describe('file link checks', () => {
  const refused = (path: string): boolean => {
    try {
      fileLinkTarget(path);
      return false;
    } catch {
      return true;
    }
  };

  test('a trailing dot or space cannot hide a program', () => {
    for (const path of ['./a.exe.', './a.exe..', './a.exe. .', './a.pdf.', './dir./a.pdf'])
      expect(refused(path)).toBe(true);
  });

  test('only a document, picture, sound, or video file can be linked', () => {
    for (const path of [
      './a.py',
      './a.pyw',
      './x.scpt',
      './a.chm',
      './a.iso',
      './a.html',
      './a.svg',
      './a.pptm',
      './a.docm',
      './a.xlam',
      './a.msix',
      './a.appref-ms',
      './a.settingcontent-ms',
      './ａ.exe',
      './a．exe',
      './a.exе',
      './.pdf',
      './README',
      './~',
      './a.exe::$DATA',
    ])
      expect(refused(path)).toBe(true);
    for (const path of ['./a.pdf', './b.DOCX', './c.png', './d.mp4', '../e.xlsx'])
      expect(refused(path)).toBe(false);
  });

  test('a device name is refused in any case and with any extension', () => {
    for (const path of ['./CON.pdf', './nul.txt', './Aux.png', './com1.pdf', './LPT9.txt'])
      expect(refused(path)).toBe(true);
    expect(refused('./CON/a.pdf')).toBe(true);
    expect(refused('./console.pdf')).toBe(false);
  });

  test('names with combining marks are taken, in NFC', () => {
    const nfd = './café.pdf';
    expect(fileLinkTarget(nfd)).toBe(encodeURIComponent('café.pdf'));
    expect(refused('./नमस्ते.pdf')).toBe(false);
    expect(refused('./สวัสดี.pdf')).toBe(false);
  });

  test('a decoded target that hides a character or steps up too far shows as written', () => {
    for (const raw of [
      '%2e%2e%2f%2e%2e%2f%2e%2e%2fetc%2fpasswd',
      'javascript%3Aalert(1).pdf',
      'report%E2%80%AEfdp.exe',
      'a%0Ab.pdf',
      '%5C%5Chost%5Cshare.pdf',
    ])
      expect(fileLinkUrl(raw)).toBe(`./${raw}`);
    expect(fileLinkUrl('docs/Q3%20report.pdf')).toBe('./docs/Q3 report.pdf');
  });

  test('the file path field adds ./ to a bare name', () => {
    expect(fileDraftUrl('report.pdf')).toBe('./report.pdf');
    expect(fileDraftUrl('../a.pdf')).toBe('../a.pdf');
    expect(fileDraftUrl('./a.pdf')).toBe('./a.pdf');
    for (const text of ['', '.', '..']) expect(fileDraftUrl(text)).toBe(text);
  });
});

describe('custom show list', () => {
  test('a repeated id lists once, and a name loses its format characters', async () => {
    const shows =
      '<p:custShowLst><p:custShow name="A‮b​" id="4"><p:sldLst/></p:custShow>' +
      '<p:custShow name="Dup" id="4"><p:sldLst/></p:custShow></p:custShowLst>';
    const e = await open(shows);
    expect(e.snapshot().customShows).toEqual([{ id: 4, name: 'Ab' }]);
    e.destroy();
  });
});

describe('file links through copy, paste, and automation', () => {
  const linked = (): Promise<PptxEditorInstance> =>
    open(SHOWS, {
      shapes: shape({
        id: 2,
        name: 'T',
        x: IN,
        y: IN,
        cx: 6 * IN,
        cy: 2 * IN,
        text:
          '<a:p><a:r><a:rPr lang="en-US"><a:hlinkClick r:id="rIdFile"/></a:rPr><a:t>File</a:t>' +
          '</a:r><a:r><a:rPr lang="en-US"><a:hlinkClick r:id="rIdTool"/></a:rPr><a:t>Tool</a:t>' +
          '</a:r></a:p>',
      }),
      rels: [
        ['rIdFile', `${R}/hyperlink`, 'docs/a%20b.pdf', 'External'],
        ['rIdTool', `${R}/hyperlink`, 'tool.exe', 'External'],
      ],
    });

  test('a pasted shape keeps a safe file link and drops an unsafe one', async () => {
    const e = await linked();
    e.setSelection({ kind: 'shapes', nodeIds: [shapeId(e)] });
    const payload = JSON.parse(JSON.stringify(e.copy())) as ClipboardPayload;
    expect(payload.relationships.map((r) => r.target)).toEqual(['./docs/a b.pdf', './tool.exe']);
    expect(e.exec({ type: 'paste', payload }).ok).toBe(true);
    const targets = rels(e).map((r) => r.relativeTarget);
    expect(targets.filter((t) => t === 'docs/a%20b.pdf')).toHaveLength(2);
    expect(targets.filter((t) => t === 'tool.exe')).toHaveLength(1);
    // A changed payload passes the same checks: a trailing dot is refused.
    const bad = {
      ...payload,
      relationships: payload.relationships.map((r) => ({ ...r, target: './a.exe.' })),
    };
    expect(e.exec({ type: 'paste', payload: bad }).ok).toBe(true);
    expect(rels(e).some((r) => r.relativeTarget?.startsWith('a.exe'))).toBe(false);
    e.destroy();
  });

  test('automation reads the address of a file link', async () => {
    const e = await linked();
    const opened = createServerAutomationHost(e.save()!);
    if (!opened.ok) throw new Error('not opened');
    const response = opened.host.execute({
      operations: [
        { op: 'getSlide', key: { index: 0 } },
        { op: 'listHyperlinks', slide: { result: 0 } },
      ],
    });
    expect(response.ok).toBe(true);
    const value = response.results[1];
    expect(value?.status === 'ok' ? value.value : null).toMatchObject({
      hyperlinks: [{ address: 'docs/a%20b.pdf' }, { address: 'tool.exe' }],
    });
    e.destroy();
  });
});
