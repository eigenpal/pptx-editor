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
import {
  customShowLinkOf,
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
