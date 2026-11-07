// Links to a slide of the presentation and to an email address: the `#` forms of `setLink`, the
// relationship and action they write, the state after slides move or go, the popup's Go to
// slide, the slide show, and the link form's helpers.

import { describe, expect, test } from 'bun:test';
import { createTranslator } from '@pptx-editor.dev/i18n';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import {
  attr,
  child,
  textOf,
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
  linkKindOf,
  mailtoParts,
  mailtoUrl,
  slideLinkChoices,
  slideLinkIndex,
  slideLinkLabel,
} from '../link-fields.ts';
import { followSlideshowLink } from '../slideshow-screen.ts';
import { SlideshowController } from '../slideshow.ts';
import { parseSlideLink } from '../../automation/commands/slide-links.ts';

const IN = 914_400;
const slide = (text: string) => ({
  shapes: shape({ id: 2, name: 'T', x: IN, y: IN, cx: 6 * IN, cy: 2 * IN, text: para(run(text)) }),
});

async function open(): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const deck = buildPptx({ slides: [slide('Go there'), slide('Two'), slide('Three')] });
  expect((await e.load(deck)).ok).toBe(true);
  return e;
}

const partOf = (e: PptxEditorInstance, i: number): string => e.session!.info.slides[i]!.partName;
const root = (e: PptxEditorInstance, i = 0): Element =>
  xmlPartOf(e.session!.store.package, partOf(e, i))!.root;
const shapeId = (e: PptxEditorInstance): string =>
  elements(find(root(e), 'p:spTree')!).find((c) => attr(find(c, 'p:cNvPr'), 'name') === 'T')!.id;
const linkEl = (r: Element): Element | undefined => find(r, 'a:hlinkClick');
/** Select "Go" on the first slide. */
const selectWord = (e: PptxEditorInstance): void =>
  e.setSelection({
    kind: 'text',
    nodeId: shapeId(e),
    anchor: { paragraph: 0, offset: 0 },
    focus: { paragraph: 0, offset: 2 },
  });
const caretLink = (e: PptxEditorInstance): string | undefined => {
  e.goToSlide(0);
  e.setSelection({
    kind: 'text',
    nodeId: shapeId(e),
    anchor: { paragraph: 0, offset: 1 },
    focus: { paragraph: 0, offset: 1 },
  });
  return e.snapshot().format?.link;
};
/** The slide relationship targets of slide 1. */
function slideRels(e: PptxEditorInstance): Map<string, string> {
  const out = new Map<string, string>();
  for (const r of relationshipsOf(e.session!.store.package, partOf(e, 0)))
    if (r.type.endsWith('/slide')) out.set(r.id, r.target ?? '');
  return out;
}

describe('setLink to a slide', () => {
  test('#slide-N writes a slide jump with a slide relationship, and the state reads it back', async () => {
    const e = await open();
    selectWord(e);
    expect(e.exec({ type: 'setLink', url: '#slide-3' }).ok).toBe(true);
    const link = linkEl(root(e))!;
    expect(attr(link, 'action')).toBe('ppaction://hlinksldjump');
    const rId = attr(link, 'r:id')!;
    expect(slideRels(e).get(rId)).toBeDefined();
    expect(caretLink(e)).toBe('#slide-3');
    // Save and reopen keep the jump.
    const reopened = readPptx(e.save()!);
    expect(reopened.ok).toBe(true);
  });

  test('a relative jump writes the show action with an empty relationship', async () => {
    const e = await open();
    selectWord(e);
    expect(e.exec({ type: 'setLink', url: '#next-slide' }).ok).toBe(true);
    const link = linkEl(root(e))!;
    expect([attr(link, 'r:id'), attr(link, 'action')]).toEqual([
      '',
      'ppaction://hlinkshowjump?jump=nextslide',
    ]);
    expect(caretLink(e)).toBe('#next-slide');
    expect(slideRels(e).size).toBe(0);
  });

  test('the link follows its slide when slides move, and goes when the slide is deleted', async () => {
    const e = await open();
    selectWord(e);
    expect(e.exec({ type: 'setLink', url: '#slide-3' }).ok).toBe(true);
    // Slide 3 moves to the second place: the link now reads slide 2.
    expect(e.exec({ type: 'moveSlide', from: 2, to: 1 }).ok).toBe(true);
    expect(caretLink(e)).toBe('#slide-2');
    // With its slide deleted, the text keeps no link to a missing slide.
    expect(e.exec({ type: 'deleteSlide', index: 1 }).ok).toBe(true);
    expect(caretLink(e)).toBeUndefined();
  });

  test('a slide past the end and a malformed # form are refused', async () => {
    const e = await open();
    selectWord(e);
    expect(e.exec({ type: 'setLink', url: '#slide-9' })).toMatchObject({
      ok: false,
      reason: 'errors.slideNotFound',
    });
    expect(e.exec({ type: 'setLink', url: '#slide-0' }).ok).toBe(false);
    expect(e.exec({ type: 'setLink', url: '#somewhere' }).ok).toBe(false);
    expect(linkEl(root(e))).toBeUndefined();
  });
});

describe('editing a link at a caret', () => {
  /** Select `from`..`to` of the first paragraph on slide 1. */
  const range = (e: PptxEditorInstance, from: number, to: number): void =>
    e.setSelection({
      kind: 'text',
      nodeId: shapeId(e),
      anchor: { paragraph: 0, offset: from },
      focus: { paragraph: 0, offset: to },
    });
  const runs = (e: PptxEditorInstance): string[] =>
    elements(find(root(e), 'a:p')!)
      .filter((r) => r.local === 'r')
      .map((r) => {
        const text = textOf(child(r, 'a:t'));
        const l = linkEl(r);
        return l ? `${text}|${attr(l, 'action')?.replace(/^.*jump=/, '')}` : text;
      });

  test('Remove and Edit change the whole jump link, and two jump links stay apart', async () => {
    const e = await open();
    // "Go there": "Go" and "there" are two links to the next slide, with a plain space between.
    range(e, 0, 2);
    expect(e.exec({ type: 'setLink', url: '#next-slide' }).ok).toBe(true);
    range(e, 3, 8);
    expect(e.exec({ type: 'setLink', url: '#next-slide' }).ok).toBe(true);
    expect(runs(e)).toEqual(['Go|nextslide', ' ', 'there|nextslide']);
    // Edit at a caret in "there" relinks all of "there" and leaves "Go".
    range(e, 5, 5);
    expect(e.exec({ type: 'setLink', url: '#last-slide' }).ok).toBe(true);
    expect(runs(e)).toEqual(['Go|nextslide', ' ', 'there|lastslide']);
    // Remove at a caret in "there" unlinks all of it.
    range(e, 5, 5);
    expect(e.exec({ type: 'setLink', url: null }).ok).toBe(true);
    expect(runs(e).filter((r) => r.includes('|'))).toEqual(['Go|nextslide']);
  });

  test('a jump link stays in its paragraph', async () => {
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const text = para(run('Go')) + para(run('there'));
    const deck = buildPptx({
      slides: [{ shapes: shape({ id: 2, name: 'T', x: IN, y: IN, cx: 6 * IN, cy: 2 * IN, text }) }],
    });
    expect((await e.load(deck)).ok).toBe(true);
    const at = (paragraph: number, from: number, to: number) =>
      e.setSelection({
        kind: 'text',
        nodeId: shapeId(e),
        anchor: { paragraph, offset: from },
        focus: { paragraph, offset: to },
      });
    // Both paragraphs link to the next slide; the runs touch across the paragraph break.
    at(0, 0, 2);
    expect(e.exec({ type: 'setLink', url: '#next-slide' }).ok).toBe(true);
    at(1, 0, 5);
    expect(e.exec({ type: 'setLink', url: '#next-slide' }).ok).toBe(true);
    // Remove at a caret in "Go" leaves the link of the next paragraph.
    at(0, 1, 1);
    expect(e.exec({ type: 'setLink', url: null }).ok).toBe(true);
    const paras = elements(find(root(e), 'p:txBody')!).filter((c) => c.local === 'p');
    const linked = paras.map((p) => elements(p).some((r) => r.local === 'r' && !!linkEl(r)));
    expect(linked).toEqual([false, true]);
  });

  test('relinking and removing leave no unused relationships', async () => {
    const e = await open();
    for (const url of ['#slide-2', '#slide-2', '#slide-3', 'mailto:a@b.c', 'https://a.test/']) {
      selectWord(e);
      expect(e.exec({ type: 'setLink', url }).ok).toBe(true);
    }
    const rels = () =>
      relationshipsOf(e.session!.store.package, partOf(e, 0)).filter(
        (r) => r.type.endsWith('/slide') || r.type.endsWith('/hyperlink')
      );
    expect(rels().length).toBe(1);
    selectWord(e);
    expect(e.exec({ type: 'setLink', url: null }).ok).toBe(true);
    expect(rels()).toEqual([]);
    // Undo brings the link and its relationship back.
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(rels().length).toBe(1);
    expect(caretLink(e)).toBe('https://a.test/');
  });
});

describe('copy and paste', () => {
  test('a pasted shape keeps its jump link, which needs no relationship', async () => {
    const e = await open();
    selectWord(e);
    expect(e.exec({ type: 'setLink', url: '#next-slide' }).ok).toBe(true);
    e.setSelection({ kind: 'shapes', nodeIds: [shapeId(e)] });
    const payload = e.copy()!;
    e.goToSlide(1);
    expect(e.exec({ type: 'paste', payload }).ok).toBe(true);
    const pasted = find(root(e, 1), 'a:hlinkClick');
    expect([attr(pasted, 'r:id'), attr(pasted, 'action')]).toEqual([
      '',
      'ppaction://hlinkshowjump?jump=nextslide',
    ]);
  });
});

describe('links with other actions', () => {
  test('an action with an external target shows that target; a custom show reads in its # form', async () => {
    const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
    const text =
      '<a:p><a:r><a:rPr lang="en-US"><a:hlinkClick r:id="rId9" action="ppaction://hlinkfile"/>' +
      '</a:rPr><a:t>File</a:t></a:r><a:r><a:rPr lang="en-US"><a:hlinkClick r:id="" ' +
      'action="ppaction://customshow?id=0"/></a:rPr><a:t> show</a:t></a:r></a:p>';
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const deck = buildPptx({
      slides: [
        {
          shapes: shape({ id: 2, name: 'T', x: IN, y: IN, cx: 6 * IN, cy: 2 * IN, text }),
          rels: [['rId9', R, 'https://a.test/report.pdf', 'External']],
        },
      ],
    });
    expect((await e.load(deck)).ok).toBe(true);
    expect(caretLink(e)).toBe('https://a.test/report.pdf');
    e.setSelection({
      kind: 'text',
      nodeId: shapeId(e),
      anchor: { paragraph: 0, offset: 6 },
      focus: { paragraph: 0, offset: 6 },
    });
    expect(e.snapshot().format?.link).toBe('#custom-show-0');
    // Save keeps the custom show action.
    const show = elements(find(root(e), 'a:p')!).filter((r) => r.local === 'r')[1]!;
    expect(attr(linkEl(show), 'action')).toBe('ppaction://customshow?id=0');
    expect(readPptx(e.save()!).ok).toBe(true);
  });
});

describe('setLink to an email address', () => {
  test('writes a mailto: link with its subject', async () => {
    const e = await open();
    selectWord(e);
    const url = mailtoUrl('ana@example.com', 'Q3 & plans');
    expect(e.exec({ type: 'setLink', url }).ok).toBe(true);
    expect(caretLink(e)).toBe('mailto:ana@example.com?subject=Q3%20%26%20plans');
  });
});

describe('Go to slide', () => {
  test('the popup shows the slide a slide link names', async () => {
    const e = await open();
    selectWord(e);
    expect(e.exec({ type: 'setLink', url: '#last-slide' }).ok).toBe(true);
    caretLink(e);
    const c = createHyperlinkPopupController(e);
    c.openReading(null);
    expect(c.getState().href).toBe('#last-slide');
    expect(c.openTarget()).toBe(true);
    expect(e.snapshot().currentSlide).toBe(2);
    expect(c.getState().mode).toBe('closed');
    c.dispose();
  });

  test('the slide show moves for a slide link and opens any other link', () => {
    const show = new SlideshowController(
      [0, 1, 2, 3].map(() => ({ hidden: false })),
      'start'
    );
    const opened: string[] = [];
    const doc = {
      createElement: () => ({
        setAttribute: (k: string, v: string) => void (k === 'href' && opened.push(v)),
        click: () => undefined,
      }),
    } as unknown as Document;
    followSlideshowLink(show, doc, '#slide-3');
    expect(show.state.index).toBe(2);
    followSlideshowLink(show, doc, '#previous-slide');
    expect(show.state.index).toBe(1);
    followSlideshowLink(show, doc, '#last-slide');
    expect(show.state.index).toBe(3);
    followSlideshowLink(show, doc, 'https://a.test/');
    expect(opened).toEqual(['https://a.test/']);
    expect(show.state.index).toBe(3);
  });
});

describe('link form helpers', () => {
  test('kinds, email parts, and slide choices round-trip through the one draft', () => {
    expect(['https://a.test', 'MAILTO:x@y.z', '#slide-2'].map(linkKindOf)).toEqual([
      'web',
      'email',
      'slide',
    ]);
    expect(mailtoParts('mailto:a%40b@c.d?subject=Hi%20there')).toEqual({
      address: 'a@b@c.d',
      subject: 'Hi there',
      rest: '',
    });
    expect(mailtoParts('https://a.test')).toEqual({ address: '', subject: '', rest: '' });
    // A plus sign stays a plus sign, and the fields the form does not edit are kept.
    const parts = mailtoParts('mailto:a@b.c?cc=d@e.f&subject=C++&body=Hi%20you');
    expect(parts).toEqual({ address: 'a@b.c', subject: 'C++', rest: 'cc=d@e.f&body=Hi%20you' });
    expect(mailtoUrl('x@y.z', parts.subject, parts.rest)).toBe(
      'mailto:x@y.z?subject=C%2B%2B&cc=d@e.f&body=Hi%20you'
    );
    // An address cannot start the fields either.
    expect(mailtoUrl('a?cc=b@c.d', '')).toBe('mailto:a%3Fcc=b@c.d');
    expect(mailtoUrl('  ', 'x')).toBe('');
    // A subject cannot add another field.
    expect(mailtoUrl('a@b.c', 'x&cc=evil@d.e')).toBe('mailto:a@b.c?subject=x%26cc%3Devil%40d.e');
    expect(slideLinkChoices(2)).toEqual([
      '#first-slide',
      '#last-slide',
      '#next-slide',
      '#previous-slide',
      '#slide-1',
      '#slide-2',
    ]);
    expect(slideLinkChoices(2).every((u) => parseSlideLink(u) !== null)).toBe(true);
    const t = createTranslator();
    expect(slideLinkLabel(t, '#slide-3')).toBe('Slide 3');
    expect(slideLinkLabel(t, '#next-slide')).toBe('Next slide');
    expect(slideLinkLabel(t, 'https://a.test')).toBe('');
    expect(slideLinkIndex('#next-slide', 2, 3)).toBe(2);
    expect(slideLinkIndex('#previous-slide', 0, 3)).toBe(0);
    expect(slideLinkIndex('#slide-4', 0, 3)).toBeNull();
  });
});
