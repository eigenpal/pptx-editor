// Links to a slide of the presentation and to an email address: the `#` forms of `setLink`, the
// relationship and action they write, the state after slides move or go, the popup's Go to
// slide, the slide show, and the link form's helpers.

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
    expect(e.exec({ type: 'setLink', url: '#slide-9' }).ok).toBe(false);
    expect(e.exec({ type: 'setLink', url: '#slide-0' }).ok).toBe(false);
    expect(e.exec({ type: 'setLink', url: '#somewhere' }).ok).toBe(false);
    expect(linkEl(root(e))).toBeUndefined();
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
    });
    expect(mailtoParts('https://a.test')).toEqual({ address: '', subject: '' });
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
