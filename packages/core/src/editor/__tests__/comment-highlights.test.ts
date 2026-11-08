// What a comment highlights on the canvas (comment-highlights.ts): a band over commented text in
// the color of the author who started the thread, open while the thread is picked or the caret
// is in it; a ring around a commented picture or shape; one band per thread, however many
// replies; the dark paper layer for light text; and markers in the roster's colors.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { PNG } from 'pngjs';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { createPptxEditor, type EditorModule, type PptxEditorInstance } from '../index.ts';

const window = new GlobalWindow();
const doc = window.document as unknown as Document;
afterAll(() => window.happyDOM.close());
const g = globalThis as { ResizeObserver?: unknown };
g.ResizeObserver ??= (window as unknown as { ResizeObserver: unknown }).ResizeObserver;

const review: EditorModule = {
  id: 'review',
  review: {
    displayModes: ['final', 'markup'],
    collectReviewItems: () => [],
    pane: { commentMarkers: 'initials' },
  },
};
const DATE = '2026-11-08T10:00:00.000';
const IN = 914_400;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

function png(): Uint8Array {
  const img = new PNG({ width: 4, height: 3 });
  img.data.fill(120);
  return new Uint8Array(PNG.sync.write(img));
}

const pic =
  `<p:pic><p:nvPicPr><p:cNvPr id="9" name="Photo"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
  '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
  `<p:spPr><a:xfrm><a:off x="${6 * IN}" y="${IN}"/><a:ext cx="${3 * IN}" cy="${2 * IN}"/>` +
  '</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';

const WHITE = '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>';

async function open(): Promise<{ e: PptxEditorInstance; host: HTMLElement }> {
  const e = createPptxEditor({
    fixedMeasurer: true,
    substitutes: false,
    modules: [review],
    author: 'Ada',
  });
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        rels: [['rIdImg', R_IMAGE, '../media/image1.png']],
        shapes:
          shape({
            id: 3,
            name: 'Body',
            x: 0,
            y: 0,
            cx: 5 * IN,
            cy: 2 * IN,
            text: para(run('The quick fox')),
          }) +
          shape({
            id: 4,
            name: 'Dark',
            x: 0,
            y: 3 * IN,
            cx: 5 * IN,
            cy: IN,
            text: para(run('Light words', '', WHITE)),
          }) +
          pic,
      },
    ],
    media: [{ name: 'image1.png', bytes: png(), contentType: 'image/png' }],
  });
  expect((await e.load(bytes)).ok).toBe(true);
  const host = doc.createElement('div') as unknown as HTMLElement;
  doc.body.append(host as unknown as Node);
  e.attach(host);
  return { e, host };
}

const idOf = (e: PptxEditorInstance, name: string): string =>
  e.snapshot().shapes!.find((s) => s.name === name)!.nodeId;

const caret = (e: PptxEditorInstance, name: string, from: number, to = from): void =>
  e.setSelection({
    kind: 'text',
    nodeId: idOf(e, name),
    anchor: { paragraph: 0, offset: from },
    focus: { paragraph: 0, offset: to },
  });

const bands = (host: HTMLElement, paper: 'light' | 'dark') =>
  [
    ...host.querySelectorAll(
      `.pptx-surface__comment-bands[data-paper='${paper}'] .pptx-surface__comment-band-group`
    ),
  ] as Element[];

describe('a band over commented text', () => {
  test("takes the starter's color, one band per thread with replies, open with the caret", async () => {
    const { e, host } = await open();
    caret(e, 'Body', 4, 9);
    expect(e.exec({ type: 'addComment', text: 'Faster?', date: DATE, author: 'Grace' }).ok).toBe(
      true
    );
    const thread = e.snapshot().comments![0]!;
    // A reply covers its thread's exact range: it draws no second band.
    e.exec({ type: 'replyToComment', commentId: thread.id, text: 'Sure', date: DATE });
    e.exec({ type: 'replyToComment', commentId: thread.id, text: 'Done', date: DATE });
    caret(e, 'Body', 0);
    const [band, ...more] = bands(host, 'light');
    expect(more.length).toBe(0);
    const grace = e.getReviewAuthors().find((a) => a.author === 'Grace')!;
    expect(band!.getAttribute('data-review-author')).toBe('Grace');
    expect(band!.getAttribute('data-review-author-slot')).toBe(String(grace.slot));
    expect(band!.getAttribute('style')).toBe(`--doc-review-author-current: ${grace.color}`);
    // At rest, then open while the caret is in the range, then at rest again.
    expect(band!.classList.contains('is-active')).toBe(false);
    expect(band!.querySelectorAll('.pptx-surface__comment-band').length).toBe(1);
    expect(band!.querySelectorAll('.pptx-surface__comment-band-rule').length).toBe(1);
    caret(e, 'Body', 6);
    expect(bands(host, 'light')[0]!.classList.contains('is-active')).toBe(true);
    caret(e, 'Body', 12);
    expect(bands(host, 'light')[0]!.classList.contains('is-active')).toBe(false);
    // Open while the thread is picked.
    e.selectComment(thread.id);
    expect(bands(host, 'light')[0]!.classList.contains('is-active')).toBe(true);
    // The marker sits at the band and shares its hooks.
    const marker = host.querySelector(`[data-comment-marker="${thread.id}"]`)!;
    expect(marker.getAttribute('data-review-author')).toBe('Grace');
    expect(marker.getAttribute('style')).toBe(band!.getAttribute('style'));
    expect(marker.querySelector('.pptx-surface__comment-marker-count-text')!.textContent).toBe('2');
  });

  test('light text is on dark paper: its band goes in the screen layer', async () => {
    const { e, host } = await open();
    caret(e, 'Dark', 0, 5);
    expect(e.exec({ type: 'addComment', text: 'Light', date: DATE }).ok).toBe(true);
    expect(bands(host, 'light').length).toBe(0);
    expect(bands(host, 'dark').length).toBe(1);
  });

  test('a resolved thread marks its band, and a moved range moves the band', async () => {
    const { e, host } = await open();
    caret(e, 'Body', 4, 9);
    e.exec({ type: 'addComment', text: 'Faster?', date: DATE });
    const x = () =>
      Number(
        bands(host, 'light')[0]!.querySelector('.pptx-surface__comment-band')!.getAttribute('x')
      );
    const before = x();
    caret(e, 'Body', 0);
    e.input('text', 'Oh ');
    expect(x()).toBeGreaterThan(before);
    const id = e.snapshot().comments![0]!.id;
    e.exec({ type: 'resolveComment', commentId: id });
    expect(bands(host, 'light')[0]!.classList.contains('is-resolved')).toBe(true);
  });
});

describe('a ring around a commented element', () => {
  test('a picture takes a ring in its author color, stronger while selected', async () => {
    const { e, host } = await open();
    const photo = idOf(e, 'Photo');
    e.setSelection({ kind: 'shapes', nodeIds: [photo] });
    expect(e.exec({ type: 'addComment', text: 'Crop?', date: DATE }).ok).toBe(true);
    expect(e.snapshot().comments![0]!.anchor).toEqual({ kind: 'shape', nodeId: photo });
    const ring = () =>
      host.querySelector('.pptx-surface__overlay .pptx-surface__comment-ring-group')!;
    const ada = e.getReviewAuthors().find((a) => a.author === 'Ada')!;
    expect(ring().getAttribute('style')).toBe(`--doc-review-author-current: ${ada.color}`);
    expect(ring().classList.contains('is-active')).toBe(true);
    // A stroke with a light halo under it, not a wash.
    expect(ring().querySelector('.pptx-surface__halo--comment-ring')).not.toBeNull();
    const width = () =>
      Number(ring().querySelector('.pptx-surface__comment-ring')!.getAttribute('stroke-width'));
    const strong = width();
    e.setSelection({ kind: 'none' });
    expect(ring().classList.contains('is-active')).toBe(false);
    expect(width()).toBeLessThan(strong);
    // No band for an element thread.
    expect(bands(host, 'light').length).toBe(0);
  });

  test('a shape selected as a whole takes a ring, not a band', async () => {
    const { e, host } = await open();
    e.setSelection({ kind: 'shapes', nodeIds: [idOf(e, 'Body')] });
    e.exec({ type: 'addComment', text: 'Whole box', date: DATE });
    expect(host.querySelectorAll('.pptx-surface__comment-ring-group')).toHaveLength(1);
    expect(bands(host, 'light').length).toBe(0);
  });
});
