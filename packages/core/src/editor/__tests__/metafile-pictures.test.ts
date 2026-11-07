// EMF and WMF pictures paint the placeholder until their PNG is ready, then the editor repaints
// and the editor, the image export, and the natural size use the PNG.

import { expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { emfRectangle, wmfRectangle } from '../../../../../scripts/fixtures/picture-bytes.ts';
import { readPptx } from '../../store/index.ts';
import { paintLayout } from '../export-image.ts';
import { createPptxEditor } from '../index.ts';
import { MetafilePictures, metafilesOf } from '../metafile-pictures.ts';

const IN = 914_400;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const SLIDE = '/ppt/slides/slide1.xml';
const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));

const pic = (id: number, rId: string): string =>
  `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="P${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
  `<p:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  `<p:spPr><a:xfrm><a:off x="${id * IN}" y="${IN}"/><a:ext cx="${IN}" cy="${IN}"/></a:xfrm>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';

/** Slide 1: an EMF, a WMF, and an EMF in a part named `.png`. Slide 2: another EMF. */
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        rels: [
          ['rIdEmf', R_IMAGE, '../media/image1.emf'],
          ['rIdWmf', R_IMAGE, '../media/image2.wmf'],
          ['rIdNamed', R_IMAGE, '../media/image3.png'],
        ],
        shapes: pic(2, 'rIdEmf') + pic(4, 'rIdWmf') + pic(6, 'rIdNamed'),
      },
      {
        layout: 'blank',
        rels: [['rIdOther', R_IMAGE, '../media/image4.emf']],
        shapes: pic(2, 'rIdOther'),
      },
    ],
    media: [
      { name: 'image1.emf', bytes: emfRectangle(), contentType: 'image/x-emf' },
      { name: 'image2.wmf', bytes: wmfRectangle(), contentType: 'image/x-wmf' },
      { name: 'image3.png', bytes: emfRectangle(), contentType: 'image/png' },
      { name: 'image4.emf', bytes: emfRectangle(300, 100), contentType: 'image/x-emf' },
    ],
  });

test('a metafile paints the placeholder, then its PNG after one repaint', async () => {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(deck())).ok).toBe(true);
  const session = editor.session!;
  let emits = 0;
  editor.subscribe(() => (emits += 1));
  const generation = session.paintGeneration;
  const fonts = session.fontGeneration;
  const laid = session.layout(0);
  // The first read starts each conversion.
  for (const rId of ['rIdEmf', 'rIdWmf', 'rIdNamed'])
    expect(session.imageUrl(SLIDE, rId)).toBeUndefined();
  await session.picturesReady(0);
  await wait(100);
  expect(emits).toBe(1);
  expect(session.paintGeneration).not.toBe(generation);
  // A converted picture repaints, with no new layout.
  expect(session.fontGeneration).toBe(fonts);
  expect(session.layout(0)).toBe(laid);
  for (const rId of ['rIdEmf', 'rIdWmf', 'rIdNamed'])
    expect(session.imageUrl(SLIDE, rId)).toStartWith('blob:');
  expect(session.picture('/ppt/media/image1.emf')?.mime).toBe('image/png');
  // The natural size is the picture frame.
  expect(session.imageSize(SLIDE, 'rIdWmf')).toEqual({ cx: IN, cy: IN / 2 });
  editor.destroy();
});

test('picturesReady waits for the pictures of one slide, or of every slide', async () => {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(deck())).ok).toBe(true);
  const session = editor.session!;
  await session.picturesReady(0);
  expect(session.picture('/ppt/media/image1.emf')).toBeDefined();
  expect(session.picture('/ppt/media/image4.emf')).toBeUndefined();
  await session.picturesReady();
  expect(session.picture('/ppt/media/image4.emf')).toBeDefined();
  editor.destroy();
});

test('an image export draws a converted metafile as a PNG', async () => {
  const win = new GlobalWindow();
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(deck())).ok).toBe(true);
  const session = editor.session!;
  const doc = win.document as unknown as Document;
  const hrefs = (): string[] => {
    const svg = paintLayout(session, session.layout(0)!, 100, 100, () => false, undefined, doc);
    return [...svg.querySelectorAll('image')].map((i) => i.getAttribute('href') ?? '');
  };
  expect(hrefs()).toHaveLength(0);
  await session.picturesReady(0);
  const after = hrefs();
  expect(after).toHaveLength(3);
  expect(after.every((h) => h.startsWith('data:image/png;base64,'))).toBe(true);
  editor.destroy();
  await win.happyDOM.close();
});

test('only the parts a slide shows count: no notes, other slides, or unused layouts', () => {
  const r = readPptx(deck());
  if (!r.ok) throw new Error('open');
  expect(metafilesOf(r.pkg, [SLIDE])).toHaveLength(3);
  expect(metafilesOf(r.pkg, ['/ppt/slides/slide2.xml'])).toHaveLength(1);
});

test('a listener that throws stops no conversion', async () => {
  let calls = 0;
  const pictures = new MetafilePictures(
    () => [],
    () => {
      calls += 1;
      throw new Error('listener');
    }
  );
  const caught: unknown[] = [];
  const report = globalThis.reportError;
  globalThis.reportError = (e: unknown) => caught.push(e);
  try {
    const a = emfRectangle(100, 50);
    const b = emfRectangle(120, 50);
    await pictures.ready([a]);
    await wait(100);
    await pictures.ready([b]);
    await wait(100);
    expect(calls).toBe(2);
    expect(pictures.drawable(b, 'image/x-emf')?.mime).toBe('image/png');
    expect(caught).toHaveLength(2);
  } finally {
    globalThis.reportError = report;
    pictures.dispose();
  }
});

test('new fonts convert again, and the old PNG draws until then', async () => {
  let fontSets = 0;
  const pictures = new MetafilePictures(
    () => ((fontSets += 1), []),
    () => {}
  );
  const bytes = emfRectangle();
  await pictures.ready([bytes]);
  const first = pictures.drawable(bytes, 'image/x-emf');
  expect(first?.mime).toBe('image/png');
  pictures.fontsChanged();
  // Until the new PNG is ready, the old one draws.
  expect(pictures.drawable(bytes, 'image/x-emf')?.bytes).toBe(first!.bytes);
  await pictures.ready([bytes]);
  expect(fontSets).toBe(2);
  expect(pictures.drawable(bytes, 'image/x-emf')?.bytes).not.toBe(first!.bytes);
  pictures.dispose();
});

test('a wait for pictures ends at its limit', async () => {
  const pictures = new MetafilePictures(
    () => [],
    () => {}
  );
  const started = performance.now();
  await pictures.ready([emfRectangle(4000, 4000)], 1);
  expect(performance.now() - started).toBeLessThan(500);
  pictures.dispose();
});

test('a disposed cache converts nothing and stays quiet', async () => {
  let changed = 0;
  const pictures = new MetafilePictures(
    () => [],
    () => (changed += 1)
  );
  const bytes = emfRectangle();
  expect(pictures.drawable(bytes, 'image/x-emf')).toBeUndefined();
  expect(pictures.busy).toBe(true);
  pictures.dispose();
  await pictures.convert(bytes);
  await wait(200);
  expect(changed).toBe(0);
  expect(pictures.busy).toBe(false);
  expect(pictures.drawable(bytes, 'image/x-emf')).toBeUndefined();
});

test('a metafile that does not convert keeps the placeholder and repaints nothing', async () => {
  let changed = 0;
  const pictures = new MetafilePictures(
    () => [],
    () => (changed += 1)
  );
  const empty = emfRectangle(0, 0);
  await pictures.convert(empty);
  await wait(100);
  expect(changed).toBe(0);
  expect(pictures.drawable(empty, 'image/x-emf')).toBeUndefined();
  expect(pictures.busy).toBe(false);
});
