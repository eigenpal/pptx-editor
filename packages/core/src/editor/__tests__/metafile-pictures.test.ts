// EMF and WMF pictures paint the placeholder until their PNG is ready, then the editor repaints
// and the editor, the image export, and the natural size use the PNG.

import { expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { emfRectangle, wmfRectangle } from '../../../../../scripts/fixtures/picture-bytes.ts';
import { paintLayout } from '../export-image.ts';
import { createPptxEditor } from '../index.ts';
import { MetafilePictures } from '../metafile-pictures.ts';

const IN = 914_400;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const SLIDE = '/ppt/slides/slide1.xml';

const pic = (id: number, rId: string): string =>
  `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="P${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
  `<p:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  `<p:spPr><a:xfrm><a:off x="${id * IN}" y="${IN}"/><a:ext cx="${IN}" cy="${IN}"/></a:xfrm>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';

/** A deck with an EMF, a WMF, and an EMF in a part named `.png`. */
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
    ],
    media: [
      { name: 'image1.emf', bytes: emfRectangle(), contentType: 'image/x-emf' },
      { name: 'image2.wmf', bytes: wmfRectangle(), contentType: 'image/x-wmf' },
      { name: 'image3.png', bytes: emfRectangle(), contentType: 'image/png' },
    ],
  });

test('a metafile paints the placeholder, then its PNG after one repaint', async () => {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(deck())).ok).toBe(true);
  const session = editor.session!;
  let emits = 0;
  editor.subscribe(() => (emits += 1));
  const generation = session.fontGeneration;
  // The first read starts each conversion.
  for (const rId of ['rIdEmf', 'rIdWmf', 'rIdNamed'])
    expect(session.imageUrl(SLIDE, rId)).toBeUndefined();
  await session.picturesReady();
  expect(emits).toBe(1);
  expect(session.fontGeneration).toBe(generation + 1);
  for (const rId of ['rIdEmf', 'rIdWmf', 'rIdNamed'])
    expect(session.imageUrl(SLIDE, rId)).toStartWith('blob:');
  expect(session.picture('/ppt/media/image1.emf')?.mime).toBe('image/png');
  // The natural size is the picture frame.
  expect(session.imageSize(SLIDE, 'rIdWmf')).toEqual({ cx: IN, cy: IN / 2 });
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
  await session.picturesReady();
  const after = hrefs();
  expect(after).toHaveLength(3);
  expect(after.every((h) => h.startsWith('data:image/png;base64,'))).toBe(true);
  editor.destroy();
  await win.happyDOM.close();
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
  await new Promise((r) => setTimeout(r, 200));
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
  expect(changed).toBe(0);
  expect(pictures.drawable(empty, 'image/x-emf')).toBeUndefined();
  expect(pictures.busy).toBe(false);
});
