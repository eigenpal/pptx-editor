/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Embedded pictures: a failed embed, and smooth sampling of the color and the alpha.

import { describe, expect, spyOn, test } from 'bun:test';
import { PNG } from 'pngjs';
import { PDFBool, PDFDocument, PDFImage, PDFName, PDFRawStream } from 'pdf-lib';
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';
import { exportPdf } from '../index.ts';

const IN = 914_400;
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function png(alpha: number): Uint8Array {
  const image = new PNG({ width: 2, height: 2 });
  image.data.fill(200);
  for (let i = 3; i < image.data.length; i += 4) image.data[i] = alpha;
  return PNG.sync.write(image);
}

const deck = (bytes: Uint8Array): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          `<p:pic><p:nvPicPr><p:cNvPr id="4" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
          `<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
          `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${IN}" cy="${IN}"/></a:xfrm>` +
          `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`,
        rels: [['rIdImg', `${R}/image`, '../media/img.png']],
      },
    ],
    media: [{ name: 'img.png', bytes, contentType: 'image/png' }],
  });

/** The image XObjects of a PDF. */
async function images(bytes: Uint8Array): Promise<PDFRawStream[]> {
  const doc = await PDFDocument.load(bytes);
  const out: PDFRawStream[] = [];
  for (const [, obj] of doc.context.enumerateIndirectObjects())
    if (obj instanceof PDFRawStream && obj.dict.get(PDFName.of('Subtype')) === PDFName.of('Image'))
      out.push(obj);
  return out;
}

describe('embedded pictures', () => {
  test('a picture that fails to embed is reported and left out, and the export completes', async () => {
    const embed = spyOn(PDFImage.prototype, 'embed').mockImplementation(async () => {
      throw new Error('broken image data');
    });
    try {
      const result = await exportPdf(deck(png(255)), { fidelityPolicy: 'best-effort' });
      expect(result.diagnostics.map((d) => d.code)).toContain('picture.unreadable');
      expect(await images(result.bytes)).toEqual([]);
    } finally {
      embed.mockRestore();
    }
  });

  test('a transparent PNG samples its color and its alpha smoothly', async () => {
    const result = await exportPdf(deck(png(128)));
    const found = await images(result.bytes);
    // The color image and its soft mask.
    expect(found).toHaveLength(2);
    for (const image of found) expect(image.dict.get(PDFName.of('Interpolate'))).toBe(PDFBool.True);
  });
});
