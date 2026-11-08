// The session's image size port: the header's
// resolution, or 72 DPI when the header states none.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { openPresentation } from '../index.ts';

const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(join(import.meta.dir, '../../../../../fixtures', name)));

describe('DocumentSession.imageSize', () => {
  test('a PNG without a resolution is 72 DPI', async () => {
    const r = await openPresentation(fixture('pictures.pptx'), { fixedMeasurer: true });
    if (!r.ok) throw new Error('open failed');
    // 128 × 96 px: 1.778 × 1.333 in at 72 DPI.
    expect(r.session.imageSize('/ppt/slides/slide1.xml', 'rIdImg')).toEqual({
      cx: 1_625_600,
      cy: 1_219_200,
    });
    r.session.dispose();
  });

  test('a stated resolution is kept', async () => {
    const r = await openPresentation(fixture('fill-fidelity.pptx'), { fixedMeasurer: true });
    if (!r.ok) throw new Error('open failed');
    const part = '/ppt/slides/slide3.xml';
    // The fixture states 144 and 150 DPI as 5669 and 5906 pixels per metre (`pHYs`), which
    // read back as 143.99 and 150.01 DPI.
    for (const [rel, ppm] of [
      ['rIdImg', 5669],
      ['rIdDpi', 5906],
    ] as const) {
      const dpi = ppm * 0.0254;
      expect(r.session.imageSize(part, rel)).toEqual({
        cx: Math.round((128 * 914_400) / dpi),
        cy: Math.round((96 * 914_400) / dpi),
      });
    }
    expect(r.session.imageSize(part, 'rIdMissing')).toBeUndefined();
    r.session.dispose();
  });
});
