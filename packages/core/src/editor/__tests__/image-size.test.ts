// The session's image size port: the header's
// resolution, or 72 DPI when the header states none.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { openPresentation } from '../index.ts';

const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(join(import.meta.dir, '../../../../../fixtures', name)));

describe('DocumentSession.imageSize', () => {
  test('a PNG without a resolution is 72 DPI; a stated 150 DPI is kept', async () => {
    const r = await openPresentation(fixture('fill-fidelity.pptx'), { fixedMeasurer: true });
    if (!r.ok) throw new Error('open failed');
    const part = '/ppt/slides/slide3.xml';
    // 128 × 96 px: 1.778 × 1.333 in at 72 DPI. The fixture states 150 DPI as 5906 pixels per
    // metre (`pHYs`), which reads back as 150.0124 DPI.
    expect(r.session.imageSize(part, 'rIdImg')).toEqual({ cx: 1_625_600, cy: 1_219_200 });
    const dpi = 5906 * 0.0254;
    expect(r.session.imageSize(part, 'rIdDpi')).toEqual({
      cx: Math.round((128 * 914_400) / dpi),
      cy: Math.round((96 * 914_400) / dpi),
    });
    expect(r.session.imageSize(part, 'rIdMissing')).toBeUndefined();
    r.session.dispose();
  });
});
