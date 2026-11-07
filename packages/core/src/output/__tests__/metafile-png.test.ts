// EMF and WMF pictures convert to a PNG on any host, within the size limit.

import { expect, test } from 'bun:test';
import { imageSize } from '@eigenpal/ooxml-core';
import { emfRectangle, wmfRectangle } from '../../../../../scripts/fixtures/picture-bytes.ts';
import { MAX_METAFILE_RASTER_EDGE, metafilePng } from '../metafile-png.ts';

test('an EMF and a WMF convert to a PNG at twice their pixel size', async () => {
  const emf = await metafilePng(emfRectangle(200, 100));
  expect(imageSize(emf!)).toEqual({ width: 400, height: 200 });
  const wmf = await metafilePng(wmfRectangle());
  expect(imageSize(wmf!)).toEqual({ width: 192, height: 96 });
});

test('a large metafile converts within the edge limit', async () => {
  const size = imageSize((await metafilePng(emfRectangle(6000, 3000)))!)!;
  expect(Math.max(size.width, size.height)).toBe(MAX_METAFILE_RASTER_EDGE);
});

test('a stopped conversion gives undefined, so a later one can run', async () => {
  expect(await metafilePng(emfRectangle(), [], AbortSignal.abort())).toBeUndefined();
});

test('other bytes and an empty metafile give null', async () => {
  expect(await metafilePng(new Uint8Array([0x89, 0x50, 0x4e, 0x47]))).toBeNull();
  expect(await metafilePng(emfRectangle(0, 0))).toBeNull();
});

test('damaged headers and records stay within the edge limit', async () => {
  const huge = emfRectangle();
  new DataView(huge.buffer).setInt32(16, 0x7fffffff, true); // bounds
  new DataView(huge.buffer).setInt32(20, 0x7fffffff, true);
  new DataView(huge.buffer).setUint32(92, 0xfffffff0, true); // first record size
  for (const bytes of [huge, emfRectangle().subarray(0, 100), wmfRectangle().subarray(0, 40)]) {
    const png = await metafilePng(bytes);
    const size = png ? imageSize(png) : null;
    expect(Math.max(size?.width ?? 0, size?.height ?? 0)).toBeLessThanOrEqual(
      MAX_METAFILE_RASTER_EDGE
    );
  }
});
