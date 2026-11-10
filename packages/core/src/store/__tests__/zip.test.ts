// Tests for the bounded ZIP reader and the deterministic writer.
//
// The first two blocks are adapted from
// docx-editor/packages/core/src/store/__tests__/zip-proto.test.ts and
// docx-editor/packages/core/src/store/package/__tests__/zip-determinism.test.ts at 26095c632
// (EigenPal, Apache-2.0), with WordprocessingML part names replaced by PresentationML ones. The
// rest are regression tests for audit findings S1, S6, K1, K4, and P6.

import { describe, expect, test } from 'bun:test';
import { deflateSync, strToU8, unzipSync, zipSync } from 'fflate';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import { normalizePartName } from '@eigenpal/ooxml-core';
import { readPptx, writePptx } from '../package/package.ts';
import { DEFAULT_ZIP_LIMITS, readZip, writeZip, type ZipLimits } from '@eigenpal/ooxml-core';

describe('dangerous part-name keys fail closed', () => {
  test('normalizePartName rejects __proto__ / constructor / prototype segments', () => {
    for (const n of ['/__proto__', '/ppt/constructor', '/prototype']) {
      const r = normalizePartName(n);
      expect(r.ok).toBe(false);
    }
    expect(normalizePartName('/ppt/presentation.xml').ok).toBe(true);
  });

  test('writeZip fails closed on such a part and does not pollute Object.prototype', () => {
    expect(() =>
      writeZip(
        new Map([
          ['/ppt/presentation.xml', new Uint8Array([1])],
          ['/__proto__', new Uint8Array([9])],
        ])
      )
    ).toThrow(/unsafe part name/);
    expect(({} as Record<string, unknown>).polluted).toBeUndefined();
  });
});

const parts = new Map([
  ['/ppt/presentation.xml', new TextEncoder().encode('<p:presentation/>')],
  ['/[Content_Types].xml', new TextEncoder().encode('<Types/>')],
]);

/** Local file header: signature(4) version(2) flags(2) method(2), then the DOS time and date. */
const METHOD_OFFSET = 8;
const DOS_TIME_OFFSET = 10;

describe('writeZip is deterministic', () => {
  test('the entry timestamp does not come from the clock', () => {
    const now = Date.now;
    try {
      const first = writeZip(parts);
      // Far enough that any clock-derived field would have to move.
      Date.now = () => now() + 7 * 24 * 60 * 60 * 1000;
      const second = writeZip(parts);
      expect([...second]).toEqual([...first]);
    } finally {
      Date.now = now;
    }
  });

  test('it is the DOS epoch, so the year field cannot underflow', () => {
    const bytes = writeZip(parts);
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const date = view.getUint16(DOS_TIME_OFFSET + 2, true);
    // DOS packs the year into the top seven bits, counted from 1980.
    expect(date >>> 9).toBe(0);
  });
});

/**
 * An archive with `records` central-directory records that all point at ONE stored local entry
 * of `payload` bytes, each declaring an uncompressed size of `declared`.
 */
function overlappingStoredZip(records: number, payload: number, declared: number): Uint8Array {
  const local = new Uint8Array(30 + 1 + payload);
  const lv = new DataView(local.buffer);
  lv.setUint32(0, 0x04034b50, true);
  lv.setUint32(18, payload, true);
  lv.setUint32(22, declared, true);
  lv.setUint16(26, 1, true);
  local[30] = 0x61;
  const directory: Uint8Array[] = [];
  for (let i = 0; i < records; i += 1) {
    const name = strToU8(`f${i}.bin`);
    const cd = new Uint8Array(46 + name.length);
    const cv = new DataView(cd.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint32(20, payload, true);
    cv.setUint32(24, declared, true);
    cv.setUint16(28, name.length, true);
    cv.setUint32(42, 0, true);
    cd.set(name, 46);
    directory.push(cd);
  }
  const cdLength = directory.reduce((n, c) => n + c.length, 0);
  const out = new Uint8Array(local.length + cdLength + 22);
  out.set(local, 0);
  let p = local.length;
  for (const cd of directory) {
    out.set(cd, p);
    p += cd.length;
  }
  const ev = new DataView(out.buffer, p, 22);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, records, true);
  ev.setUint16(10, records, true);
  ev.setUint32(12, cdLength, true);
  ev.setUint32(16, local.length, true);
  return out;
}

describe('S1: stored entries cannot bypass the size caps', () => {
  test('directory records that share one stored entry are refused', () => {
    // 1 MiB of input that used to inflate to 200 MiB under a 1 MiB cap.
    const zip = overlappingStoredZip(200, 1024 * 1024, 0);
    const r = readZip(zip, { maxEntries: 10_000, maxTotalBytes: 1024 * 1024, maxRatio: 200 });
    expect(r).toMatchObject({ ok: false, reason: 'overlapping-entries' });
  });

  test('with the default limits too', () => {
    const r = readZip(overlappingStoredZip(600, 1024 * 1024, 0));
    expect(r).toMatchObject({ ok: false, reason: 'overlapping-entries' });
  });

  test('a stored entry counts its real size, not the size it declares', () => {
    const zip = overlappingStoredZip(1, 2 * 1024 * 1024, 0);
    expect(readZip(zip, { maxEntries: 10, maxTotalBytes: 1024 * 1024 })).toMatchObject({
      ok: false,
      reason: 'too-large',
      limit: 'zip.maxTotalBytes',
    });
    expect(readZip(zip, { maxEntries: 10, maxTotalBytes: 4 * 1024 * 1024 }).ok).toBe(true);
  });

  test('an entry whose data runs past the end of the input is refused', () => {
    const zip = zipSync({ 'a.bin': new Uint8Array(1000) }, { level: 0 });
    const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
    const cd = zip.length - 22 - (46 + 5);
    view.setUint32(cd + 20, 5000, true);
    expect(readZip(zip).ok).toBe(false);
  });

  test('ordinary archives still read', () => {
    const zip = zipSync({ 'a.xml': strToU8('<a/>'), 'b/c.bin': new Uint8Array(10) });
    const r = readZip(zip);
    expect(r.ok && [...r.entries.keys()]).toEqual(['/a.xml', '/b/c.bin']);
  });
});

describe('S6: limits are validated, clamped, and frozen', () => {
  const zip = zipSync({ 'a.bin': new Uint8Array(1000), 'b.bin': new Uint8Array(1000) });

  test('a limit that is not a finite, non-negative number refuses instead of turning caps off', () => {
    const bad: ZipLimits[] = [
      { maxEntries: Number.NaN, maxTotalBytes: 1e9 },
      { maxEntries: 10, maxTotalBytes: Number.NaN },
      { maxEntries: 10, maxTotalBytes: 1e9, maxRatio: Number.NaN },
      { maxEntries: -1, maxTotalBytes: 1e9 },
      { maxEntries: 10, maxTotalBytes: Number.POSITIVE_INFINITY },
      { maxEntries: 10, maxTotalBytes: 1e9, maxRatio: -1 },
      { maxEntries: 1.5, maxTotalBytes: 1e9 },
    ];
    for (const limits of bad)
      expect(readZip(zip, limits)).toMatchObject({ ok: false, reason: 'invalid-limits' });
    expect(readZip(zip, null as unknown as ZipLimits)).toMatchObject({
      ok: false,
      reason: 'invalid-limits',
    });
  });

  test('a missing field takes the default, so the caps stay on', () => {
    expect(readZip(zip, {} as ZipLimits).ok).toBe(true);
    expect(readZip(zip, { maxEntries: 1 } as ZipLimits)).toMatchObject({
      ok: false,
      reason: 'too-many-entries',
    });
  });

  test('the defaults cannot be changed by a host', () => {
    expect(Object.isFrozen(DEFAULT_ZIP_LIMITS)).toBe(true);
    expect(() => {
      (DEFAULT_ZIP_LIMITS as { maxTotalBytes: number }).maxTotalBytes = Number.MAX_VALUE;
    }).toThrow();
  });
});

describe('K4: where the ratio cap starts, as in DOCX Editor', () => {
  test('a deck with a 6 MB blank bitmap opens (about 1022:1 for that entry)', () => {
    const bmp = new Uint8Array(54 + 1920 * 1080 * 3);
    bmp.fill(0xff, 54);
    bmp[0] = 0x42;
    bmp[1] = 0x4d;
    const deck = buildPptx({
      slides: [{ shapes: shape({ id: 2, x: 0, y: 0, cx: 100, cy: 100 }) }],
      media: [{ name: 'blank.bmp', bytes: bmp, contentType: 'image/bmp' }],
    });
    const r = readPptx(deck);
    expect(r.ok).toBe(true);
  });

  test('a small deck with a larger blank bitmap is refused above 8 MiB', () => {
    const bmp = new Uint8Array(54 + 2560 * 1440 * 3);
    bmp.fill(0xff, 54);
    bmp[0] = 0x42;
    bmp[1] = 0x4d;
    const deck = buildPptx({
      slides: [{ shapes: shape({ id: 2, x: 0, y: 0, cx: 100, cy: 100 }) }],
      media: [{ name: 'blank.bmp', bytes: bmp, contentType: 'image/bmp' }],
    });
    expect(readZip(deck)).toMatchObject({ ok: false, reason: 'too-large', limit: 'zip.maxRatio' });
  });

  test('a deck that opens with a blank bitmap stored saves to a deck that reopens', () => {
    const bmp = new Uint8Array(54 + 2560 * 1440 * 3);
    bmp.fill(0xff, 54);
    bmp[0] = 0x42;
    bmp[1] = 0x4d;
    const built = unzipSync(
      buildPptx({
        slides: [{ shapes: shape({ id: 2, x: 0, y: 0, cx: 100, cy: 100 }) }],
        media: [{ name: 'blank.bmp', bytes: bmp, contentType: 'image/bmp' }],
      })
    );
    // The same deck with the bitmap stored, as some writers save it: about 1:1, so it opens.
    const stored = zipSync(
      Object.fromEntries(
        Object.entries(built).map(([name, bytes]) => [
          name,
          name.endsWith('.bmp') ? ([bytes, { level: 0 }] as const) : bytes,
        ])
      ) as Parameters<typeof zipSync>[0]
    );
    const opened = readPptx(stored);
    if (!opened.ok) throw new Error('fixture');
    const reopened = readPptx(writePptx(opened.pkg));
    expect(reopened.ok).toBe(true);
    // It deflates an 11 MB part twice; give a loaded machine room.
  }, 30_000);

  test('a highly compressible part opens under a strict ratio', () => {
    const xml = strToU8('<a:p><a:r><a:t>x</a:t></a:r></a:p>'.repeat(4000));
    const zip = zipSync({ 'a.xml': xml });
    expect(xml.length / zip.length).toBeGreaterThan(200);
    expect(readZip(zip, { maxEntries: 10, maxTotalBytes: 1e9, maxRatio: 2 }).ok).toBe(true);
  });

  test('an archive that claims to inflate far beyond the ratio is refused before inflation', () => {
    // A deflated entry whose directory record declares 200 MiB: refused from the header alone.
    const zip = zipSync({ 'bomb.bin': new Uint8Array(1024) });
    const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
    const cd = zip.length - 22 - (46 + 'bomb.bin'.length);
    view.setUint32(cd + 24, 200 * 1024 * 1024, true);
    expect(readZip(zip)).toMatchObject({ ok: false, reason: 'too-large', limit: 'zip.maxRatio' });
  });
});

describe('K1: part names are decoded once', () => {
  test('a name with an encoded percent sign saves under the same name', () => {
    const zip = zipSync({
      'a%2541.xml': strToU8('<a/>'),
      'b%25zz.bin': strToU8('x'),
      'ppt/media/image%201.png': strToU8('y'),
    });
    const first = readZip(zip);
    expect(first.ok && [...first.entries.keys()]).toEqual([
      '/a%41.xml',
      '/b%zz.bin',
      '/ppt/media/image 1.png',
    ]);
    if (!first.ok) return;
    const saved = writeZip(first.entries);
    expect(Object.keys(unzipSync(saved))).toEqual([
      'a%2541.xml',
      'b%25zz.bin',
      'ppt/media/image%201.png',
    ]);
    const again = readZip(saved);
    expect(again.ok && [...again.entries.keys()]).toEqual([...first.entries.keys()]);
  });
});

describe('P6: compressed media is stored, not deflated again', () => {
  const PNG = Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82]);

  function methodOf(zip: Uint8Array, name: string): number {
    const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
    const target = strToU8(name);
    for (let p = 0; p + 30 < zip.length; p += 1) {
      if (view.getUint32(p, true) !== 0x04034b50) continue;
      const length = view.getUint16(p + 26, true);
      const entry = zip.subarray(p + 30, p + 30 + length);
      if (length === target.length && entry.every((b, i) => b === target[i]))
        return view.getUint16(p + METHOD_OFFSET, true);
    }
    throw new Error(`no entry ${name}`);
  }

  test('PNG and JPEG are stored; XML is deflated; the package reads back the same', () => {
    const jpeg = Uint8Array.from([0xff, 0xd8, 0xff, 0xe0, 1, 2, 3]);
    const deck = buildPptx({
      slides: [
        {
          shapes: '',
          ext: '<p:extLst><p:ext uri="urn:test" r:id="png"/><p:ext uri="urn:test:jpeg" r:id="jpeg"/></p:extLst>',
          rels: [
            [
              'png',
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
              '../media/image1.png',
            ],
            [
              'jpeg',
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
              '../media/image2.jpeg',
            ],
          ],
        },
      ],
      media: [
        { name: 'image1.png', bytes: PNG, contentType: 'image/png' },
        { name: 'image2.jpeg', bytes: jpeg, contentType: 'image/jpeg' },
      ],
    });
    const pkg = readPptx(deck);
    if (!pkg.ok) throw new Error('fixture');
    const saved = writePptx(pkg.pkg);
    expect(methodOf(saved, 'ppt/media/image1.png')).toBe(0);
    expect(methodOf(saved, 'ppt/media/image2.jpeg')).toBe(0);
    expect(methodOf(saved, 'ppt/presentation.xml')).toBe(8);
    const reopened = readPptx(saved);
    expect(reopened.ok).toBe(true);
    if (!reopened.ok) return;
    const media = reopened.pkg.parts.get('/ppt/media/image1.png');
    expect(media?.kind === 'binary' && [...media.bytes]).toEqual([...PNG]);
  });

  test('a deflated archive still reads when an entry is not compressible', () => {
    const random = deflateSync(strToU8('seed'));
    expect(readZip(zipSync({ 'x.bin': random })).ok).toBe(true);
  });
});
