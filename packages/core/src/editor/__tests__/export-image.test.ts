// Slide pictures: the standalone SVG holds its pictures as data URLs (never a blob URL that
// dies with the page), has the asked pixel size, and refuses sizes out of range. The PNG path
// needs a real canvas: the e2e suite checks it in Chromium.

import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { createPptxEditor } from '../index.ts';
import { exportSlideSvg, sfntFsType } from '../export-image.ts';
import { CommandRefusal } from '../../automation/commands/op-batch.ts';

// The engine run has no DOM: lend this file happy-dom's document only while it runs.
const win = new GlobalWindow();
const g = globalThis as unknown as Record<string, unknown>;
const saved = { document: g.document, XMLSerializer: g.XMLSerializer };
beforeAll(() => {
  g.document = win.document;
  g.XMLSerializer = win.XMLSerializer;
});
afterAll(() => {
  g.document = saved.document;
  g.XMLSerializer = saved.XMLSerializer;
  void win.happyDOM.close();
});

const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='
  ),
  (c) => c.charCodeAt(0)
);
const PIC =
  '<p:pic><p:nvPicPr><p:cNvPr id="5" name="Picture 4"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
  '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
  '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';

async function open() {
  const bytes = buildPptx({
    slides: [
      {
        shapes: PIC,
        rels: [
          [
            'rIdImg',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            '../media/image1.png',
          ],
        ],
      },
    ],
    media: [{ name: 'image1.png', bytes: PNG, contentType: 'image/png' }],
  });
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e.session!;
}

describe('exportSlideSvg', () => {
  test('pictures are inline data URLs at the asked pixel size', async () => {
    const session = await open();
    const svg = exportSlideSvg(session, 0, { width: 320 });
    expect(svg.startsWith('<?xml')).toBe(true);
    expect(svg).toContain('data:image/png;base64,');
    expect(svg).not.toContain('blob:');
    expect(svg).toContain('width="320"');
    expect(svg).toContain('height="180"');
  });

  test('a size out of range is refused with a typed reason', async () => {
    const session = await open();
    let reason = '';
    try {
      exportSlideSvg(session, 0, { width: 4 });
    } catch (e) {
      reason = e instanceof CommandRefusal ? e.reason : 'other';
    }
    expect(reason).toBe('errors.exportSize');
  });
});

describe('sfntFsType', () => {
  test('reads the OS/2 fsType and rejects a cut table', () => {
    const bytes = new Uint8Array(28 + 10);
    const view = new DataView(bytes.buffer);
    view.setUint16(4, 1);
    view.setUint32(12, 0x4f532f32);
    view.setUint32(20, 28);
    view.setUint16(28 + 8, 0x0008);
    expect(sfntFsType(bytes)).toBe(8);
    expect(sfntFsType(bytes.subarray(0, 30))).toBeUndefined();
    expect(sfntFsType(new Uint8Array(4))).toBeUndefined();
  });
});
