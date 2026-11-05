import { describe, expect, test } from 'bun:test';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

// A JPEG draws its stored pixels and ignores its EXIF orientation (check deck r54-1).
// So Crop > Fill, Reset Picture & Size, and Change Picture without a host size use the stored
// pixels that the header states.

const IN = 914_400;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

/** A big-endian TIFF block with one IFD0 entry: tag 0x0112 (Orientation) = `orientation`. */
function tiff(orientation: number): number[] {
  const u16 = (v: number) => [v >> 8, v & 0xff];
  const u32 = (v: number) => [v >>> 24, (v >> 16) & 0xff, (v >> 8) & 0xff, v & 0xff];
  return [
    0x4d,
    0x4d,
    ...u16(42),
    ...u32(8),
    ...u16(1),
    ...u16(0x0112),
    ...u16(3),
    ...u32(1),
  ].concat([...u16(orientation), 0, 0, ...u32(0)]);
}

/** A JPEG header: SOI, an APP1 `Exif` segment, a SOF0 frame of `w`×`h` stored pixels, EOI. */
function jpeg(w: number, h: number, orientation: number): Uint8Array {
  const body = [0x45, 0x78, 0x69, 0x66, 0, 0, ...tiff(orientation)];
  const app1 = [0xff, 0xe1, (body.length + 2) >> 8, (body.length + 2) & 0xff, ...body];
  const sof = [0xff, 0xc0, 0, 17, 8, h >> 8, h & 0xff, w >> 8, w & 0xff, 3];
  for (let c = 0; c < 9; c++) sof.push(0);
  return new Uint8Array([0xff, 0xd8, ...app1, ...sof, 0xff, 0xd9]);
}

/** A deck with one 4×3 in picture of a JPEG stored 40×30 px, with EXIF orientation 6. */
async function openTurned(): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const shapes =
    '<p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture 2"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
    '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    `<p:spPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${4 * IN}" cy="${3 * IN}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
  const bytes = buildPptx({
    slides: [{ layout: 'blank', rels: [['rIdImg', R_IMAGE, '../media/image1.jpeg']], shapes }],
    media: [{ name: 'image1.jpeg', bytes: jpeg(40, 30, 6), contentType: 'image/jpeg' }],
  });
  expect((await e.load(bytes)).ok).toBe(true);
  e.selectAll();
  return e;
}

describe('picture commands use the stored pixels, not the EXIF orientation', () => {
  test('Crop > Fill takes the stored aspect (4:3): a 4:3 frame needs no crop', async () => {
    const e = await openTurned();
    expect(e.exec({ type: 'cropImage', crop: { kind: 'fill' } }).ok).toBe(true);
    expect(e.snapshot().image!.crop).toEqual({ left: 0, top: 0, right: 0, bottom: 0 });
  });

  test('Reset Picture & Size takes the stored size', async () => {
    const e = await openTurned();
    expect(e.exec({ type: 'resetImage', size: true }).ok).toBe(true);
    const g = e.snapshot().geometry!;
    // 40×30 px at 96 DPI, not turned to 30×40.
    expect([g.cx, g.cy]).toEqual([40 * 9525, 30 * 9525]);
  });

  test('Change Picture without a host size takes the stored aspect of an EXIF JPEG', async () => {
    const e = await openTurned();
    const replaced = e.exec({ type: 'replaceImage', bytes: jpeg(80, 60, 8), mime: 'image/jpeg' });
    expect(replaced.ok).toBe(true);
    const g = e.snapshot().geometry!;
    // A 4:3 image fills the 4×3 in frame; turned, it would fit as 2.25×3 in.
    expect([g.cx, g.cy]).toEqual([4 * IN, 3 * IN]);
  });
});
