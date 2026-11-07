// A synthetic deck of large pictures for the browser memory bench (`bun run
// bench:memory:browser`): each slide holds one full-slide picture, its own media part, so the
// browser decodes a different bitmap on every slide.

import { strToU8, zlibSync } from 'fflate';
import { buildPptx } from '../fixtures/builder.ts';

const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

const CRC = Array.from({ length: 256 }, (_, n) => {
  let c = n;
  for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});

function crc32(bytes: Uint8Array): number {
  let c = 0xffffffff;
  for (const b of bytes) c = CRC[(c ^ b) & 0xff]! ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(type: string, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(12 + data.length);
  const view = new DataView(out.buffer);
  view.setUint32(0, data.length);
  out.set(strToU8(type), 4);
  out.set(data, 8);
  view.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)));
  return out;
}

/**
 * A `w`×`h` RGB PNG of gradient bands that `seed` shifts, so each picture has its own pixels. The
 * gradient compresses well: the file is small, but the decoded bitmap is `w`×`h`×4 bytes.
 */
export function gradientPng(w: number, h: number, seed: number): Uint8Array {
  const raw = new Uint8Array((w * 3 + 1) * h);
  for (let y = 0; y < h; y += 1) {
    const row = y * (w * 3 + 1);
    for (let x = 0; x < w; x += 1) {
      const at = row + 1 + x * 3;
      raw[at] = (x + seed * 37) & 0xff;
      raw[at + 1] = ((y >> 6) * 13 + seed * 11) & 0xff;
      raw[at + 2] = (seed * 53) & 0xff;
    }
  }
  const ihdr = new Uint8Array(13);
  const view = new DataView(ihdr.buffer);
  view.setUint32(0, w);
  view.setUint32(4, h);
  ihdr.set([8, 2, 0, 0, 0], 8);
  const parts = [
    new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', zlibSync(raw, { level: 6 })),
    chunk('IEND', new Uint8Array()),
  ];
  const png = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const p of parts) (png.set(p, at), (at += p.length));
  return png;
}

/** A deck of `n` slides, each with one full-slide `w`×`h` picture of its own. */
export function pictureDeck(n: number, w = 1920, h = 1080): Uint8Array {
  const slides = Array.from({ length: n }, (_, i) => ({
    shapes:
      `<p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture ${i + 1}" descr="Picture ${i + 1}"/>` +
      '<p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>',
    rels: [['rIdImg', R_IMAGE, `../media/picture${i + 1}.png`] as const],
  }));
  const media = Array.from({ length: n }, (_, i) => ({
    name: `picture${i + 1}.png`,
    bytes: gradientPng(w, h, i + 1),
    contentType: 'image/png',
  }));
  return buildPptx({ slides, media });
}
