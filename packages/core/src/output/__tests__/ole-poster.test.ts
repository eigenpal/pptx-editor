// An OLE object shows the picture of its `p:oleObj/p:pic` (ECMA-376 Part 1, 19.3.1.36) in the
// frame's box. The object is never run. With no picture, the frame keeps its labeled placeholder.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { buildSlideStructureDeck } from '../../../../../scripts/fixtures/slide-structure.ts';
import { resolveSlide, type ResolvedGraphicFrame } from '../../cascade/index.ts';
import { createFixedMeasurer, layoutSlide, type SlideLayout } from '../../layout/index.ts';
import { readPptx } from '../../store/index.ts';
import { paintSlide } from '../index.ts';

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const IN = 914_400;
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg=='
  ),
  (c) => c.charCodeAt(0)
);

function layoutOf(bytes: Uint8Array, index: number): SlideLayout {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('open failed');
  return layoutSlide(resolveSlide(r.pkg, index)!, createFixedMeasurer());
}

function paint(layout: SlideLayout): SVGSVGElement {
  return paintSlide(doc, layout, {
    images: (part, rId) => `blob:${part}#${rId}`,
    idPrefix: 't',
  });
}

const xfrm = (x: number, y: number, cx: number, cy: number, tag = 'a:xfrm'): string =>
  `<${tag}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></${tag}>`;

/** An OLE frame at (1 in, 2 in), 3 in by 2 in. `pic` is the poster's own box, or none. */
function oleFrame(pic: 'same' | 'elsewhere' | 'none', wrap = false): string {
  const box = pic === 'elsewhere' ? xfrm(6 * IN, 4 * IN, IN, IN) : xfrm(IN, 2 * IN, 3 * IN, 2 * IN);
  const poster =
    pic === 'none'
      ? ''
      : '<p:pic><p:nvPicPr><p:cNvPr id="0" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
        '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
        `<p:spPr>${box}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  const ole = (body: string): string =>
    `<p:oleObj name="Worksheet" r:id="rIdOle" progId="Package"><p:embed/>${body}</p:oleObj>`;
  const data = wrap
    ? '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
      `<mc:Choice xmlns:v="urn:schemas-microsoft-com:vml" Requires="v">${ole('')}</mc:Choice>` +
      `<mc:Fallback>${ole(poster)}</mc:Fallback></mc:AlternateContent>`
    : ole(poster);
  return (
    '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Object 2"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
    xfrm(IN, 2 * IN, 3 * IN, 2 * IN, 'p:xfrm') +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">${data}</a:graphicData></a:graphic></p:graphicFrame>`
  );
}

/** A PNG header that states `w` × `h` pixels; only the header is read. */
function pngHeader(w: number, h: number): Uint8Array {
  const b = PNG.slice();
  const v = new DataView(b.buffer);
  v.setUint32(16, w);
  v.setUint32(20, h);
  return b;
}

type ImageRel = readonly [string, string, string] | readonly [string, string, string, 'External'];

/** A deck with one slide. `image` is the poster relationship's target and the part's bytes. */
function deck(
  shapes: string,
  image: { readonly rel?: ImageRel; readonly bytes?: Uint8Array; readonly name?: string } = {}
): Uint8Array {
  const name = image.name ?? 'image1.png';
  return buildPptx({
    slides: [
      {
        shapes,
        layout: 'blank',
        rels: [
          image.rel ?? ['rIdImg', `${R}/image`, `../media/${name}`],
          ['rIdOle', `${R}/package`, '../embeddings/object1.bin'],
        ],
      },
    ],
    media: [{ name, bytes: image.bytes ?? PNG, contentType: 'image/png' }],
    binaryParts: [
      {
        name: 'ppt/embeddings/object1.bin',
        bytes: new Uint8Array([1, 2, 3]),
        contentType: 'application/vnd.openxmlformats-officedocument.oleObject',
      },
    ],
  });
}

function frameOf(layout: SlideLayout): ResolvedGraphicFrame {
  const item = layout.items.find((i) => i.kind === 'frame');
  if (item?.kind !== 'frame') throw new Error('no frame');
  return item.element;
}

describe('OLE poster picture (ECMA-376 Part 1, 19.3.1.36)', () => {
  test('the slide-structure fixture paints the poster, not a placeholder', () => {
    const layout = layoutOf(buildSlideStructureDeck(), 3);
    const frame = frameOf(layout);
    expect(frame.frameKind).toBe('ole');
    expect(frame.poster?.fill.type).toBe('picture');
    expect(layout.diagnostics.some((d) => d.startsWith('graphic frame rendered'))).toBe(false);
    const root = paint(layout);
    const g = root.querySelector('[data-kind="frame"]')!;
    expect(g.getAttribute('data-node-id')).toBe(frame.nodeId);
    expect(g.querySelector('image')?.getAttribute('href')).toContain('rIdImg');
    // No placeholder label and no dashed box.
    expect(g.querySelector('text')).toBeNull();
    expect(g.querySelector('rect[stroke-dasharray]')).toBeNull();
    // Only the frame is a hit target.
    expect(root.querySelectorAll(`[data-node-id="${frame.nodeId}"]`).length).toBe(1);
  });

  test('the poster fills the frame box, whatever box the picture states', () => {
    const layout = layoutOf(deck(oleFrame('elsewhere')), 0);
    const item = layout.items.find((i) => i.kind === 'frame');
    if (item?.kind !== 'frame' || !item.poster) throw new Error('no poster');
    expect(item.poster.element.xfrm).toMatchObject({ x: IN, y: 2 * IN, cx: 3 * IN, cy: 2 * IN });
    expect(item.poster.matrix).toEqual(item.matrix);
  });

  test('an mc:AlternateContent takes the picture of its fallback branch', () => {
    const layout = layoutOf(deck(oleFrame('same', true)), 0);
    expect(frameOf(layout).poster?.fill.type).toBe('picture');
  });

  test('an OLE object with no picture keeps the labeled placeholder', () => {
    const layout = layoutOf(deck(oleFrame('none')), 0);
    expect(frameOf(layout).poster).toBeUndefined();
    expect(layout.diagnostics).toContain('graphic frame rendered as placeholder: ole');
    const g = paint(layout).querySelector('[data-kind="frame"]')!;
    expect(g.querySelector('text')?.textContent).toBe('ole');
  });

  // Every poster value is untrusted: a picture that cannot be drawn keeps the placeholder.
  test.each([
    [
      'an external target',
      { rel: ['rIdImg', `${R}/image`, 'https://example.com/a.png', 'External'] },
    ],
    ['a missing part', { rel: ['rIdImg', `${R}/image`, '../media/missing.png'] }],
    [
      'a relationship to another part',
      { rel: ['rIdImg', `${R}/image`, '../embeddings/object1.bin'] },
    ],
    ['bytes that are not an image', { bytes: new Uint8Array([0x01, 0x00, 0x00, 0x00, 0x6c]) }],
    ['an EMF picture', { name: 'image1.emf', bytes: new Uint8Array(44) }],
    ['a picture of 10 000 × 10 000 pixels', { bytes: pngHeader(10_000, 10_000) }],
    ['a picture of zero pixels', { bytes: pngHeader(0, 1) }],
  ] as const)('%s gives no poster, so the placeholder shows', (_, image) => {
    const layout = layoutOf(deck(oleFrame('same'), image), 0);
    expect(frameOf(layout).poster).toBeUndefined();
    expect(layout.diagnostics).toContain('graphic frame rendered as placeholder: ole');
  });

  test('a poster in a group follows the group transform', () => {
    const group =
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${6 * IN}" cy="${4 * IN}"/>` +
      `<a:chOff x="0" y="0"/><a:chExt cx="${12 * IN}" cy="${8 * IN}"/></a:xfrm></p:grpSpPr>` +
      `${oleFrame('same')}</p:grpSp>`;
    const layout = layoutOf(deck(group), 0);
    const g = layout.items[0];
    if (g?.kind !== 'group') throw new Error('no group');
    const item = g.children[0];
    if (item?.kind !== 'frame' || !item.poster) throw new Error('no poster');
    expect(item.poster.matrix).toEqual(item.matrix);
    expect(item.poster.element.xfrm.cx).toBe(1.5 * IN);
  });
});
