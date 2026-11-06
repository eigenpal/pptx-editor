import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '../../store/index.ts';
import { resolveSlide } from '../../cascade/index.ts';
import { createFixedMeasurer, layoutSlide } from '../../layout/index.ts';
import { paintSlide } from '../index.ts';

// Picture paint: the exact `a:biLevel` threshold (ECMA-376 Part 1, 20.1.8.11) and the outline
// drawn outside a picture, as reference PDFs show, with its dash rhythm.

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

function paint(blip: string, spPr = '', srcRect = ''): SVGSVGElement {
  const r = readPptx(
    buildPptx({
      slides: [
        {
          layout: 'blank',
          rels: [['rIdImg', R_IMAGE, '../media/image1.png']],
          shapes:
            '<p:pic><p:nvPicPr><p:cNvPr id="2" name="P"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
            `<p:blipFill><a:blip r:embed="rIdImg">${blip}</a:blip>${srcRect}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
            '<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>' +
            `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${spPr}</p:spPr></p:pic>`,
        },
      ],
      media: [
        { name: 'image1.png', bytes: new Uint8Array([137, 80, 78, 71]), contentType: 'image/png' },
      ],
    })
  );
  if (!r.ok) throw new Error('open failed');
  const layout = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer());
  return paintSlide(doc, layout, { images: () => 'blob:test', idPrefix: 't' });
}

describe('picture paint', () => {
  test('black and white thresholds at 25 %, 50 %, and 75 %, not in 10 % steps', () => {
    for (const thresh of [25_000, 50_000, 75_000]) {
      const svg = paint(`<a:biLevel thresh="${thresh}"/>`);
      const fn = svg.querySelector('feComponentTransfer feFuncR')!;
      expect(fn.getAttribute('type')).toBe('linear');
      const slope = Number(fn.getAttribute('slope'));
      const intercept = Number(fn.getAttribute('intercept'));
      // The ramp crosses from black to white at the threshold.
      const t = thresh / 100_000;
      expect(slope * t + intercept).toBeCloseTo(1, 6);
      expect(slope * (t - 1 / 255) + intercept).toBeLessThan(0);
    }
  });

  test('a picture outline is drawn outside the picture', () => {
    const svg = paint(
      '',
      '<a:ln w="38100"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln>'
    );
    const outline = [...svg.querySelectorAll('path')].find((p) => p.getAttribute('mask'))!;
    // Twice the line width, with the inside masked away: the visible band is outside.
    expect(Number(outline.getAttribute('stroke-width'))).toBe(76_200);
    const mask = svg.querySelector(`mask${outline.getAttribute('mask')!.slice(4, -1)}`)!;
    expect(mask.querySelector('path')!.getAttribute('fill')).toBe('#000');
  });

  test('a crop that leaves nothing visible draws uncropped, as the commands read it', () => {
    // Two concurrent crops merged edge by edge: 60 % off each side.
    const svg = paint('', '', '<a:srcRect l="60000" r="60000"/>');
    const image = svg.querySelector('image')!;
    expect(Number(image.getAttribute('width'))).toBe(1_828_800);
    expect(Number(image.getAttribute('x'))).toBe(0);
  });

  test('a dotted picture outline keeps the dot rhythm with butt caps', () => {
    // A round dot (`sysDot`, cap="rnd") is one line width long, one line width apart. The
    // wider outside stroke would make its round caps reach twice as far.
    const svg = paint(
      '',
      '<a:ln w="38100" cap="rnd"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:prstDash val="sysDot"/></a:ln>'
    );
    const outline = [...svg.querySelectorAll('path')].find((p) => p.getAttribute('mask'))!;
    expect(outline.getAttribute('stroke-linecap')).toBe('butt');
    const [dash, gap] = outline.getAttribute('stroke-dasharray')!.split(' ').map(Number);
    expect(dash).toBeCloseTo(38_100, 0);
    expect(gap).toBeCloseTo(38_100, 0);
  });

  test('a picture draws its stored pixels: the browser ignores EXIF orientation', () => {
    // A JPEG's EXIF orientation does not count; a browser applies it to an
    // SVG `<image>` unless `image-orientation: none` is set.
    const image = paint('').querySelector('image')!;
    expect(image.getAttribute('style')).toBe('image-orientation: none');
  });
});

describe('picture accessibility', () => {
  /** Three pictures: P with a description, D decorative, N with a blank one; and a shape S. */
  function layout() {
    const pic = (id: number, name: string, y: number) =>
      `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${name}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
      '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
      `<p:spPr><a:xfrm><a:off x="914400" y="${y}"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
    const r = readPptx(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            rels: [['rIdImg', R_IMAGE, '../media/image1.png']],
            shapes:
              pic(2, 'P', 0) +
              pic(3, 'D', 1_000_000) +
              pic(4, 'N', 2_000_000) +
              '<p:sp><p:nvSpPr><p:cNvPr id="5" name="S"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
              '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
              '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>',
          },
        ],
        media: [
          {
            name: 'image1.png',
            bytes: new Uint8Array([137, 80, 78, 71]),
            contentType: 'image/png',
          },
        ],
      })
    );
    if (!r.ok) throw new Error('open failed');
    return layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer());
  }
  const alt: Record<string, { description: string; decorative: boolean }> = {
    P: { description: 'A dusk sky over a lake', decorative: false },
    D: { description: 'Ignored', decorative: true },
    N: { description: '  ', decorative: false },
    S: { description: 'A box', decorative: false },
  };

  test('a described picture is an img named by its description; the others are hidden', () => {
    const lay = layout();
    const ids = new Map(lay.items.map((i) => [i.element.name, i.element.nodeId]));
    const names = new Map(lay.items.map((i) => [i.element.nodeId, i.element.name]));
    const svg = paintSlide(doc, lay, {
      images: () => 'blob:test',
      idPrefix: 't',
      altText: (_part, id) => alt[names.get(id)!],
    });
    const g = (name: string) => svg.querySelector(`[data-node-id="${ids.get(name)}"]`)!;
    expect(g('P').getAttribute('role')).toBe('img');
    expect(g('P').getAttribute('aria-label')).toBe('A dusk sky over a lake');
    expect(g('P').getAttribute('aria-hidden')).toBeNull();
    for (const name of ['D', 'N']) {
      expect(g(name).getAttribute('aria-hidden')).toBe('true');
      expect(g(name).getAttribute('role')).toBeNull();
      expect(g(name).getAttribute('aria-label')).toBeNull();
    }
    // A shape keeps its content (its text) in the tree: no image role, not hidden.
    expect(g('S').getAttribute('role')).toBeNull();
    expect(g('S').getAttribute('aria-hidden')).toBeNull();
    // The group names the whole picture; the image element carries no ARIA.
    expect(svg.querySelector('image[role], image[aria-label]')).toBeNull();
  });

  test('without the alt text port the paint has no ARIA, and the drawing is the same', () => {
    const lay = layout();
    const plain = paintSlide(doc, lay, { images: () => 'blob:test', idPrefix: 't' });
    expect(plain.querySelector('[role], [aria-label], [aria-hidden]')).toBeNull();
    const named = paintSlide(doc, lay, {
      images: () => 'blob:test',
      idPrefix: 't',
      altText: () => ({ description: 'x', decorative: false }),
    });
    expect(named.querySelectorAll('[role="img"]').length).toBe(3);
    for (const el of named.querySelectorAll('[role]')) {
      el.removeAttribute('role');
      el.removeAttribute('aria-label');
    }
    expect(named.outerHTML).toBe(plain.outerHTML);
  });
});
