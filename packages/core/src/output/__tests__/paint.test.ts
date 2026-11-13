import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { PNG } from 'pngjs';
import {
  buildPptx,
  cell,
  para,
  run,
  shape,
  table,
} from '../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '../../store/index.ts';
import { resolveSlide } from '../../cascade/index.ts';
import { createFixedMeasurer, layoutSlide } from '../../layout/index.ts';
import { paintSlide } from '../index.ts';
import { uprightPieces } from '../paint-text.ts';

// A DOM for this file only: the engine test run registers no global DOM.
const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

function paint(shapes: string): SVGSVGElement {
  const r = readPptx(buildPptx({ slides: [{ shapes, layout: 'blank' }] }));
  if (!r.ok) throw new Error('open failed');
  const layout = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer());
  return paintSlide(doc, layout, { images: () => undefined, idPrefix: 't' });
}

const outlined = (prst: string): string =>
  shape({
    id: 2,
    x: 1_000_000,
    y: 1_000_000,
    cx: 2_000_000,
    cy: 1_000_000,
    prst,
    fill: '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>',
    line:
      '<a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>' +
      '<a:headEnd type="oval"/><a:tailEnd type="triangle"/></a:ln>',
  });

/** The drawn line ends of a slide: the path data of each. */
function lineEnds(root: SVGSVGElement): string[] {
  return [...root.querySelectorAll('[data-line-end]')].map((p) => p.getAttribute('d') ?? '');
}

describe('line ends only on open paths (R3)', () => {
  test('a rectangle with a tail end draws no arrowhead', () => {
    expect(lineEnds(paint(outlined('rect')))).toEqual([]);
  });

  test('a callout puts line ends on its leader line, never on its box', () => {
    // An oval head (four curves) and a triangle tail (three corners).
    const ends = lineEnds(paint(outlined('borderCallout1')));
    expect(ends.length).toBe(2);
    expect(ends[0].match(/C/g)?.length).toBe(4);
    expect(ends[1].match(/L/g)?.length).toBe(2);
  });

  test('an open line keeps both ends, and its stroke stops inside the triangle', () => {
    const root = paint(outlined('line'));
    expect(lineEnds(root).length).toBe(2);
    // A 1.5 pt line takes the heads of a 2 pt line: the med triangle is 3 × 25 400 EMU long,
    // and the line stops half a unit inside its base, 63 500 EMU before the end point.
    const stroke = [...root.querySelectorAll('path')].find((p) => p.getAttribute('stroke-width'));
    const [x0, y0, x1, y1] = (stroke?.getAttribute('d') ?? '').match(/[\d.]+/g)!.map(Number);
    expect([x0, y0]).toEqual([0, 0]);
    expect(Math.hypot(2_000_000 - x1, 1_000_000 - y1)).toBeCloseTo(63_500, 0);
  });
});

describe('text decoration alpha (R8)', () => {
  test('every underline and strikethrough bar takes the text alpha', () => {
    const root = paint(
      shape({
        id: 2,
        x: 0,
        y: 0,
        cx: 4_000_000,
        cy: 1_000_000,
        text:
          para(
            run(
              'Faded',
              'u="dbl" strike="dblStrike"',
              '<a:solidFill><a:srgbClr val="C00000"><a:alpha val="40000"/></a:srgbClr></a:solidFill>'
            )
          ) + para(run('Solid', 'u="dbl"')),
      })
    );
    const text = root.querySelector('[data-pptx-text]')!;
    const glyphs = [...text.querySelectorAll('text')];
    expect(glyphs[0].getAttribute('fill-opacity')).toBe('0.4');
    // One underline path of two bars and two strike bars for the faded run, one underline path
    // for the solid one.
    const unders = [...text.querySelectorAll('path')];
    const strikes = [...text.querySelectorAll('rect')];
    expect(unders.map((u) => (u.getAttribute('d') ?? '').split('Z').length - 1)).toEqual([2, 2]);
    expect(strikes.length).toBe(2);
    expect([...unders, ...strikes].map((b) => b.getAttribute('fill-opacity'))).toEqual([
      '0.4',
      null,
      '0.4',
      '0.4',
    ]);
  });
});

describe('underline color (EG_TextUnderlineFill)', () => {
  test('a:uFill paints the underline bars; strikethrough and glyphs keep the text fill', () => {
    const root = paint(
      shape({
        id: 2,
        x: 0,
        y: 0,
        cx: 4_000_000,
        cy: 1_000_000,
        text:
          para(
            run(
              'Red line',
              'u="dbl" strike="sngStrike"',
              '<a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>' +
                '<a:uFill><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:uFill>'
            )
          ) + para(run('Text line', 'u="sng"', '<a:uFillTx/>')),
      })
    );
    const text = root.querySelector('[data-pptx-text]')!;
    const glyphs = [...text.querySelectorAll('text')];
    const fillOf = (b: Element): (string | null)[] => [
      b.getAttribute('fill'),
      b.getAttribute('fill-opacity'),
    ];
    const unders = [...text.querySelectorAll('path')].map(fillOf);
    const strikes = [...text.querySelectorAll('rect')].map(fillOf);
    // The double underline in the underline color with its alpha, then the strike bar in the
    // text color; the second run's underline follows its text.
    expect(unders[0][1]).toBe('0.5');
    expect(unders[0][0]).not.toBe(glyphs[0].getAttribute('fill'));
    expect(strikes).toEqual([[glyphs[0].getAttribute('fill'), null]]);
    expect(unders[1]).toEqual([glyphs[1].getAttribute('fill'), null]);
  });
});

describe('underline outline (EG_TextUnderlineLine)', () => {
  const underlines = (children: string): Element[] => {
    const root = paint(
      shape({
        id: 2,
        x: 0,
        y: 0,
        cx: 4_000_000,
        cy: 1_000_000,
        text: para(run('Outlined', 'u="sng"', children)),
      })
    );
    return [...root.querySelectorAll('[data-pptx-text] path')];
  };

  test('a:uLn with a fill strokes the underline; its width does not change the bar', () => {
    const plain = underlines('');
    const [outlined] = underlines(
      '<a:uLn w="25400"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:uLn>'
    );
    expect(outlined.getAttribute('d')).toBe(plain[0].getAttribute('d'));
    expect(outlined.getAttribute('stroke-width')).toBe('2');
    expect(outlined.getAttribute('stroke')).toBeTruthy();
  });

  test('a:uLn without a fill, and a:uLnTx on text without an outline, draw no outline', () => {
    expect(underlines('<a:uLn w="76200"/>')[0].getAttribute('stroke')).toBeNull();
    expect(underlines('<a:uLnTx/>')[0].getAttribute('stroke')).toBeNull();
  });
});

describe('table border caps', () => {
  const border = (dash: string): string =>
    ['lnL', 'lnR', 'lnT', 'lnB']
      .map(
        (n) =>
          `<a:${n} w="25400" cap="flat"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>` +
          `${dash ? `<a:prstDash val="${dash}"/>` : ''}</a:${n}>`
      )
      .join('');
  const lines = (dash: string): Element[] => {
    const root = paint(
      table({
        id: 2,
        x: 0,
        y: 0,
        cx: 2_000_000,
        cy: 400_000,
        columns: [2_000_000],
        tblPr: '<a:tblPr/>',
        rows: `<a:tr h="400000">${cell('x', { tcPr: border(dash) })}</a:tr>`,
      })
    );
    return [...root.querySelectorAll('[data-kind="table-borders"] line')];
  };

  test('solid borders meet with square caps', () => {
    const solid = lines('');
    expect(solid.length).toBe(4);
    for (const l of solid) expect(l.getAttribute('stroke-linecap')).toBe('square');
  });

  test('dotted borders keep their flat caps, so the dots stay apart', () => {
    const dotted = lines('sysDot');
    expect(dotted.length).toBe(4);
    for (const l of dotted) {
      expect(l.getAttribute('stroke-linecap')).toBe('butt');
      expect(l.getAttribute('stroke-dasharray')).toBe('25400 25400');
    }
  });
});

describe('effects without the Pro drawing renderers', () => {
  const all = (extra = ''): string =>
    shape({
      id: 2,
      x: 1_000_000,
      y: 1_000_000,
      cx: 2_000_000,
      cy: 1_000_000,
      fill: '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>',
      effects:
        '<a:effectLst><a:glow rad="63500"><a:srgbClr val="FF0000"/></a:glow>' +
        '<a:innerShdw blurRad="63500" dist="50800" dir="5400000"><a:srgbClr val="000000"/></a:innerShdw>' +
        extra +
        '<a:reflection blurRad="6350" stA="50000" endPos="55000" dist="5000" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>' +
        '<a:softEdge rad="12700"/></a:effectLst>',
    });

  test('paint no glow, inner shadow, reflection, or soft edge', () => {
    const root = paint(all());
    expect(root.querySelectorAll('filter').length).toBe(0);
    expect(root.querySelectorAll('feMorphology, mask, linearGradient').length).toBe(0);
    expect(root.querySelector('[data-effect="reflection"]')).toBeNull();
    // The shape itself still paints.
    expect(root.querySelector('[data-kind="shape"] path')).not.toBeNull();
  });

  test('still paint an outer shadow', () => {
    const root = paint(
      all(
        '<a:outerShdw blurRad="38100" dist="38100" dir="2700000"><a:srgbClr val="000000"/></a:outerShdw>'
      )
    );
    expect(root.querySelectorAll('[data-effect="shadow"]').length).toBe(1);
    // Only the shadow's filter: no glow or body filter, and no reflection.
    expect(root.querySelectorAll('filter').length).toBe(1);
    expect(root.querySelector('[data-effect="reflection"]')).toBeNull();
  });
});

describe('an empty text body', () => {
  test('paints the same markup as a shape with no text body', () => {
    const box = (text?: string): string =>
      shape({ id: 2, x: 1_000_000, y: 1_000_000, cx: 2_000_000, cy: 1_000_000, text });
    // The body lays out one line for the caret; that line draws nothing, so paint adds nothing.
    const empty = paint(box('<a:p><a:endParaRPr lang="en-US"/></a:p>'));
    const none = paint(box());
    expect(empty.innerHTML).toBe(none.innerHTML);
    // A body with a character does add markup, so the comparison can fail.
    expect(paint(box(para(run('x')))).innerHTML).not.toBe(none.innerHTML);
  });
});

describe('text elements', () => {
  const body = (vert: string, text: string): string =>
    shape({
      id: 2,
      x: 1_000_000,
      y: 1_000_000,
      cx: 1_000_000,
      cy: 4_000_000,
      text: para(run(text)),
      bodyPr: `<a:bodyPr vert="${vert}" wrap="square"><a:noAutofit/></a:bodyPr>`,
    });

  test('stacked letters paint as one text element, each letter at its own place', () => {
    const texts = [...paint(body('wordArtVert', 'STACK ME')).querySelectorAll('text')];
    expect(texts).toHaveLength(1);
    const text = texts[0]!;
    expect(text.textContent).toBe('STACK ME');
    expect(text.getAttribute('text-anchor')).toBe('middle');
    expect(text.getAttribute('x')!.split(' ')).toHaveLength(8);
    expect(new Set(text.getAttribute('y')!.split(' ')).size).toBe(8);
    expect(text.parentElement!.getAttribute('transform')).toBe('rotate(-90)');
  });

  test('a character outside the basic plane keeps a text element of its own', () => {
    const texts = [...paint(body('wordArtVert', 'A\u{1F600}B')).querySelectorAll('text')];
    expect(texts.map((t) => t.textContent)).toEqual(['A', '\u{1F600}', 'B']);
  });

  test('the text group sets the kerning and ligature style once for its runs', () => {
    const root = paint(body('horz', 'one two'));
    const group = root.querySelector('[data-pptx-text]')!;
    expect(group.getAttribute('style')).toContain('font-kerning:none');
    // `white-space` does not inherit into a `<text>` in every browser: each one sets it.
    expect(root.querySelector('text')!.getAttribute('style')).toBe('white-space:pre');
  });
});

describe('tab glyph paint', () => {
  test('tabs separate glyph spans without changing upright or surrogate positions', () => {
    expect(uprightPieces('A\t縦𠀋\tB', 'eastAsian')).toEqual([
      { start: 0, end: 1, upright: false },
      { start: 2, end: 3, upright: true },
      { start: 3, end: 5, upright: true },
      { start: 6, end: 7, upright: false },
    ]);
    expect(uprightPieces('A\t𠀋', 'all')).toEqual([
      { start: 0, end: 1, upright: true },
      { start: 2, end: 4, upright: true },
    ]);
    expect(uprightPieces('A\t縦𠀋\tB', 'none')).toEqual([
      { start: 0, end: 1, upright: false },
      { start: 2, end: 5, upright: false },
      { start: 6, end: 7, upright: false },
    ]);
  });
  test('keeps text positions and decorations without sending tabs to SVG text layout', () => {
    const opened = readPptx(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: shape({
              id: 2,
              x: 0,
              y: 0,
              cx: 6_000_000,
              cy: 4_000_000,
              text: para(run('A\t\tB', 'sz="1800" u="sng"')),
            }),
          },
        ],
      })
    );
    if (!opened.ok) throw new Error('open failed');
    const layout = layoutSlide(resolveSlide(opened.pkg, 0)!, createFixedMeasurer());
    const item = layout.items[0];
    if (item.kind !== 'shape' || !item.text) throw new Error('no text');
    const source = item.text.lines[0].runs[0];
    const root = paintSlide(doc, layout, { images: () => undefined, idPrefix: 'tabs' });
    const text = [...root.querySelectorAll('text')];
    expect(text.map((t) => t.textContent)).toEqual(['A', 'B']);
    expect(Number(text[0].getAttribute('x'))).toBeCloseTo(source.x[0] / 12_700, 3);
    expect(Number(text[1].getAttribute('x'))).toBeCloseTo(source.x[3] / 12_700, 3);
    // The underline is a path of one rectangle: M x0 y L x1 y L x1 y1 L x0 y1 Z.
    const d = root.querySelector('[data-pptx-text] path')?.getAttribute('d') ?? '';
    const xs = [...d.matchAll(/[ML]([-\d.]+) /g)].map((m) => Number(m[1]));
    expect(Math.max(...xs) - Math.min(...xs)).toBeCloseTo((source.x[4] - source.x[0]) / 12_700, 3);
  });

  test('a long tab-only run emits no SVG text node', () => {
    const root = paint(
      shape({
        id: 2,
        x: 0,
        y: 0,
        cx: 6_000_000,
        cy: 4_000_000,
        text: para(run('\t'.repeat(60_000), 'sz="1800"')),
      })
    );
    expect(root.querySelectorAll('text').length).toBe(0);
  });
});

describe('warped text', () => {
  test('a picture bullet draws on the parallelogram that the warp gives it', () => {
    const pPr =
      '<a:pPr marL="342900" indent="-342900"><a:buBlip><a:blip r:embed="rIdBu"/></a:buBlip></a:pPr>';
    const r = readPptx(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: shape({
              id: 2,
              x: 914_400,
              y: 914_400,
              cx: 2_743_200,
              cy: 1_828_800,
              bodyPr:
                '<a:bodyPr wrap="none"><a:prstTxWarp prst="textWave1"><a:avLst/></a:prstTxWarp></a:bodyPr>',
              text: para(run('Item'), pPr),
            }),
            rels: [['rIdBu', R_IMAGE, '../media/image1.png']],
          },
        ],
        media: [{ name: 'image1.png', bytes: pngBytes(), contentType: 'image/png' }],
      })
    );
    if (!r.ok) throw new Error('open failed');
    const layout = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer());
    const url = 'data:image/png;base64,';
    const root = paintSlide(doc, layout, { images: () => url, idPrefix: 't' });
    const image = root.querySelector('image[data-run="bullet"]');
    expect(image?.getAttribute('transform')).toMatch(/^matrix\(/);
    expect(image?.getAttribute('width')).toBe('1');
  });

  test('paints every glyph outline when a view culls the lines of other text', () => {
    const r = readPptx(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: shape({
              id: 2,
              x: 914_400,
              y: 914_400,
              cx: 2_743_200,
              cy: 1_828_800,
              bodyPr:
                '<a:bodyPr wrap="none"><a:prstTxWarp prst="textWave1"><a:avLst/></a:prstTxWarp></a:bodyPr>',
              text: para(run('Wave')),
            }),
          },
        ],
      })
    );
    if (!r.ok) throw new Error('open failed');
    const layout = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer());
    const opts = { images: () => undefined, idPrefix: 't' };
    const all = paintSlide(doc, layout, opts).querySelectorAll('[data-pptx-text] path').length;
    expect(all).toBeGreaterThan(0);
    // Far below the warped lines' boxes: an unwarped line there would not paint.
    const view = { x: 0, y: 6_000_000, width: 100, height: 100 };
    const culled = paintSlide(doc, layout, { ...opts, view });
    expect(culled.querySelectorAll('[data-pptx-text] path').length).toBe(all);
  });
});

const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

function pngBytes(): Uint8Array<ArrayBuffer> {
  const image = new PNG({ width: 4, height: 4 });
  image.data.fill(200);
  return new Uint8Array(PNG.sync.write(image));
}
