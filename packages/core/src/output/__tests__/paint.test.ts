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
    const bars = [...text.querySelectorAll('rect')];
    // Two underline bars and two strike bars for the faded run, two bars for the solid one.
    expect(bars.length).toBe(6);
    expect(bars.slice(0, 4).map((b) => b.getAttribute('fill-opacity'))).toEqual([
      '0.4',
      '0.4',
      '0.4',
      '0.4',
    ]);
    expect(bars.slice(4).map((b) => b.getAttribute('fill-opacity'))).toEqual([null, null]);
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
    const bars = [...text.querySelectorAll('rect')];
    const fills = bars.map((b) => [b.getAttribute('fill'), b.getAttribute('fill-opacity')]);
    // Two underline bars in the underline color with its alpha, then the strike bar in the
    // text color; the second run's underline follows its text.
    expect(fills.slice(0, 2)).toEqual([
      [fills[0][0], '0.5'],
      [fills[0][0], '0.5'],
    ]);
    expect(fills[0][0]).not.toBe(glyphs[0].getAttribute('fill'));
    expect(fills[2]).toEqual([glyphs[0].getAttribute('fill'), null]);
    expect(fills[3]).toEqual([glyphs[1].getAttribute('fill'), null]);
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
});

const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

function pngBytes(): Uint8Array<ArrayBuffer> {
  const image = new PNG({ width: 4, height: 4 });
  image.data.fill(200);
  return new Uint8Array(PNG.sync.write(image));
}
