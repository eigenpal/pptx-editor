import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
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

/** Paths that carry a line-end marker, and whether each is closed. */
function markedPaths(root: SVGSVGElement): { closed: boolean; start: boolean; end: boolean }[] {
  return [...root.querySelectorAll('path')]
    .filter((p) => p.getAttribute('marker-start') || p.getAttribute('marker-end'))
    .map((p) => ({
      closed: /Z/i.test(p.getAttribute('d') ?? ''),
      start: p.hasAttribute('marker-start'),
      end: p.hasAttribute('marker-end'),
    }));
}

describe('line ends only on open paths (R3)', () => {
  test('a rectangle with a tail end draws no arrowhead and defines no marker', () => {
    const root = paint(outlined('rect'));
    expect(markedPaths(root)).toEqual([]);
    expect(root.querySelectorAll('marker').length).toBe(0);
  });

  test('a callout puts line ends on its leader line, never on its box', () => {
    const marked = markedPaths(paint(outlined('borderCallout1')));
    expect(marked.length).toBe(1);
    expect(marked[0]).toEqual({ closed: false, start: true, end: true });
  });

  test('an open line keeps both ends', () => {
    expect(markedPaths(paint(outlined('line')))).toEqual([
      { closed: false, start: true, end: true },
    ]);
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

describe('effect filter regions', () => {
  test('a shadow and a glow keep the outline that paints past the box', () => {
    const root = paint(
      shape({
        id: 2,
        x: 1_000_000,
        y: 1_000_000,
        cx: 2_000_000,
        cy: 1_000_000,
        fill: '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>',
        line: '<a:ln w="88900"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>',
        effects:
          '<a:effectLst><a:glow rad="12700"><a:srgbClr val="FF0000"/></a:glow>' +
          '<a:outerShdw blurRad="0" dist="0"><a:srgbClr val="000000"/></a:outerShdw></a:effectLst>',
      })
    );
    const filters = [...root.querySelectorAll('filter')];
    expect(filters.length).toBe(2);
    // Half the 7 pt outline lies outside the box; neither region may cut it off.
    for (const f of filters)
      expect(Number(f.getAttribute('x'))).toBeLessThanOrEqual(1_000_000 - 88_900);
  });

  test('the region keeps a miter join and arrowheads', () => {
    const region = (line: string, prst = 'rect'): number => {
      const root = paint(
        shape({
          id: 2,
          x: 1_000_000,
          y: 1_000_000,
          cx: 2_000_000,
          cy: 1_000_000,
          prst,
          line,
          effects:
            '<a:effectLst><a:glow rad="12700"><a:srgbClr val="FF0000"/></a:glow></a:effectLst>',
        })
      );
      return Number(root.querySelector('filter')!.getAttribute('x'));
    };
    const solid = '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>';
    // A miter limit of 8 reaches up to 4 widths past the corner.
    expect(region(`<a:ln w="25400">${solid}<a:miter lim="800000"/></a:ln>`)).toBeLessThanOrEqual(
      1_000_000 - 4 * 25_400
    );
    // An arrowhead on an open path reaches 3 widths past its end.
    expect(
      region(`<a:ln w="25400">${solid}<a:round/><a:tailEnd type="triangle"/></a:ln>`, 'line')
    ).toBeLessThanOrEqual(1_000_000 - 3 * 25_400);
  });

  test('a soft edge and an inner shadow keep a callout tail outside the box', () => {
    const root = paint(
      shape({
        id: 2,
        x: 1_000_000,
        y: 1_000_000,
        cx: 2_000_000,
        cy: 1_000_000,
        prst: 'wedgeRectCallout',
        // The tail tip lies 1.5 box heights below the box.
        avLst: '<a:gd name="adj1" fmla="val 0"/><a:gd name="adj2" fmla="val 200000"/>',
        fill: '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>',
        effects:
          '<a:effectLst><a:innerShdw blurRad="63500" dist="50800" dir="5400000">' +
          '<a:srgbClr val="000000"/></a:innerShdw><a:softEdge rad="12700"/></a:effectLst>',
      })
    );
    const f = root.querySelector('filter')!;
    const bottom = Number(f.getAttribute('y')) + Number(f.getAttribute('height'));
    expect(bottom).toBeGreaterThanOrEqual(1_000_000 + 2.5 * 1_000_000);
  });

  test('a turned callout keeps its tail in the region too', () => {
    const root = paint(
      shape({
        id: 2,
        x: 1_000_000,
        y: 1_000_000,
        cx: 2_000_000,
        cy: 1_000_000,
        rot: 5_400_000,
        prst: 'wedgeRectCallout',
        avLst: '<a:gd name="adj1" fmla="val 0"/><a:gd name="adj2" fmla="val 200000"/>',
        fill: '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>',
        effects: '<a:effectLst><a:softEdge rad="12700"/></a:effectLst>',
      })
    );
    // Turned 90° about the box center (2 000 000, 1 500 000), the tail tip moves to x = 0.
    const f = root.querySelector('filter')!;
    expect(Number(f.getAttribute('x'))).toBeLessThanOrEqual(0);
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
