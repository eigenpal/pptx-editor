// The `line-ends-dashes` fixture: every line end type (`ST_LineEndType`, ECMA-376 Part 1,
// 20.1.10.33) at every width and length (`ST_LineEndWidth`, `ST_LineEndLength`) on both ends,
// at three line widths; every preset dash (`ST_PresetLineDashVal`, 20.1.10.48) and two custom
// dashes (20.1.8.21 custDash) at three line widths with each line cap (20.1.10.31); and line ends
// and dashes on diagonal, bent, and curved connectors and on closed shapes.

import { shape, solid, type FixtureSlide } from './builder.ts';

const IN = 914_400;
const n = (v: number): number => Math.round(v);

const END_TYPES = ['triangle', 'stealth', 'diamond', 'oval', 'arrow'] as const;
const SIZES = ['sm', 'med', 'lg'] as const;

/** The preset dashes, in schema order, and two custom dashes. */
const DASHES = [
  'solid', 'dot', 'dash', 'lgDash', 'dashDot', 'lgDashDot', 'lgDashDotDot',
  'sysDash', 'sysDot', 'sysDashDot', 'sysDashDotDot', 'cust1', 'cust2',
] as const; // prettier-ignore

/** Custom dashes: dash and space lengths in thousandths of a percent of the line width. */
const CUSTOM: Record<string, string> = {
  cust1: '<a:custDash><a:ds d="600000" sp="200000"/></a:custDash>',
  cust2: '<a:custDash><a:ds d="300000" sp="150000"/><a:ds d="100000" sp="150000"/></a:custDash>',
};

const dashXml = (dash: string): string => CUSTOM[dash] ?? `<a:prstDash val="${dash}"/>`;

const end = (tag: 'headEnd' | 'tailEnd', type: string, w: string, len: string): string =>
  type === 'none' ? '' : `<a:${tag} type="${type}" w="${w}" len="${len}"/>`;

/** One connector from (x, y) by (cx, cy), with flips for a direction up or left. */
function connector(o: {
  id: number;
  x: number;
  y: number;
  cx: number;
  cy: number;
  prst?: string;
  ln: string;
  flipH?: boolean;
  flipV?: boolean;
  rot?: number;
}): string {
  const attrs = `${o.rot ? ` rot="${n(o.rot)}"` : ''}${o.flipH ? ' flipH="1"' : ''}${o.flipV ? ' flipV="1"' : ''}`;
  return (
    `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${o.id}" name="Connector ${o.id}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>` +
    `<p:spPr><a:xfrm${attrs}><a:off x="${n(o.x)}" y="${n(o.y)}"/><a:ext cx="${n(o.cx)}" cy="${n(o.cy)}"/></a:xfrm>` +
    `<a:prstGeom prst="${o.prst ?? 'straightConnector1'}"><a:avLst/></a:prstGeom>${o.ln}</p:spPr></p:cxnSp>`
  );
}

const line = (w: number, color: string, inner: string, cap = 'flat'): string =>
  `<a:ln w="${w}" cap="${cap}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill>${inner}</a:ln>`;

/**
 * Five columns, one per line end type, and nine rows, one per width and length. Each connector
 * has the same end at its head and its tail.
 */
function endsSlide(width: number): FixtureSlide {
  const shapes: string[] = [];
  let id = 2;
  END_TYPES.forEach((type, col) => {
    SIZES.forEach((w, wi) =>
      SIZES.forEach((len, li) => {
        const row = wi * 3 + li;
        shapes.push(
          connector({
            id: id++,
            x: (0.45 + col * 2.6) * IN,
            y: (0.45 + row * 0.76) * IN,
            cx: 2 * IN,
            cy: 0,
            ln: line(width, '1F4E79', `<a:round/>${end('headEnd', type, w, len)}${end('tailEnd', type, w, len)}`),
          })
        );
      })
    );
  });
  return { layout: 'blank', shapes: shapes.join('') };
}

/** Three columns, one per line width, and one row per dash, all with the cap `cap`. */
function dashesSlide(cap: 'flat' | 'rnd' | 'sq'): FixtureSlide {
  const shapes: string[] = [];
  let id = 2;
  [9525, 28575, 57150].forEach((width, col) =>
    DASHES.forEach((dash, row) =>
      shapes.push(
        connector({
          id: id++,
          x: (0.5 + col * 4.25) * IN,
          y: (0.4 + row * 0.53) * IN,
          cx: 3.8 * IN,
          cy: 0,
          ln: line(width, 'C00000', `${dashXml(dash)}<a:round/>`, cap),
        })
      )
    )
  );
  return { layout: 'blank', shapes: shapes.join('') };
}

/** Line ends and dashes on diagonal, bent, and curved connectors, and dashes on closed shapes. */
function mixedSlide(): FixtureSlide {
  const shapes: string[] = [];
  let id = 2;
  // Diagonals in four directions, with a dashed line under each pair of ends.
  [
    { flipH: false, flipV: false, head: 'oval', tail: 'triangle', dash: 'dash' },
    { flipH: true, flipV: false, head: 'diamond', tail: 'stealth', dash: 'sysDot' },
    { flipH: false, flipV: true, head: 'arrow', tail: 'arrow', dash: 'lgDashDot' },
    { flipH: true, flipV: true, head: 'triangle', tail: 'oval', dash: 'solid' },
  ].forEach((o, i) =>
    shapes.push(
      connector({
        id: id++,
        x: (0.5 + i * 1.6) * IN,
        y: 0.5 * IN,
        cx: 1.2 * IN,
        cy: 1.6 * IN,
        flipH: o.flipH,
        flipV: o.flipV,
        ln: line(28575, '385723', `${dashXml(o.dash)}<a:round/>${end('headEnd', o.head, 'med', 'med')}${end('tailEnd', o.tail, 'lg', 'lg')}`),
      })
    )
  );
  // A rotated straight connector.
  shapes.push(
    connector({
      id: id++,
      x: 7 * IN,
      y: 1.3 * IN,
      cx: 2.5 * IN,
      cy: 0,
      rot: 30 * 60000,
      ln: line(19050, '7030A0', `<a:prstDash val="dashDot"/><a:round/>${end('headEnd', 'stealth', 'lg', 'sm')}${end('tailEnd', 'triangle', 'sm', 'lg')}`),
    })
  );
  // Bent and curved connectors, with ends that follow the last segment's direction.
  shapes.push(
    connector({
      id: id++,
      x: 10 * IN,
      y: 0.5 * IN,
      cx: 2.6 * IN,
      cy: 1.8 * IN,
      prst: 'bentConnector3',
      ln: line(28575, 'C00000', `<a:prstDash val="sysDash"/><a:miter lim="800000"/>${end('headEnd', 'diamond', 'lg', 'lg')}${end('tailEnd', 'triangle', 'med', 'med')}`),
    }),
    connector({
      id: id++,
      x: 0.5 * IN,
      y: 3 * IN,
      cx: 3.2 * IN,
      cy: 1.6 * IN,
      prst: 'curvedConnector3',
      ln: line(28575, '1F4E79', `<a:prstDash val="lgDash"/><a:round/>${end('headEnd', 'oval', 'sm', 'sm')}${end('tailEnd', 'stealth', 'lg', 'lg')}`),
    }),
    connector({
      id: id++,
      x: 4.2 * IN,
      y: 3 * IN,
      cx: 3 * IN,
      cy: 1.6 * IN,
      prst: 'curvedConnector3',
      flipV: true,
      ln: line(38100, 'ED7D31', `<a:prstDash val="sysDashDotDot"/><a:round/>${end('headEnd', 'arrow', 'lg', 'med')}${end('tailEnd', 'arrow', 'sm', 'sm')}`, 'rnd'),
    })
  );
  // Dashes around closed shapes, which get no line ends.
  shapes.push(
    shape({
      id: id++,
      x: 7.8 * IN,
      y: 2.9 * IN,
      cx: 2.2 * IN,
      cy: 1.6 * IN,
      fill: solid('DEEBF7'),
      line: line(38100, '1F4E79', '<a:prstDash val="dash"/><a:miter lim="800000"/><a:headEnd type="triangle"/><a:tailEnd type="triangle"/>'),
    }),
    shape({
      id: id++,
      x: 10.4 * IN,
      y: 2.9 * IN,
      cx: 2.2 * IN,
      cy: 1.6 * IN,
      prst: 'ellipse',
      fill: '<a:noFill/>',
      line: line(28575, '7030A0', '<a:prstDash val="sysDot"/><a:round/>', 'rnd'),
    }),
    shape({
      id: id++,
      x: 0.6 * IN,
      y: 5.1 * IN,
      cx: 3 * IN,
      cy: 1.8 * IN,
      prst: 'roundRect',
      fill: solid('FFF2CC'),
      line: line(19050, 'C00000', '<a:prstDash val="lgDashDotDot"/><a:round/>'),
    }),
    // An open arc gets its line ends along the curve.
    shape({
      id: id++,
      x: 4.4 * IN,
      y: 5 * IN,
      cx: 2.6 * IN,
      cy: 2 * IN,
      prst: 'arc',
      fill: '<a:noFill/>',
      line: line(28575, '385723', `<a:prstDash val="dash"/><a:round/>${end('headEnd', 'stealth', 'med', 'med')}${end('tailEnd', 'triangle', 'lg', 'med')}`),
    }),
    // A line shape (not a connector) with ends, and a hairline-width connector with ends.
    shape({
      id: id++,
      x: 7.8 * IN,
      y: 5.3 * IN,
      cx: 4.6 * IN,
      cy: 0,
      prst: 'line',
      line: line(63500, '404040', `<a:prstDash val="sysDashDot"/><a:round/>${end('headEnd', 'oval', 'med', 'med')}${end('tailEnd', 'diamond', 'med', 'med')}`),
    }),
    connector({
      id: id++,
      x: 7.8 * IN,
      y: 6.4 * IN,
      cx: 4.6 * IN,
      cy: 0,
      ln: line(3175, '000000', `<a:round/>${end('headEnd', 'triangle', 'lg', 'lg')}${end('tailEnd', 'arrow', 'lg', 'lg')}`),
    })
  );
  return { layout: 'blank', shapes: shapes.join('') };
}

export const lineEndsDashesFixtures: Record<string, FixtureSlide[]> = {
  'line-ends-dashes': [
    endsSlide(9525),
    endsSlide(28575),
    endsSlide(57150),
    dashesSlide('flat'),
    dashesSlide('rnd'),
    dashesSlide('sq'),
    mixedSlide(),
  ],
};
